"""
VIA ネットネット株スクリーナー（グレアム流）
日本株を対象に NCAV戦略でスクリーニングする独立ツール。

判定基準:
  NCAV = 流動資産 - 負債総額
  時価総額 <= NCAV x (2/3)   ← グレアムの安全域基準
  営業CF(直近期) >= 0         ← NCAVの維持可能性チェック

追加チェック項目:
  ① 過去3期営業CFの一貫性（全期間黒字か）
  ② バリュートラップ・リスク兆候（配当・自社株買い・売上トレンド）
  ③ 中小企業リスク（Claude APIによる定性分析、オプション）

使い方:
  python via_netnet.py          全銘柄スクリーニング
  python via_netnet.py --test   先頭20銘柄のみテスト実行

所要時間: 全銘柄(約4000社)で約60〜90分（yfinance呼び出しのため）
"""

import sys, os, time, warnings
import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime

warnings.filterwarnings('ignore')

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR   = os.path.join(SCRIPT_DIR, "via_cache")
OUTPUT_CSV  = os.path.join(SCRIPT_DIR, "via_netnet_results.csv")
OUTPUT_HTML = os.path.join(SCRIPT_DIR, "via_netnet_results.html")

SLEEP_SEC        = 0.4
NCAV_MARGIN      = 2/3   # グレアムの安全域基準（時価総額がNCAVの2/3以下）
FINS_CACHE_DAYS  = 30


# ── 銘柄リスト取得 ──

def load_jp_tickers():
    """jpx_tickers.xls から東証銘柄リストを読み込む"""
    for ext in ["xlsx", "xls"]:
        manual = os.path.join(SCRIPT_DIR, f"jpx_tickers.{ext}")
        if os.path.exists(manual):
            try:
                engine = "openpyxl" if ext == "xlsx" else "xlrd"
                df = pd.read_excel(manual, dtype=str, engine=engine)
                code_col = None
                for col in df.columns:
                    if any(k in str(col) for k in ["コード", "code", "Code"]):
                        code_col = col
                        break
                if code_col is None:
                    continue
                codes = df[code_col].dropna().astype(str).str.strip()
                codes_4 = codes[codes.str.match(r'^\d{4}$')]
                codes_5 = codes[codes.str.match(r'^\d{5}$')]
                all_codes = pd.concat([codes_4, codes_5]).drop_duplicates().tolist()
                tickers = [c + ".T" for c in all_codes]
                print(f"  JPXファイルから {len(tickers)} 銘柄取得")
                return tickers
            except Exception as e:
                print(f"  JPXファイル読込エラー: {e}")
    print("  JPXファイルが見つかりません。jpx_tickers.xls を配置してください。")
    return []


# ── 財務データ取得・NCAV計算 ──

def _cache_path(ticker):
    os.makedirs(CACHE_DIR, exist_ok=True)
    safe = ticker.replace(".", "_")
    return os.path.join(CACHE_DIR, f"netnet_{safe}.json")

def _load_cache(ticker):
    import json
    path = _cache_path(ticker)
    if not os.path.exists(path):
        return None
    age_days = (time.time() - os.path.getmtime(path)) / 86400
    if age_days > FINS_CACHE_DAYS:
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

def _save_cache(ticker, data):
    import json
    path = _cache_path(ticker)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)


def get_netnet_data(ticker):
    """
    yfinanceから流動資産・負債総額・営業CF・株価・市場規模を取得し、
    NCAVと判定結果、追加リスクチェック結果を返す。
    """
    cached = _load_cache(ticker)
    if cached:
        return cached

    try:
        tk = yf.Ticker(ticker)
        bs = tk.balance_sheet
        cf = tk.cash_flow
        fi = tk.income_stmt

        if bs is None or bs.empty:
            return None

        def _latest(df, keys, idx=0):
            for k in keys:
                if k in df.index:
                    vals = df.loc[k].dropna()
                    if len(vals) > idx:
                        return float(vals.iloc[idx])
            return None

        def _series(df, keys):
            for k in keys:
                if k in df.index:
                    return df.loc[k].dropna().tolist()
            return []

        current_assets = _latest(bs, ["Current Assets"])
        total_liab      = _latest(bs, ["Total Liabilities Net Minority Interest",
                                        "Total Liab"])

        ocf_series = _series(cf, ["Operating Cash Flow",
                                   "Cash Flow From Continuing Operating Activities"])
        ocf = ocf_series[0] if ocf_series else None

        if current_assets is None or total_liab is None:
            return None

        ncav = current_assets - total_liab

        # 株価・市場規模
        price = None
        market_cap = None
        shares = None
        name = ticker
        sector = ""
        try:
            price = tk.fast_info.last_price
            shares = tk.fast_info.shares
            info = tk.info
            name = info.get("longName") or info.get("shortName") or ticker
            sector = info.get("sector", "")

            shares_info = info.get("sharesOutstanding")
            if shares_info and (not shares or shares < 1000):
                shares = shares_info

            if price and shares and shares >= 1000:
                market_cap = price * shares
            else:
                market_cap = tk.fast_info.market_cap
        except Exception:
            pass

        if not price or not market_cap or not shares or shares < 1000:
            return None

        ncav_per_share = ncav / shares if shares else None

        # 判定
        ncav_positive = ncav > 0
        below_2_3 = (market_cap <= ncav * NCAV_MARGIN) if ncav_positive else False
        ocf_positive = (ocf is not None and ocf >= 0)
        is_netnet = ncav_positive and below_2_3

        # ── ① 過去3期営業CFの一貫性チェック ──
        ocf_3y = ocf_series[:3]  # 直近3期（新しい順）
        ocf_3y_all_positive = (
            len(ocf_3y) >= 3 and all(v is not None and v >= 0 for v in ocf_3y)
        )
        ocf_3y_count_positive = sum(1 for v in ocf_3y if v is not None and v >= 0)

        # ── ③ バリュートラップ・リスク兆候（機械判定できる部分） ──
        # 配当履歴: 直近1年以内に配当実績があるか
        has_dividend = False
        try:
            div = tk.dividends
            if div is not None and not div.empty:
                last_div_date = div.index[-1]
                # タイムゾーン考慮せず年数だけ比較
                years_since = (pd.Timestamp.now(tz=last_div_date.tz) - last_div_date).days / 365
                has_dividend = years_since <= 2
        except Exception:
            pass

        # 自社株買い: 直近期に実施したか
        buyback_series = _series(cf, ["Repurchase Of Capital Stock",
                                       "Common Stock Payments"])
        has_buyback = bool(buyback_series and buyback_series[0] and buyback_series[0] < 0)

        # 売上高トレンド: 直近3期で減少傾向か（先細りシグナル）
        rev_series = _series(fi, ["Total Revenue"])
        rev_3y = rev_series[:3]
        revenue_declining = (
            len(rev_3y) >= 3 and rev_3y[0] < rev_3y[1] < rev_3y[2]
        ) if len(rev_3y) >= 3 else None  # 古い→新しい順なので逆転に注意

        # 実際は新しい順で入っているため declining は rev_3y[0] < rev_3y[2] のロジックを使う
        revenue_declining = (
            len(rev_3y) >= 3 and rev_3y[0] < rev_3y[2]
        ) if len(rev_3y) >= 3 else None

        # バリュートラップ・スコア（資本還元シグナルが何もない場合に高リスク）
        # 配当なし AND 自社株買いなし → 株主還元意識が低い可能性
        value_trap_flag = (not has_dividend) and (not has_buyback)

        result = {
            "ticker": ticker,
            "name": name,
            "sector": sector,
            "price": round(price, 2),
            "market_cap": round(market_cap, 0),
            "current_assets": round(current_assets, 0),
            "total_liabilities": round(total_liab, 0),
            "ncav": round(ncav, 0),
            "ncav_per_share": round(ncav_per_share, 2) if ncav_per_share else None,
            "ocf": round(ocf, 0) if ocf is not None else None,
            "ocf_positive": ocf_positive,
            "ncav_ratio": round(market_cap / ncav, 3) if ncav_positive else None,
            "is_netnet": is_netnet,
            "margin_pct": round((ncav * NCAV_MARGIN - market_cap) / (ncav * NCAV_MARGIN) * 100, 1)
                          if ncav_positive and ncav * NCAV_MARGIN > 0 else None,
            # ① 営業CF一貫性
            "ocf_3y_all_positive": ocf_3y_all_positive,
            "ocf_3y_count_positive": ocf_3y_count_positive,
            "ocf_3y_values": [round(v, 0) if v is not None else None for v in ocf_3y],
            # ③ バリュートラップ兆候
            "has_dividend": has_dividend,
            "has_buyback": has_buyback,
            "revenue_declining": revenue_declining,
            "value_trap_flag": value_trap_flag,
        }
        _save_cache(ticker, result)
        return result

    except Exception:
        return None


# ── HTML生成 ──

def make_html(rows, total_input, generated):
    netnet = [r for r in rows if r.get("is_netnet")]
    netnet_ocf = [r for r in netnet if r.get("ocf_positive")]
    netnet_3y  = [r for r in netnet if r.get("ocf_3y_all_positive")]
    netnet_safe = [r for r in netnet if r.get("ocf_3y_all_positive") and not r.get("value_trap_flag")]

    def fmt_money(v):
        if v is None: return "—"
        return f"{v:,.0f}"

    trs = []
    for r in sorted(rows, key=lambda x: (x.get("ncav_ratio") if x.get("ncav_ratio") is not None else 999)):
        if not r.get("is_netnet"):
            continue

        ocf_badge = ('<span class="ocf-ok">CF+</span>' if r.get("ocf_positive")
                     else '<span class="ocf-bad">CF-</span>')
        ocf_sort = "1" if r.get("ocf_positive") else "0"

        # ① 過去3期CF一貫性バッジ
        n3 = r.get("ocf_3y_count_positive", 0)
        if r.get("ocf_3y_all_positive"):
            cf3y_badge = '<span class="cf3-ok">3期連続◎</span>'
            cf3y_sort = "3"
        elif n3 >= 2:
            cf3y_badge = f'<span class="cf3-mid">{n3}/3期</span>'
            cf3y_sort = "2"
        elif n3 == 1:
            cf3y_badge = f'<span class="cf3-bad">{n3}/3期</span>'
            cf3y_sort = "1"
        else:
            cf3y_badge = '<span class="cf3-bad">データ不足</span>'
            cf3y_sort = "0"

        # ③ バリュートラップ兆候バッジ
        div_mark  = '✓' if r.get("has_dividend") else '—'
        bb_mark   = '✓' if r.get("has_buyback") else '—'
        if r.get("value_trap_flag"):
            trap_badge = '<span class="trap-warn">⚠ 株主還元なし</span>'
            trap_sort = "0"
        else:
            trap_badge = '<span class="trap-ok">還元あり</span>'
            trap_sort = "1"

        rev_decl = r.get("revenue_declining")
        rev_badge = ('<span class="rev-bad">↓減少</span>' if rev_decl
                     else '<span class="rev-ok">→維持</span>' if rev_decl is False
                     else '<span class="rev-na">—</span>')

        trs.append(
            f'<tr>'
            f'<td><b>{r["ticker"]}</b></td>'
            f'<td>{str(r["name"])[:26]}</td>'
            f'<td>{str(r.get("sector",""))[:16]}</td>'
            f'<td class="num">{r["price"]}</td>'
            f'<td class="num">{fmt_money(r["market_cap"])}</td>'
            f'<td class="num">{fmt_money(r["ncav"])}</td>'
            f'<td class="num ratio">{r.get("ncav_ratio","—")}</td>'
            f'<td class="num">{r.get("margin_pct","—")}%</td>'
            f'<td data-sort="{ocf_sort}">{ocf_badge}</td>'
            f'<td data-sort="{cf3y_sort}" title="直近3期: {r.get("ocf_3y_values")}">{cf3y_badge}</td>'
            f'<td data-sort="{trap_sort}" title="配当:{div_mark} 自社株買い:{bb_mark}">{trap_badge}</td>'
            f'<td>{rev_badge}</td>'
            f'</tr>'
        )

    return f"""<!DOCTYPE html>
<html lang="ja"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>VIA ネットネット株スクリーニング</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:system-ui,sans-serif;font-size:13px;background:#f5f5f0;color:#1a1a1a;padding:1rem}}
h1{{font-size:18px;font-weight:500;margin-bottom:.4rem}}
.meta{{font-size:12px;color:#888;margin-bottom:.5rem;line-height:1.6}}
.cfg{{font-size:11px;padding:8px 12px;border-radius:6px;margin-bottom:6px;line-height:1.7;
      background:#eef3fb;border-left:3px solid #185FA5}}
.cfg2{{font-size:11px;padding:8px 12px;border-radius:6px;margin-bottom:10px;line-height:1.7;
      background:#fef3e2;border-left:3px solid #BA7517}}
.summary{{display:flex;gap:10px;margin-bottom:.8rem;flex-wrap:wrap}}
.sc{{background:#fff;border:1px solid #e0e0d8;border-radius:8px;padding:8px 14px;min-width:130px}}
.sc .sv{{font-size:22px;font-weight:600}} .sc .sl{{font-size:11px;color:#888}}
table{{border-collapse:collapse;width:100%;white-space:nowrap;background:#fff}}
th{{background:#f0f0e8;font-weight:500;padding:6px 9px;border-bottom:2px solid #ccc;font-size:12px;cursor:pointer}}
th:hover{{background:#e4e4d8}}
td{{padding:5px 9px;border-bottom:1px solid #eee;font-size:12px}}
tr:hover td{{background:#f9f9f5}}
td.num{{text-align:right;font-variant-numeric:tabular-nums}}
td.ratio{{font-weight:700;color:#0F6E56}}
.ocf-ok{{color:#0F6E56;font-weight:700}}
.ocf-bad{{color:#993C1D}}
.cf3-ok{{color:#0F6E56;font-weight:700;background:#e8f8f2;padding:2px 6px;border-radius:4px}}
.cf3-mid{{color:#854F0B;background:#fef3e2;padding:2px 6px;border-radius:4px}}
.cf3-bad{{color:#993C1D;background:#fbe9e7;padding:2px 6px;border-radius:4px}}
.trap-ok{{color:#0F6E56}}
.trap-warn{{color:#BA7517;font-weight:700}}
.rev-ok{{color:#0F6E56}}
.rev-bad{{color:#993C1D;font-weight:700}}
.rev-na{{color:#aaa}}
</style></head><body>
<h1>VIA ネットネット株スクリーニング（グレアム流） <a href="via_results.html" style="font-size:13px;font-weight:500;color:#185FA5;margin-left:12px;text-decoration:none;border:1px solid #185FA5;border-radius:6px;padding:3px 10px;vertical-align:middle">📈 通常のVIAスクリーニングはこちら</a></h1>
<p class="meta">生成: {generated} ／ 対象: {total_input}銘柄（日本株）</p>
<div class="cfg">
  <b>NCAV判定</b>: 流動資産−負債総額 を計算し、時価総額がNCAVの2/3以下の銘柄を抽出。
</div>
<div class="cfg2">
  <b>追加リスクチェック</b>:　
  <b>3期CF</b>=過去3期営業CFが全て黒字か（NCAV維持の信頼性）　／　
  <b>株主還元</b>=配当または自社株買いの実績（資本配分意識のチェック。バリュートラップ回避の参考）　／　
  <b>売上trend</b>=直近3期の売上推移（事業の先細りシグナル）
</div>
<div class="summary">
  <div class="sc"><div class="sv" style="color:#185FA5">{len(netnet)}</div><div class="sl">ネットネット株</div></div>
  <div class="sc"><div class="sv" style="color:#0F6E56">{len(netnet_ocf)}</div><div class="sl">うち営業CF黒字</div></div>
  <div class="sc"><div class="sv" style="color:#0F6E56">{len(netnet_3y)}</div><div class="sl">うち3期連続CF黒字</div></div>
  <div class="sc"><div class="sv" style="color:#185FA5">{len(netnet_safe)}</div><div class="sl">3期黒字×株主還元あり</div></div>
</div>
<table id="tbl">
<thead><tr>
  <th onclick="sortTable(0)">Ticker</th><th onclick="sortTable(1)">名称</th>
  <th onclick="sortTable(2)">セクター</th><th onclick="sortTable(3)">株価</th>
  <th onclick="sortTable(4)">時価総額</th><th onclick="sortTable(5)">NCAV</th>
  <th onclick="sortTable(6)">時価総額/NCAV</th><th onclick="sortTable(7)">安全域余地</th>
  <th onclick="sortTable(8)">営業CF</th><th onclick="sortTable(9)">3期CF</th>
  <th onclick="sortTable(10)">株主還元</th><th onclick="sortTable(11)">売上trend</th>
</tr></thead>
<tbody>{"".join(trs)}</tbody>
</table>
<script>
var sd = {{}};
function sortTable(c) {{
  var tb = document.querySelector('#tbl tbody');
  var rows = Array.from(tb.rows);
  sd[c] = !sd[c];
  rows.sort(function(a, b) {{
    var av = a.cells[c] ? (a.cells[c].getAttribute('data-sort') || a.cells[c].textContent.trim()) : '';
    var bv = b.cells[c] ? (b.cells[c].getAttribute('data-sort') || b.cells[c].textContent.trim()) : '';
    var an = parseFloat(av.replace(/[^0-9.-]/g, ''));
    var bn = parseFloat(bv.replace(/[^0-9.-]/g, ''));
    if (!isNaN(an) && !isNaN(bn)) return sd[c] ? an - bn : bn - an;
    return sd[c] ? av.localeCompare(bv) : bv.localeCompare(av);
  }});
  rows.forEach(function(r) {{ tb.appendChild(r); }});
}}
</script>
</body></html>"""


# ── メイン ──

def main():
    test_mode = "--test" in sys.argv
    force_refresh = "--refresh" in sys.argv

    print("=" * 60)
    print("VIA ネットネット株スクリーナー（グレアム流）")
    print(f"判定基準: 時価総額 <= NCAV x {NCAV_MARGIN:.3f} （流動資産-負債総額の2/3）")
    print("追加チェック: 過去3期CF一貫性 / 株主還元実績 / 売上トレンド")
    print("=" * 60)

    print("\n[銘柄リスト取得]")
    tickers = load_jp_tickers()
    if not tickers:
        print("銘柄リストが空です。終了します。")
        return

    if test_mode:
        tickers = tickers[:20]
        print(f"テストモード: 先頭{len(tickers)}銘柄のみ実行")

    if force_refresh:
        import glob
        for f in glob.glob(os.path.join(CACHE_DIR, "netnet_*.json")):
            os.remove(f)
        print("キャッシュを全削除しました（--refresh指定）")

    total = len(tickers)
    print(f"\n対象: {total} 銘柄")
    print(f"推定時間: 約 {round(total * SLEEP_SEC / 60)} 分\n")

    rows = []
    netnet_count = 0
    for i, ticker in enumerate(tickers):
        pct = (i + 1) / total * 100
        print(f"[{i+1:4d}/{total}] {ticker:<10} ({pct:5.1f}%)", end=" ", flush=True)
        r = get_netnet_data(ticker)
        if r:
            rows.append(r)
            if r.get("is_netnet"):
                netnet_count += 1
                ocf_str = "CF+" if r.get("ocf_positive") else "CF-"
                cf3_str = "3期黒字" if r.get("ocf_3y_all_positive") else f"{r.get('ocf_3y_count_positive',0)}/3期"
                trap_str = "還元あり" if not r.get("value_trap_flag") else "還元なし"
                print(f"-> NCAV比率:{r['ncav_ratio']}  {ocf_str}  {cf3_str}  {trap_str}  ★ネットネット")
            else:
                print("-> 対象外")
        else:
            print("-> データなし")
        time.sleep(SLEEP_SEC)

    df = pd.DataFrame(rows)
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"\nCSV出力: {OUTPUT_CSV}")

    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    html = make_html(rows, total, generated)
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML出力: {OUTPUT_HTML}")

    print("\n" + "=" * 60)
    print(f"ネットネット株: {netnet_count} 銘柄")
    print("=" * 60)

if __name__ == "__main__":
    main()
