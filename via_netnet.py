"""
VIA ネットネット株スクリーナー（グレアム流）
日本株を対象に NCAV戦略でスクリーニングする独立ツール。

判定基準:
  NCAV = 流動資産 - 負債総額
  時価総額 <= NCAV x (2/3)   ← グレアムの安全域基準
  営業CF(直近期) >= 0         ← NCAVの維持可能性チェック

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


# ── 銘柄リスト取得（既存via_screener.pyのロジックを再利用） ──

def load_jp_tickers():
    """jpx_tickers.xls から東証銘柄リストを読み込む（via_screener.pyと同じ仕組み）"""
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
    NCAVと判定結果を返す。
    """
    cached = _load_cache(ticker)
    if cached:
        return cached

    try:
        tk = yf.Ticker(ticker)
        bs = tk.balance_sheet
        cf = tk.cash_flow

        if bs is None or bs.empty:
            return None

        def _latest(df, keys):
            for k in keys:
                if k in df.index:
                    vals = df.loc[k].dropna()
                    if not vals.empty:
                        return float(vals.iloc[0])  # 最新年度（先頭列）
            return None

        current_assets = _latest(bs, ["Current Assets"])
        total_liab      = _latest(bs, ["Total Liabilities Net Minority Interest",
                                        "Total Liab"])
        ocf = None
        if cf is not None and not cf.empty:
            ocf = _latest(cf, ["Operating Cash Flow",
                                "Cash Flow From Continuing Operating Activities"])

        if current_assets is None or total_liab is None:
            return None

        ncav = current_assets - total_liab

        # 株価・時価総額（fast_infoから取得、異常値は info で補正）
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

            # fast_info.sharesが異常値（極端に小さい）の場合はinfoから取得し直す
            shares_info = info.get("sharesOutstanding")
            if shares_info and (not shares or shares < 1000):
                shares = shares_info

            # market_capも同様にinfoで補正・再計算
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
            "ncav_ratio": round(market_cap / ncav, 3) if ncav_positive else None,  # 1.0未満が割安
            "is_netnet": is_netnet,
            "margin_pct": round((ncav * NCAV_MARGIN - market_cap) / (ncav * NCAV_MARGIN) * 100, 1)
                          if ncav_positive and ncav * NCAV_MARGIN > 0 else None,
        }
        _save_cache(ticker, result)
        return result

    except Exception:
        return None


# ── HTML生成 ──

def make_html(rows, total_input, generated):
    netnet = [r for r in rows if r.get("is_netnet")]
    netnet_ocf = [r for r in netnet if r.get("ocf_positive")]

    def fmt_money(v):
        if v is None: return "—"
        return f"{v:,.0f}"

    trs = []
    for r in sorted(rows, key=lambda x: (x.get("ncav_ratio") if x.get("ncav_ratio") is not None else 999)):
        if not r.get("is_netnet"):
            continue
        ocf_badge = ('<span class="ocf-ok">CF+</span>' if r.get("ocf_positive")
                     else '<span class="ocf-bad">CF-</span>')
        trs.append(
            f'<tr>'
            f'<td><b>{r["ticker"]}</b></td>'
            f'<td>{str(r["name"])[:28]}</td>'
            f'<td>{str(r.get("sector",""))[:18]}</td>'
            f'<td class="num">{r["price"]}</td>'
            f'<td class="num">{fmt_money(r["market_cap"])}</td>'
            f'<td class="num">{fmt_money(r["current_assets"])}</td>'
            f'<td class="num">{fmt_money(r["total_liabilities"])}</td>'
            f'<td class="num">{fmt_money(r["ncav"])}</td>'
            f'<td class="num">{r.get("ncav_per_share","—")}</td>'
            f'<td class="num ratio">{r.get("ncav_ratio","—")}</td>'
            f'<td class="num">{r.get("margin_pct","—")}%</td>'
            f'<td>{ocf_badge}</td>'
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
.cfg{{font-size:11px;padding:8px 12px;border-radius:6px;margin-bottom:10px;line-height:1.7;
      background:#eef3fb;border-left:3px solid #185FA5}}
.summary{{display:flex;gap:10px;margin-bottom:.8rem;flex-wrap:wrap}}
.sc{{background:#fff;border:1px solid #e0e0d8;border-radius:8px;padding:8px 14px;min-width:130px}}
.sc .sv{{font-size:22px;font-weight:600}} .sc .sl{{font-size:11px;color:#888}}
table{{border-collapse:collapse;width:100%;white-space:nowrap;background:#fff}}
th{{background:#f0f0e8;font-weight:500;padding:6px 9px;border-bottom:2px solid #ccc;font-size:12px}}
td{{padding:5px 9px;border-bottom:1px solid #eee;font-size:12px}}
tr:hover td{{background:#f9f9f5}}
td.num{{text-align:right;font-variant-numeric:tabular-nums}}
td.ratio{{font-weight:700;color:#0F6E56}}
.ocf-ok{{color:#0F6E56;font-weight:700}}
.ocf-bad{{color:#993C1D}}
</style></head><body>
<h1>VIA ネットネット株スクリーニング（グレアム流） <a href="via_results.html" style="font-size:13px;font-weight:500;color:#185FA5;margin-left:12px;text-decoration:none;border:1px solid #185FA5;border-radius:6px;padding:3px 10px;vertical-align:middle">📈 通常のVIAスクリーニングはこちら</a></h1>
<p class="meta">生成: {generated} ／ 対象: {total_input}銘柄（日本株）</p>
<div class="cfg">
  <b>判定基準</b>: NCAV（流動資産−負債総額）を計算し、時価総額がNCAVの2/3以下の銘柄を抽出。
  CF+は直近期の営業キャッシュフローが黒字（NCAV維持可能性のチェック）。
  NCAV比率が低いほど割安度が高い。
</div>
<div class="summary">
  <div class="sc"><div class="sv" style="color:#185FA5">{len(netnet)}</div><div class="sl">ネットネット株</div></div>
  <div class="sc"><div class="sv" style="color:#0F6E56">{len(netnet_ocf)}</div><div class="sl">うち営業CF黒字</div></div>
</div>
<table>
<thead><tr>
  <th>Ticker</th><th>名称</th><th>セクター</th><th>株価</th>
  <th>時価総額</th><th>流動資産</th><th>負債総額</th><th>NCAV</th>
  <th>NCAV/株</th><th>時価総額/NCAV</th><th>安全域余地</th><th>営業CF</th>
</tr></thead>
<tbody>{"".join(trs)}</tbody>
</table>
</body></html>"""


# ── メイン ──

def main():
    test_mode = "--test" in sys.argv

    print("=" * 60)
    print("VIA ネットネット株スクリーナー（グレアム流）")
    print(f"判定基準: 時価総額 <= NCAV x {NCAV_MARGIN:.3f} （流動資産-負債総額の2/3）")
    print("=" * 60)

    print("\n[銘柄リスト取得]")
    tickers = load_jp_tickers()
    if not tickers:
        print("銘柄リストが空です。終了します。")
        return

    if test_mode:
        tickers = tickers[:20]
        print(f"テストモード: 先頭{len(tickers)}銘柄のみ実行")

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
                print(f"-> NCAV:{r['ncav']:,.0f}  時価総額:{r['market_cap']:,.0f}  "
                      f"比率:{r['ncav_ratio']}  {ocf_str}  ★ネットネット")
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
