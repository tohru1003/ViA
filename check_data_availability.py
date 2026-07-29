"""
市場・商品区分ごとの財務データ取得可否の実測

「ETF・REITは指標が取れない」「外国株式はどうか」を推測で決めず、
スクリーナーが実際に使っているデータ経路を叩いて確認する。

  - J-Quants (via_screener.get_jp_financials): 日本株の財務データ本体
  - yfinance: 株価と補助的な指標

区分ごとにサンプルを取り、EPS/ROE が返るかを数える。

使い方:
  python check_data_availability.py
"""

import sys, os, re, json

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

import pandas as pd
import via_screener as vs

RESULT_FILE = os.path.join(SCRIPT_DIR, "data_availability_result.json")

# 各区分から何銘柄ずつ試すか。多いとAPI負荷が上がるので控えめにする。
SAMPLES_PER_MARKET = 3

CODE_RE = re.compile(r'^\d{4}$|^\d{5}$|^\d{3}[A-Za-z]$')


def load_rows():
    path = None
    for ext in ("xlsx", "xls"):
        p = os.path.join(SCRIPT_DIR, f"jpx_tickers.{ext}")
        if os.path.exists(p):
            path = p
            break
    if path is None:
        return None

    ext = os.path.splitext(path)[1].lower()
    engine = "openpyxl" if ext == ".xlsx" else "xlrd"
    df = pd.read_excel(path, dtype=str, engine=engine)

    code_col = next((c for c in df.columns if "コード" in str(c)), None)
    name_col = next((c for c in df.columns if "銘柄名" in str(c)), None)
    mkt_col = next((c for c in df.columns if "市場" in str(c)), None)
    if not (code_col and mkt_col):
        return None

    out = []
    for _, r in df.iterrows():
        code = str(r[code_col]).strip()
        if not CODE_RE.match(code):
            continue
        out.append({
            "code": code,
            "name": str(r[name_col]) if name_col else "",
            "market": str(r[mkt_col]).strip(),
        })
    return out


def probe_jquants(code):
    """J-Quants経由で財務データが返るか"""
    try:
        fin = vs.get_jp_financials(code)
    except Exception as e:
        return {"ok": False, "error": str(e)[:100]}

    if not fin:
        return {"ok": False, "error": "None が返却"}

    eps = [v for v in (fin.get("eps") or []) if v is not None]
    roe = [v for v in (fin.get("roe") or []) if v is not None]
    return {
        "ok": bool(eps or roe),
        "eps_count": len(eps),
        "roe_count": len(roe),
        "latest_EPS": fin.get("latest_EPS"),
        "latest_ROE": fin.get("latest_ROE"),
    }


def probe_yfinance(code):
    """yfinance で株価が引けるか"""
    try:
        import yfinance as yf
        hist = yf.Ticker(f"{code}.T").history(period="5d")
        return {"ok": bool(hist is not None and len(hist)), "rows": len(hist) if hist is not None else 0}
    except Exception as e:
        return {"ok": False, "error": str(e)[:100]}


def main():
    rows = load_rows()
    if not rows:
        print("jpx_tickers.xls を読めませんでした。")
        return 1

    by_market = {}
    for r in rows:
        by_market.setdefault(r["market"], []).append(r)

    print(f"読込: {len(rows)}銘柄 / {len(by_market)}区分")
    print(f"各区分から最大{SAMPLES_PER_MARKET}銘柄を実測します。")

    report = {}
    for market in sorted(by_market, key=lambda m: -len(by_market[m])):
        items = by_market[market][:SAMPLES_PER_MARKET]
        print()
        print("=" * 64)
        print(f"[{market}]  該当 {len(by_market[market])}銘柄")
        print("=" * 64)

        entries = []
        for it in items:
            jq = probe_jquants(it["code"])
            yf_ = probe_yfinance(it["code"])
            entries.append({**it, "jquants": jq, "yfinance": yf_})

            jq_txt = ("OK eps={} roe={}".format(jq.get("eps_count"), jq.get("roe_count"))
                      if jq["ok"] else "NG " + str(jq.get("error", ""))[:40])
            yf_txt = "OK" if yf_["ok"] else "NG"
            print(f"  {it['code']:6s} {it['name'][:22]:24s} JQ:{jq_txt:26s} YF:{yf_txt}")

        report[market] = {
            "total": len(by_market[market]),
            "jquants_ok": sum(1 for e in entries if e["jquants"]["ok"]),
            "yfinance_ok": sum(1 for e in entries if e["yfinance"]["ok"]),
            "sampled": len(entries),
            "entries": entries,
        }

    print()
    print("=" * 64)
    print("区分別サマリ（サンプル中の成功数）")
    print("=" * 64)
    for market, r in sorted(report.items(), key=lambda kv: -kv[1]["total"]):
        print(f"  {r['total']:5d}銘柄  JQ {r['jquants_ok']}/{r['sampled']}  "
              f"YF {r['yfinance_ok']}/{r['sampled']}  {market}")

    with open(RESULT_FILE, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"\n結果出力: {RESULT_FILE}")

    compact = ";".join(
        f"{m.replace('（','(').replace('）',')')}={r['jquants_ok']}/{r['sampled']}"
        for m, r in sorted(report.items(), key=lambda kv: -kv[1]["total"]))
    print(f"RESULT {compact}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
