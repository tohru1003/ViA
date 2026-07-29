"""
JPX上場銘柄一覧で取りこぼしているコードの調査

_parse_jpx() は 4桁/5桁の数字コードしか抽出しないため、
2024年からJPXが発行している英字入りの新形式コード（417A 等）が捨てられている。
実際に何が捨てられているのかを、形式別に分類して確認する。

使い方:
  python check_new_codes.py
"""

import sys, os, re, json

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

import pandas as pd

RESULT_FILE = os.path.join(SCRIPT_DIR, "new_codes_result.json")

# _parse_jpx() が現在拾っているパターン
CURRENT = re.compile(r'^\d{4}$|^\d{5}$')

# 分類用パターン
PATTERNS = [
    ("digit4", re.compile(r'^\d{4}$')),
    ("digit5", re.compile(r'^\d{5}$')),
    ("digit3_alpha1", re.compile(r'^\d{3}[A-Za-z]$')),
    ("digit4_alpha1", re.compile(r'^\d{4}[A-Za-z]$')),
]


def classify(code):
    for name, pat in PATTERNS:
        if pat.match(code):
            return name
    return "other"


def probe_alphanumeric_tickers():
    """
    英字入りコードが yfinance で実際に引けるかを確認する。

    引けないなら対象に加えても失敗が増えるだけなので、
    _parse_jpx() を直す前に確かめておく。
    """
    targets = [
        ("147A.T", "ソラコム / グロース"),
        ("167A.T", "リョーサン菱洋HD / プライム"),
        ("141A.T", "トライアルHD / グロース"),
        ("7203.T", "トヨタ / 対照用の既存4桁コード"),
    ]

    print()
    print("=" * 60)
    print("yfinance で英字入りティッカーが引けるかの確認")
    print("=" * 60)

    try:
        import yfinance as yf
    except Exception as e:
        print(f"  yfinance をインポートできません: {e}")
        return {"error": str(e)}

    out = {}
    for ticker, label in targets:
        try:
            hist = yf.Ticker(ticker).history(period="1mo")
            rows = 0 if hist is None else len(hist)
            last = None
            if rows:
                last = round(float(hist["Close"].iloc[-1]), 2)
            out[ticker] = {"rows": rows, "last_close": last}
            state = f"OK {rows}行 終値={last}" if rows else "NG データなし"
            print(f"    {ticker:9s} {state:28s} {label}")
        except Exception as e:
            out[ticker] = {"error": str(e)[:120]}
            print(f"    {ticker:9s} NG {str(e)[:60]:25s} {label}")
    return out


def main():
    path = None
    for ext in ("xlsx", "xls"):
        p = os.path.join(SCRIPT_DIR, f"jpx_tickers.{ext}")
        if os.path.exists(p):
            path = p
            break
    if path is None:
        print("jpx_tickers.xls / .xlsx が見つかりません。")
        return 1

    ext = os.path.splitext(path)[1].lower()
    engine = "openpyxl" if ext == ".xlsx" else "xlrd"
    df = pd.read_excel(path, dtype=str, engine=engine)
    print(f"読込: {path}")
    print(f"  {df.shape[0]}行 × {df.shape[1]}列")

    code_col = next((c for c in df.columns if "コード" in str(c)), None)
    name_col = next((c for c in df.columns if "銘柄名" in str(c)), None)
    mkt_col = next((c for c in df.columns if "市場" in str(c)), None)
    if code_col is None:
        print(f"コード列が見つかりません: {list(df.columns)}")
        return 1

    rows = df[[c for c in (code_col, name_col, mkt_col) if c]].dropna(subset=[code_col])
    rows = rows.copy()
    rows[code_col] = rows[code_col].astype(str).str.strip()

    buckets = {}
    for _, r in rows.iterrows():
        code = r[code_col]
        kind = classify(code)
        buckets.setdefault(kind, []).append({
            "code": code,
            "name": str(r[name_col]) if name_col else "",
            "market": str(r[mkt_col]) if mkt_col else "",
        })

    print()
    print("=" * 60)
    print("コード形式別の内訳")
    print("=" * 60)
    for kind, items in sorted(buckets.items(), key=lambda kv: -len(kv[1])):
        picked = "拾えている" if CURRENT.match(items[0]["code"]) else "捨てている"
        print(f"\n[{kind}] {len(items)}件  ({picked})")
        for it in items[:8]:
            print(f"    {it['code']}  {it['market']}  {it['name']}")
        if len(items) > 8:
            print(f"    ... 他 {len(items) - 8}件")

    missed = [it for kind, items in buckets.items()
              if not CURRENT.match(items[0]["code"]) for it in items]
    picked = [it for kind, items in buckets.items()
              if CURRENT.match(items[0]["code"]) for it in items]

    def tally(items):
        out = {}
        for it in items:
            out[it["market"]] = out.get(it["market"], 0) + 1
        return out

    by_market = tally(missed)

    print()
    print("=" * 60)
    print(f"取りこぼし合計: {len(missed)}件（市場区分別）")
    print("=" * 60)
    for mkt, n in sorted(by_market.items(), key=lambda kv: -kv[1]):
        print(f"    {n:5d}  {mkt}")

    # 現在拾えているコードの内訳。ETF/REIT/PRO Market が
    # どれだけ混ざっているかを確認し、市場区分での選別可否を判断する。
    picked_by_market = tally(picked)
    print()
    print("=" * 60)
    print(f"現在拾えている{len(picked)}件の市場区分別内訳")
    print("=" * 60)
    for mkt, n in sorted(picked_by_market.items(), key=lambda kv: -kv[1]):
        print(f"    {n:5d}  {mkt}")

    # 内国株式のみに絞った場合の銘柄数
    def is_domestic(mkt):
        return "内国株式" in mkt

    domestic_total = (sum(n for m, n in picked_by_market.items() if is_domestic(m))
                      + sum(n for m, n in by_market.items() if is_domestic(m)))
    print()
    print(f"  内国株式のみに絞った場合: {domestic_total}銘柄")

    probe = probe_alphanumeric_tickers()

    result = {
        "total_rows": int(df.shape[0]),
        "counts_by_kind": {k: len(v) for k, v in buckets.items()},
        "missed_total": len(missed),
        "missed_by_market": by_market,
        "picked_total": len(picked),
        "picked_by_market": picked_by_market,
        "domestic_only_total": domestic_total,
        "yfinance_probe": probe,
        "missed_samples": missed[:40],
    }
    with open(RESULT_FILE, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n結果出力: {RESULT_FILE}")

    probe_ok = sum(1 for v in probe.values()
                   if isinstance(v, dict) and v.get("rows"))
    print(f"RESULT rows={df.shape[0]} missed={len(missed)} "
          f"domestic={domestic_total} probe_ok={probe_ok}/{len(probe)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
