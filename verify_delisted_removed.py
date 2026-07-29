"""
日本株ユニバースの検証

フルスクリーニングを回さずに、銘柄リスト生成の段階だけを実行して確認する。

  1. ハードコードリストから削除した17銘柄が最終リストに現れないこと
  2. コード再利用で現在も上場している2銘柄（4556 / 8729）が残っていること
  3. 以前に問題となった 9384（内外トランスライン）が現れないこと
  4. 2024年以降の英字入りコード（417A形式）が対象に入っていること
  5. 財務データが取れない区分（ETF・PRO Market等）が除外されていること

使い方:
  python verify_delisted_removed.py
"""

import sys, os, json

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

import via_screener as vs

# 端末のエンコーディングに左右されず結果を読めるようにするための出力先
RESULT_FILE = os.path.join(SCRIPT_DIR, "verify_result.json")

# check_regional_delisted.py の点検で「廃止済み かつ 現行リストに無い」と
# 判定してハードコードリストから削除したコード
REMOVED = [
    "2743", "6839", "7315", "8245", "8279", "9260",   # nagoya
    "6556", "7092", "8087", "9386",                   # fukuoka
    "4764", "8740", "9070",                           # sapporo
    "5413", "7205", "8355", "9613",                   # nikkei225
]

# 廃止一覧に名前があるが現在も上場中のため意図的に残したコード
KEPT = ["4556", "8729"]

# 当初の発端となった銘柄
ORIGINAL_ISSUE = ["9384"]

# 2024年以降の英字入りコード。いずれも内国株式なので対象に含まれるべき。
NEW_FORMAT = {
    "130A": "Veritas In Silico / グロース",
    "141A": "トライアルHD / グロース",
    "147A": "ソラコム / グロース",
    "167A": "リョーサン菱洋HD / プライム",
}

# 内国株式以外の区分。実測でJ-Quantsから財務データが返らない、
# または普通株と同じ基準で扱えないため除外されるべき。
NON_DOMESTIC = {
    "1305": "ETF（iFreeETF TOPIX）",
    "131A": "PRO Market（CCNグループ）",
    "2971": "REIT（エスコンジャパンリート）",
    "4875": "外国株式（メディシノバ）",
    "8301": "出資証券（日本銀行）",
}


def main():
    print("[銘柄リストを生成]")
    print("-" * 50)
    jp = vs.load_jp_tickers()
    jp = list(dict.fromkeys(jp))
    jp = vs.filter_delisted_jp(jp)
    print(f"  最終JP銘柄数: {len(jp)}")
    print(f"  現行リスト由来: {len(vs.JP_LISTED_AUTHORITATIVE)}")

    codes = {t.replace(".T", "") for t in jp}

    print()
    print("=" * 50)
    print("検証結果")
    print("=" * 50)

    failures = []

    print("\n[1] 削除した廃止銘柄が含まれないこと")
    leaked = [c for c in REMOVED if c in codes]
    for c in REMOVED:
        mark = "NG 残存" if c in codes else "OK"
        print(f"    {c}: {mark}")
    if leaked:
        failures.append(f"削除済みのはずの銘柄が残存: {', '.join(leaked)}")

    print("\n[2] コード継続銘柄が残っていること")
    dropped = [c for c in KEPT if c not in codes]
    for c in KEPT:
        mark = "OK" if c in codes else "NG 消えた"
        print(f"    {c}: {mark}")
    if dropped:
        failures.append(f"上場中の銘柄が誤って除外: {', '.join(dropped)}")

    print("\n[3] 発端の銘柄が含まれないこと")
    for c in ORIGINAL_ISSUE:
        mark = "NG 残存" if c in codes else "OK"
        print(f"    {c}: {mark}")
        if c in codes:
            failures.append(f"{c} が残存")

    print("\n[4] 英字入りコードが対象に入っていること")
    missing_new = [c for c in NEW_FORMAT if c not in codes]
    for c, label in NEW_FORMAT.items():
        mark = "OK" if c in codes else "NG 未収録"
        print(f"    {c}: {mark}  {label}")
    if missing_new:
        failures.append(f"英字入りコードが未収録: {', '.join(missing_new)}")

    print("\n[5] 内国株式以外の区分が除外されていること")
    leftover = [c for c in NON_DOMESTIC if c in codes]
    for c, label in NON_DOMESTIC.items():
        mark = "NG 残存" if c in codes else "OK"
        print(f"    {c}: {mark}  {label}")
    if leftover:
        failures.append(f"内国株式以外が残存: {', '.join(leftover)}")

    print()
    print("=" * 50)
    if failures:
        print(f"検証失敗: {len(failures)}件")
        for f in failures:
            print(f"  - {f}")
    else:
        print("検証成功: すべての条件を満たしています")
    print("=" * 50)

    result = {
        "jp_total": len(jp),
        "authoritative_total": len(vs.JP_LISTED_AUTHORITATIVE),
        "alnum_code_total": sum(1 for c in codes if not c.isdigit()),
        "removed_still_present": leaked,
        "kept_wrongly_dropped": dropped,
        "original_issue_present": [c for c in ORIGINAL_ISSUE if c in codes],
        "new_format_missing": missing_new,
        "non_domestic_leftover": leftover,
        "passed": not failures,
    }
    with open(RESULT_FILE, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2)
    print(f"結果出力: {RESULT_FILE}")

    # 端末の文字化けや貼り付け事故に強い、ASCIIのみの一行サマリ
    print("RESULT jp={} auth={} alnum={} leaked={} dropped={} "
          "newmissing={} nondomestic={} verdict={}".format(
              len(jp),
              len(vs.JP_LISTED_AUTHORITATIVE),
              result["alnum_code_total"],
              ",".join(leaked) or "none",
              ",".join(dropped) or "none",
              ",".join(missing_new) or "none",
              ",".join(leftover) or "none",
              "PASS" if not failures else "FAIL",
          ))

    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
