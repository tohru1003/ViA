"""
地方取引所ハードコードリストの棚卸し

via_screener.py にハードコードされた地方取引所（名古屋/福岡/札幌）および
日経225フォールバックの銘柄コードを、JPXの上場廃止銘柄一覧
（当年ページ + バックナンバー全年度）と突き合わせて、
廃止済みコードが残っていないか点検する。

持株会社化などで同じコードが新会社へ引き継がれた銘柄は廃止一覧にも載るため、
現行の上場銘柄一覧（jpx_tickers.xls）とも照合し、
「廃止一覧にあり、かつ現行リストに無い」ものだけを削除対象として報告する。

使い方:
  python check_regional_delisted.py
"""

import sys, re, os, json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# via_screener のインポート時に Windows 端末向けの UTF-8 出力設定が行われる。
# ここで別途 sys.stdout を差し替えると、そのラッパーが回収される際に
# 出力先が閉じられてしまうため、差し替えは行わない。
import via_screener as vs

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_FILE = os.path.join(SCRIPT_DIR, "regional_delisted_report.json")


def fetch_delisted_with_archives():
    """当年ページとバックナンバー全ページから上場廃止コードを取得"""
    print("JPX上場廃止一覧を取得中...")
    delisted, pages = vs._scrape_jpx_delisted_codes(verbose=True)
    print(f"  合計: {len(delisted)}件 / {len(pages)}ページ（重複除去後）")
    if len(pages) <= 1:
        print("  ※ バックナンバーを取得できませんでした。")
        print("     当年分のみでの点検になるため、過去の廃止銘柄は見逃されます。")
    return delisted, pages


def current_listed_codes():
    """
    現在のJPX上場銘柄一覧（jpx_tickers.xls / .xlsx）のコード集合を返す。

    持株会社化・会社分割などで同じコードが新会社に引き継がれた銘柄は、
    上場廃止一覧に載っていても現在も上場中である。誤削除を防ぐため、
    廃止判定はこの現行リストとの突き合わせで確定させる。
    """
    for ext in ("xlsx", "xls"):
        path = os.path.join(SCRIPT_DIR, f"jpx_tickers.{ext}")
        if not os.path.exists(path):
            continue
        tickers = vs._parse_jpx(path)
        if tickers:
            return {t.replace(".T", "").upper() for t in tickers}
    return set()


def hardcoded_groups():
    """via_screener.py のハードコードリストを再取得"""
    src = open(os.path.join(SCRIPT_DIR, "via_screener.py"), encoding="utf-8").read()

    groups = {}
    for name in ("nagoya", "fukuoka", "sapporo"):
        m = re.search(rf"{name}\s*=\s*\[(.*?)\]", src, re.DOTALL)
        groups[name] = re.findall(r'"([0-9A-Za-z]{4,5})"', m.group(1)) if m else []

    # 日経225フォールバック
    nk = vs._nk225()
    groups["nikkei225"] = [t.replace(".T", "") for t in nk]
    return groups


def main():
    delisted, pages = fetch_delisted_with_archives()
    if not delisted:
        print("上場廃止リストが取得できませんでした。中止します。")
        return

    listed = current_listed_codes()
    if listed:
        print(f"現行の上場銘柄一覧: {len(listed)}銘柄")
    else:
        print("※ jpx_tickers.xls が読めませんでした。")
        print("   コード継続銘柄を判別できないため、削除対象の確定は保留します。")

    groups = hardcoded_groups()

    print()
    print("=" * 60)
    print("ハードコードリストの点検結果")
    print("=" * 60)

    report = {
        "pages": pages,
        "delisted_total": len(delisted),
        "listed_total": len(listed),
        "remove": {},
        "keep": {},
    }
    total_remove = total_keep = 0

    for name, codes in groups.items():
        hits = sorted(c for c in codes if c.upper() in delisted)
        # 廃止一覧にあり、かつ現行リストに無いものだけを削除対象とする
        remove = [c for c in hits if c.upper() not in listed] if listed else []
        keep = [c for c in hits if c.upper() in listed] if listed else hits

        report["remove"][name] = remove
        report["keep"][name] = keep
        total_remove += len(remove)
        total_keep += len(keep)

        print(f"\n[{name}] {len(codes)}銘柄  廃止一覧ヒット{len(hits)}件")
        if remove:
            print(f"  削除対象（現在は非上場）: {', '.join(remove)}")
        if keep:
            print(f"  保持（コード継続で上場中）: {', '.join(keep)}")
        if not hits:
            print("  問題なし")

    print()
    print("=" * 60)
    print(f"削除対象: {total_remove}件 / コード継続: {total_keep}件")
    print("=" * 60)

    with open(REPORT_FILE, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"レポート出力: {REPORT_FILE}")


if __name__ == "__main__":
    main()
