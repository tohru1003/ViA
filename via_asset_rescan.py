"""
VIA 資産過小評価スコア 再計算専用スクリプト
既存のvia_results.csvを読み込み、VIA通過の日本株に対して
資産過小評価スコア（NCAV+投資有価証券+賃貸等不動産含み益-税効果）を
再計算してCSV・HTMLを更新する。

全銘柄スクリーニングをやり直す必要がないため高速（EDINETキャッシュ活用）。

使い方:
  python via_asset_rescan.py
"""

import sys, os, time
import pandas as pd
import yfinance as yf
import warnings
warnings.filterwarnings('ignore')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

import via_screener as v  # calc_asset_undervaluation等を再利用

CSV_FILE = os.path.join(SCRIPT_DIR, "via_results.csv")


def main():
    print("=" * 60)
    print("VIA 資産過小評価スコア 再計算")
    print("=" * 60)

    if not os.path.exists(CSV_FILE):
        print(f"エラー: {CSV_FILE} が見つかりません")
        return

    df = pd.read_csv(CSV_FILE, dtype=str, encoding="utf-8-sig")
    print(f"CSV読み込み: {len(df)} 行")

    passed_mask = df["grade"].isin(["STRICT", "RELAXED"])
    passed_jp = df[passed_mask & (df["market"] == "JP")]
    print(f"対象（VIA通過の日本株）: {len(passed_jp)} 銘柄")

    if v._edinet is None:
        print("エラー: via_edinet_moduleが読み込めません")
        return

    for idx, row in passed_jp.iterrows():
        ticker = row["ticker"]
        print(f"  {ticker:<10}", end=" ", flush=True)
        try:
            tk_tmp = yf.Ticker(ticker)
            mcap = tk_tmp.fast_info.market_cap
        except Exception:
            mcap = None

        asset_result = v.calc_asset_undervaluation(ticker, row["market"], mcap)
        if asset_result:
            def _s(v):
                return str(v) if v is not None else None
            df.at[idx, "ncav"] = _s(asset_result.get("ncav"))
            df.at[idx, "ncav_plus_tax_adjusted"] = _s(asset_result.get("ncav_plus_tax_adjusted"))
            df.at[idx, "asset_undervaluation_ratio"] = _s(asset_result.get("undervaluation_ratio"))
            df.at[idx, "is_asset_undervalued"] = _s(asset_result.get("is_undervalued"))
            vd_data = asset_result.get("valuation_diff_data") or {}
            rp_data = asset_result.get("rental_property_data") or {}
            df.at[idx, "valuation_diff"] = _s(vd_data.get("valuation_diff"))
            df.at[idx, "rental_property_gain"] = _s(rp_data.get("unrealized_gain"))
            df.at[idx, "rental_property_book_value"] = _s(rp_data.get("book_value"))
            df.at[idx, "rental_property_fair_value"] = _s(rp_data.get("fair_value"))
            ratio_disp = asset_result.get("undervaluation_ratio")
            rp_str = (f" 不動産含み益:{rp_data.get('unrealized_gain',0):,.0f}円"
                      if rp_data.get("unrealized_gain") else "")
            print(f"-> 比率:{ratio_disp}{rp_str}  "
                  f"{'★過小評価' if asset_result.get('is_undervalued') else ''}")
        else:
            print("-> データ取得不可")
        time.sleep(0.5)

    df.to_csv(CSV_FILE, index=False, encoding="utf-8-sig")
    print(f"\nCSV更新完了: {CSV_FILE}")

    # HTML再生成
    print("\nHTML再生成中...")
    try:
        import importlib
        import via_rebuild_html as vrh
        importlib.reload(vrh)
        import json
        moat_path = os.path.join(SCRIPT_DIR, "via_moat.json")
        moat_data = {}
        if os.path.exists(moat_path):
            with open(moat_path, "r", encoding="utf-8") as f:
                moat_data = json.load(f)
        df_html = pd.read_csv(CSV_FILE, dtype=str, encoding="utf-8-sig")
        df_html = vrh.recalc_scores(df_html)
        from datetime import datetime
        generated = datetime.now().strftime("%Y-%m-%d %H:%M")
        html = vrh.make_html(df_html, len(df_html), generated, moat_data)
        out_html = os.path.join(SCRIPT_DIR, "via_results.html")
        with open(out_html, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"HTML再生成完了: {out_html}")
    except Exception as e:
        print(f"HTML再生成失敗: {e}")
        import traceback
        traceback.print_exc()

    print("\n完了")


if __name__ == "__main__":
    main()
