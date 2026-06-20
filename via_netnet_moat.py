"""
VIA ネットネット株専用モート分析スクリプト
via_netnet_results.csv のネットネット株（is_netnet=True）を対象に、
Claude APIでエコノミックモートを分析し via_moat.json に追記する。

via_moat_update.py の call_claude() を再利用するため、
両ファイルが同じフォルダにある必要があります。

使い方:
  python via_netnet_moat.py
"""

import sys, os, json, time
import pandas as pd
import yfinance as yf
import warnings
warnings.filterwarnings('ignore')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

import via_moat_update as vmu  # call_claude() を再利用

CSV_FILE  = os.path.join(SCRIPT_DIR, "via_netnet_results.csv")
MOAT_FILE = os.path.join(SCRIPT_DIR, "via_moat.json")
SLEEP_SEC = 1.0


def log(msg):
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


def get_financial_metrics(ticker):
    """モート分析用の簡易財務指標をyfinanceから取得"""
    try:
        tk = yf.Ticker(ticker)
        info = tk.info

        roe = info.get("returnOnEquity")
        roa = info.get("returnOnAssets")
        gm  = info.get("grossMargins")

        roe = round(roe * 100, 1) if roe else None
        roa = round(roa * 100, 1) if roa else None
        gm  = round(gm * 100, 1)  if gm  else None

        # ROIC簡易近似（取得困難なためROAで代替）
        roic = roa

        # 成長率: EPS成長率の簡易代替（売上高成長率）
        growth = info.get("revenueGrowth")
        growth = round(growth * 100, 1) if growth else None

        # D/E
        de = info.get("debtToEquity")
        de = round(de, 1) if de else None

        return roe, roa, roic, gm, growth, de
    except Exception:
        return None, None, None, None, None, None


def main():
    log("=" * 55)
    log("VIA ネットネット株 モート分析開始")
    log("=" * 55)

    if not os.path.exists(CSV_FILE):
        log(f"エラー: {CSV_FILE} が見つかりません")
        return

    df = pd.read_csv(CSV_FILE, dtype=str, encoding="utf-8-sig")
    df["is_netnet_bool"] = df["is_netnet"].astype(str).str.lower() == "true"
    netnet = df[df["is_netnet_bool"]].copy()
    log(f"ネットネット株: {len(netnet)} 社")

    # 既存モートデータ読み込み
    moat_data = {}
    if os.path.exists(MOAT_FILE):
        with open(MOAT_FILE, "r", encoding="utf-8") as f:
            moat_data = json.load(f)
    log(f"既存モート登録: {len(moat_data)} 銘柄")

    targets = [t for t in netnet["ticker"].tolist() if t not in moat_data]
    log(f"分析対象（未登録のみ）: {len(targets)} 銘柄")

    success, fail = 0, 0
    for i, ticker in enumerate(targets):
        row = netnet[netnet["ticker"] == ticker].iloc[0]
        name = row.get("name", ticker)
        sector = row.get("sector", "")

        log(f"[{i+1:3d}/{len(targets)}] {ticker:<10} {str(name)[:24]:<26} "
            f"({(i+1)/len(targets)*100:5.1f}%)")

        roe, roa, roic, gm, growth, de = get_financial_metrics(ticker)

        try:
            result = vmu.call_claude(
                ticker=ticker, name=name, sector=sector, market="JP",
                roe=roe, roa=roa, roic=roic, gm=gm, growth=growth, de=de,
                score="ネットネット株（NCAV戦略）"
            )
            moat_data[ticker] = result
            grade = result.get("moat_grade", "?").upper()
            bonus = result.get("bonus", 0)
            types = result.get("moat_type", [])
            log(f"  -> {grade}  bonus={bonus}  types={types}")
            success += 1
        except Exception as e:
            log(f"  [ERROR] 失敗: {e}")
            fail += 1

        time.sleep(SLEEP_SEC)

    with open(MOAT_FILE, "w", encoding="utf-8") as f:
        json.dump(moat_data, f, ensure_ascii=False, indent=2)

    log("=" * 55)
    log(f"完了: 成功={success}  失敗={fail}  合計{len(moat_data)}銘柄登録")
    log(f"保存先: {MOAT_FILE}")
    log("=" * 55)


if __name__ == "__main__":
    main()
