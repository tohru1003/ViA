"""
VIA スクリーナー — 財務判定 + 購買ターゲット価格比較 + 厳格/緩和ラベル
使い方:
  pip install yfinance pandas numpy requests openpyxl
  python via_screener.py
"""

import sys, io
# Windows端末でのUTF-8出力を強制
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import yfinance as yf
import pandas as pd
import numpy as np
import os, time, warnings, urllib.request, ssl, sys
from datetime import datetime

try:
    import via_edinet_module as _edinet
except ImportError:
    _edinet = None

warnings.filterwarnings('ignore')

# ── 設定 ──────────────────────────────────────────────
WACC             = 8.0
SLEEP_SEC        = 0.5
JQUANTS_API_KEY  = "TMhn3YeV3FTDnbCaW6pm6irEqvbocWKoooVYGySQPFI"
JQUANTS_BASE     = "https://api.jquants.com/v2"

STRICT = dict(
    pass_ratio  = 0.80,
    de_thr      = 50,
    roe_thr     = 15,
    roa_thr     = 7,
    roic_thr    = 15,
    consec_drop = 2,
    max_years   = 9,    # 最大9年分で評価（古すぎるデータを除外）
    min_years   = 5,    # EPS上昇期間として必要な最低年数（厳格基準）
)
RELAXED = dict(
    pass_ratio   = 0.70,
    de_thr       = 100,
    roe_thr      = 12,
    roa_thr      = 5,
    roic_thr     = 10,
    consec_drop  = 3,
    recent_years = 4,   # 直近4年のデータのみで評価
    min_years    = 4,   # EPS上昇期間として必要な最低年数（緩和基準）
)

# ── J-Quants API設定（日本株）──
JQUANTS_API_KEY = "TMhn3YeV3FTDnbCaW6pm6irEqvbocWKoooVYGySQPFI"
JQUANTS_BASE    = "https://api.jquants.com/v2"

# ── EODHD API設定（米国株）──
EODHD_TOKEN = "6a17e87bdceab5.12312454"
EODHD_BASE  = "https://eodhd.com/api"

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
# ── 財務データキャッシュ設定 ──
CACHE_DIR          = os.path.join(SCRIPT_DIR, "via_cache")
FINS_CACHE_DAYS    = 30    # 財務データの有効期限（日）
EARNINGS_CACHE_DAYS= 7     # 決算発表チェックの有効期限（日）
PRICE_CACHE_HOURS  = 12    # 株価キャッシュの有効期限（時間）

DCF_CONT_YEARS  = 10
DCF_DISC_RATE   = 10
DCF_INFL_RATE   = 2
DCF_SURV_YEARS  = 10
DCF_MARGIN_SAFE = 30
DCF_ADJ_CAP     = 12

OUTPUT_CSV  = "via_results.csv"
OUTPUT_HTML = "via_results.html"
# ──────────────────────────────────────────────────────


# ── 銘柄リスト取得 ─────────────────────────────────────

def _make_ssl_ctx():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode    = ssl.CERT_NONE
    return ctx

def _http_get(url, timeout=30):
    req = urllib.request.Request(url, headers={
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"),
        "Accept":     "*/*",
        "Referer":    "https://www.jpx.co.jp/",
    })
    with urllib.request.urlopen(req, timeout=timeout,
                                context=_make_ssl_ctx()) as r:
        return r.read()


def load_jp_tickers():
    """東証全銘柄を自動ダウンロード。失敗時は日経225。"""

    # ── 手動配置ファイルを最優先で確認 ──
    # jpx_tickers.xlsx / .xls がある場合は年齢に関係なく使用
    for ext in ["xlsx", "xls"]:
        manual = os.path.join(SCRIPT_DIR, f"jpx_tickers.{ext}")
        if os.path.exists(manual):
            age_h = (time.time() - os.path.getmtime(manual)) / 3600
            print(f"  JPXファイル検出: {manual} ({age_h:.1f}h前)")
            result = _parse_jpx(manual)
            if result:
                # 地方取引所独自銘柄を追加
                regional = _regional_exchanges()
                # 重複除去（東証リストにない銘柄のみ追加）
                result_set = set(result)
                added = [t for t in regional if t not in result_set]
                result = result + added
                print(f"  → 東証+地方取引所: {len(result)} 銘柄 "
                      f"（地方独自追加: {len(added)}）")
                return result
            else:
                print(f"  → パース失敗。ファイルの形式を確認してください。")

    # ── JPX公式からダウンロード（xlsx優先、xls次点） ──
    # JPX公式URL（xlsのみ提供）
    JPX_URL = ("https://www.jpx.co.jp/markets/statistics-equities/misc/"
               "tvdivq0000001vg2-att/data_j.xls")
    save = os.path.join(SCRIPT_DIR, "jpx_tickers.xls")

    print("  東証上場銘柄一覧をダウンロード中...")
    print(f"  URL: {JPX_URL}")

    # requests が使えれば requests で、なければ urllib で試みる
    data = None
    headers = {
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Accept": "application/vnd.ms-excel,*/*",
        "Accept-Language": "ja,en;q=0.9",
        "Referer": "https://www.jpx.co.jp/markets/statistics-equities/misc/01.html",
        "Connection": "keep-alive",
    }

    # 方法1: requests（セッション付き）
    try:
        import requests as req_lib
        session = req_lib.Session()
        session.headers.update(headers)
        # まずトップページにアクセスしてCookieを取得
        try:
            session.get("https://www.jpx.co.jp/markets/statistics-equities/misc/01.html",
                       timeout=10, verify=False)
        except Exception:
            pass
        r = session.get(JPX_URL, timeout=30, verify=False)
        if r.status_code == 200 and len(r.content) > 1000:
            data = r.content
            print(f"  → requests で取得成功 ({len(data):,} bytes)")
        else:
            print(f"  → requests HTTP {r.status_code}")
    except ImportError:
        print("  requests未インストール → urllibで試みます")
    except Exception as e:
        print(f"  requests 失敗: {e}")

    # 方法2: urllib（requestsが失敗した場合）
    if data is None:
        try:
            data = _http_get(JPX_URL)
            print(f"  → urllib で取得成功 ({len(data):,} bytes)")
        except Exception as e:
            print(f"  urllib 失敗: {e}")

    # ファイル保存・パース
    if data and len(data) > 1000:
        try:
            with open(save, "wb") as f:
                f.write(data)
            result = _parse_jpx(save)
            if result:
                print(f"  → 東証: {len(result)} 銘柄")
                return result
            else:
                print("  → パース失敗（コード列が見つからない）")
        except Exception as e:
            print(f"  保存/パースエラー: {e}")

    # ── フォールバック ──
    print()
    print("  ★ JPXダウンロード失敗")
    print("  ★ 手動でダウンロードする場合:")
    print("  ★   https://www.jpx.co.jp/markets/statistics-equities/misc/01.html")
    print("  ★   → 「上場銘柄一覧」Excelアイコンをクリック")
    print(f"  ★   → ダウンロードしたファイルを {SCRIPT_DIR} に")
    print("  ★     「jpx_tickers.xls」または「jpx_tickers.xlsx」として保存")
    print("  → 日経225にフォールバック")
    nk = _nk225()
    regional = _regional_exchanges()
    nk_set = set(nk)
    added = [t for t in regional if t not in nk_set]
    return nk + added


def _parse_jpx(path):
    """東証ExcelからTickerリストを生成"""
    try:
        ext = os.path.splitext(path)[1].lower()
        engine = "openpyxl" if ext == ".xlsx" else "xlrd"
        df = pd.read_excel(path, dtype=str, engine=engine)

        print(f"  ファイル読込: {df.shape[0]}行 × {df.shape[1]}列")
        print(f"  列名: {list(df.columns[:8])}")

        # 「コード」列を探す（複数パターン対応）
        code_col = None
        for col in df.columns:
            col_str = str(col).strip()
            if any(k in col_str for k in ["コード", "code", "Code", "ticker", "銘柄コード"]):
                code_col = col
                break
        if code_col is None:
            # 最初の列に数字が多ければそれをコード列とみなす
            for col in df.columns:
                sample = df[col].dropna().astype(str).str.strip()
                numeric = sample[sample.str.match(r'^\d{4,5}$')]
                if len(numeric) > 100:
                    code_col = col
                    print(f"  コード列を自動検出: 「{col}」")
                    break

        if code_col is None:
            print(f"  コード列が見つかりません。列名を確認してください: {list(df.columns)}")
            return []

        codes = df[code_col].dropna().astype(str).str.strip()
        # 4桁または5桁の数字コードを抽出（東証は4桁、一部5桁あり）
        codes_4 = codes[codes.str.match(r'^\d{4}$')]
        codes_5 = codes[codes.str.match(r'^\d{5}$')]
        print(f"  4桁コード: {len(codes_4)}件  5桁コード: {len(codes_5)}件")

        all_codes = pd.concat([codes_4, codes_5]).drop_duplicates().tolist()
        if len(all_codes) < 10:
            print(f"  コード数が少なすぎます({len(all_codes)}件)。列「{code_col}」のサンプル: {codes.head(5).tolist()}")
            return []

        tickers = [c + ".T" for c in all_codes]
        return tickers
    except Exception as e:
        print(f"  Excelパースエラー ({path}): {e}")
        import traceback
        traceback.print_exc()
        return []


def _nk225():
    codes = [
        "1332","1605","1721","1801","1802","1803","1808","1812","1925","1928",
        "1963","2002","2269","2282","2413","2432","2501","2502","2503","2531",
        "2768","2801","2802","2871","2914","3086","3099","3101","3105","3289",
        "3401","3402","3405","3407","3436","3659","3861","3863","4004","4005",
        "4042","4043","4061","4062","4063","4151","4183","4188","4208","4272",
        "4307","4324","4452","4502","4503","4506","4507","4519","4523","4528",
        "4543","4568","4578","4631","4689","4704","4751","4755","4901","4902",
        "4911","5001","5019","5020","5101","5108","5110","5201","5202","5214",
        "5233","5301","5332","5333","5334","5401","5406","5411","5412","5413",
        "5444","5463","5471","5631","5706","5707","5711","5713","5714","5715",
        "5801","5802","5803","5901","6103","6113","6178","6301","6302","6305",
        "6326","6361","6367","6471","6472","6473","6479","6501","6503","6504",
        "6506","6526","6594","6645","6674","6701","6702","6703","6724","6752",
        "6753","6758","6762","6770","6841","6857","6861","6902","6952","6954",
        "6971","6976","6981","6988","7003","7004","7011","7012","7013","7186",
        "7201","7202","7203","7205","7211","7261","7267","7269","7270","7272",
        "7731","7733","7735","7741","7751","7752","7762","7832","7911","7912",
        "7951","8001","8002","8003","8015","8031","8035","8053","8058","8113",
        "8233","8252","8267","8306","8308","8309","8316","8331","8354","8355",
        "8411","8601","8604","8630","8697","8725","8729","8750","8766","8795",
        "8801","8802","8804","8830","9001","9005","9007","9008","9009","9020",
        "9021","9022","9064","9101","9104","9107","9202","9301","9432","9433",
        "9434","9501","9502","9503","9531","9532","9602","9613","9735","9766",
    ]
    print(f"  → 日経225: {len(codes)} 銘柄（フォールバック）")
    return [c + ".T" for c in codes]

def _regional_exchanges():
    """名古屋・福岡・札幌 独自上場銘柄（東証非重複）"""
    nagoya = [
        "2743","3543","3918","4119","5816","5917","6161","6247","6378",
        "6387","6463","6551","6839","7220","7256","7264","7315","7322",
        "7363","7427","7528","7599","7643","7841","8119","8142","8217",
        "8244","8245","8279","8281","9260","9273","9322","9324","9325",
    ]
    fukuoka = [
        "2764","3077","3134","3172","3354","3371","3395","3443","3521",
        "3565","3607","3662","3693","4025","4082","4556","4571","5943",
        "6040","6250","6381","6556","7060","7092","7177","7683","8087",
        "8574","9028","9380","9381","9384","9386","9389","9423","9514",
        "9540","9553","9616","9628","9663","9678","9686","9699","9715",
    ]
    sapporo = [
        "2764","3093","3135","3548","4764","6072","6999","7821","8093",
        "8740","9070","9279","9446","9467","9535","9633","9658",
    ]
    codes = list(set(nagoya + fukuoka + sapporo))
    tickers = [c + ".T" for c in codes]
    print(f"  → 地方取引所（名古屋/福岡/札幌）独自銘柄: {len(tickers)}")
    return tickers



def load_us_tickers():
    """NYSE+NASDAQ全銘柄をGitHubから取得。失敗時はS&P500。"""

    # ── キャッシュ確認 ──
    cache = os.path.join(SCRIPT_DIR, "us_tickers_cache.txt")
    if os.path.exists(cache):
        age_h = (time.time() - os.path.getmtime(cache)) / 3600
        if age_h < 24:
            with open(cache) as f:
                t = [l.strip() for l in f if l.strip()]
            print(f"  US キャッシュ使用: {len(t)} 銘柄 ({age_h:.1f}h前)")
            return t

    # ── GitHubからダウンロード ──
    url = ("https://raw.githubusercontent.com/"
           "rreichel3/US-Stock-Symbols/main/all/all_tickers.txt")
    try:
        print("  NYSE+NASDAQ全銘柄をダウンロード中...")
        data = _http_get(url, timeout=20)
        raw  = [l.strip() for l in data.decode().strip().split('\n') if l.strip()]
        t    = [x for x in raw if x.isalpha() and 1 <= len(x) <= 5]
        print(f"  → NYSE+NASDAQ: {len(t)} 銘柄")
        with open(cache, "w") as f:
            f.write("\n".join(t))
        return t
    except Exception as e:
        print(f"  GitHubダウンロード失敗: {e} → S&P500にフォールバック")
        return _sp500()


def _sp500():
    url = ("https://raw.githubusercontent.com/datasets/"
           "s-and-p-500-companies/main/data/constituents.csv")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=15) as r:
        sp500 = pd.read_csv(r)
    t = sp500["Symbol"].str.replace(".", "-", regex=False).tolist()
    print(f"  → S&P500: {len(t)} 銘柄（フォールバック）")
    return t


def fetch_jquants(path, params=""):
    """J-Quants API V2 からデータ取得"""
    import ssl
    url = f"{JQUANTS_BASE}{path}?{params}" if params else f"{JQUANTS_BASE}{path}"
    req = urllib.request.Request(url, headers={"x-api-key": JQUANTS_API_KEY})
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    with urllib.request.urlopen(req, timeout=15, context=ctx) as r:
        return json.loads(r.read())


def get_jp_financials(code4):
    """
    J-Quants から日本株の財務データを取得して
    process_ticker と同じ形式の辞書を返す。
    code4: 4桁コード（例: "7792"）
    """
    import json as _json
    try:
        data = fetch_jquants("/fins/summary", f"code={code4}")
        stmts = data.get("data", [])
        if not stmts:
            return None

        # 年次決算（FY）のみ抽出、古い順
        fy = sorted(
            [s for s in stmts if "FY" in s.get("DocType", "")],
            key=lambda x: x.get("DiscDate", "")
        )
        if not fy:
            return None

        def to_f(v):
            try: return float(v) if v and v != "" else None
            except: return None

        eps_arr  = [to_f(s.get("EPS"))  for s in fy]
        np_arr   = [to_f(s.get("NP"))   for s in fy]
        op_arr   = [to_f(s.get("OP"))   for s in fy]
        sales_arr= [to_f(s.get("Sales"))for s in fy]
        eq_arr   = [to_f(s.get("Eq"))   for s in fy]
        ta_arr   = [to_f(s.get("TA"))   for s in fy]
        eqar_arr = [to_f(s.get("EqAR")) for s in fy]
        cfo_arr  = [to_f(s.get("CFO"))  for s in fy]
        cfi_arr  = [to_f(s.get("CFI"))  for s in fy]

        # ROE = NP/Eq×100, ROA = NP/TA×100
        roe_arr = [
            np_arr[i]/eq_arr[i]*100
            if np_arr[i] is not None and eq_arr[i] and eq_arr[i] != 0 else None
            for i in range(min(len(np_arr), len(eq_arr)))
        ]
        roa_arr = [
            np_arr[i]/ta_arr[i]*100
            if np_arr[i] is not None and ta_arr[i] and ta_arr[i] != 0 else None
            for i in range(min(len(np_arr), len(ta_arr)))
        ]

        # D/E = (1-EqAR)/EqAR×100
        de_arr = [
            (1 - eqar_arr[i]) / eqar_arr[i] * 100
            if eqar_arr[i] is not None and 0 < eqar_arr[i] < 1 else None
            for i in range(len(eqar_arr))
        ]

        # FCF = CFO + CFI
        fcf_arr = [
            cfo_arr[i] + cfi_arr[i]
            if cfo_arr[i] is not None and cfi_arr[i] is not None else None
            for i in range(min(len(cfo_arr), len(cfi_arr)))
        ]

        # 営業利益率（粗利代替）= OP/Sales×100
        gm_arr = [
            op_arr[i]/sales_arr[i]*100
            if op_arr[i] is not None and sales_arr[i] and sales_arr[i] != 0 else None
            for i in range(min(len(op_arr), len(sales_arr)))
        ]

        # ROIC = OP×0.75 / TA （簡易計算）
        roic_arr = [
            op_arr[i] * 0.75 / ta_arr[i] * 100
            if op_arr[i] is not None and ta_arr[i] and ta_arr[i] != 0 else None
            for i in range(min(len(op_arr), len(ta_arr)))
        ]

        def last(arr):
            vals = [v for v in arr if v is not None and not np.isnan(float(v))]
            return round(vals[-1], 2) if vals else None

        return {
            "eps":     eps_arr,
            "gm":  gm_arr,
            "ocf":     cfo_arr,
            "fcf":     fcf_arr,
            "roe": roe_arr,
            "roa": roa_arr,
            "de":  de_arr,
            "roic":roic_arr,
            "latest_EPS":  last(eps_arr),
            "latest_GM":   last(gm_arr),
            "latest_ROE":  last(roe_arr),
            "latest_ROA":  last(roa_arr),
            "latest_DE":   last(de_arr),
            "latest_ROIC": last(roic_arr),
            "source": "jquants",
            "fy_years": len(fy),
        }
    except Exception as e:
        return None



# ── エコノミックモート評価 ────────────────────────────
import json as _moat_json

def load_moat_data():
    """via_moat.json から定性モートデータを読み込む"""
    path = os.path.join(SCRIPT_DIR, "via_moat.json")
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return _moat_json.load(f)
        except Exception:
            pass
    return {}

_MOAT_DATA = {}   # 起動時に読み込む

def calc_moat_score(row_data, moat_db):
    """
    定量スコア（70点）＋定性ボーナス（30点）でモートスコアを計算。
    row_data: dict（latest_ROE, latest_ROA, latest_ROIC, latest_GM, dcf_raw_rateなど）
    moat_db: via_moat.jsonの内容
    戻り値: dict {score, quant, bonus, grade, types, comment, risk}
    """
    def to_f(v, default=0):
        try: return float(v) if v not in (None,"","nan","None") else default
        except: return default

    roe    = to_f(row_data.get("latest_ROE"))
    roa    = to_f(row_data.get("latest_ROA"))
    roic   = to_f(row_data.get("latest_ROIC"))
    gm     = to_f(row_data.get("latest_GM"))
    growth = to_f(row_data.get("dcf_raw_rate"))

    # 定量スコア（70点満点）
    q_roe    = min(roe   / 20  * 20, 20) if roe   > 0 else 0
    q_roa    = min(roa   / 10  * 15, 15) if roa   > 0 else 0
    q_roic   = min(roic  / 15  * 15, 15) if roic  > 0 else 0
    q_gm     = min(gm    / 40  * 15, 15) if gm    > 0 else 0
    q_growth = min(growth/ 15  *  5,  5) if growth > 0 else 0
    quant = round(q_roe + q_roa + q_roic + q_gm + q_growth, 1)

    # 定性ボーナス（via_moat.jsonから）
    ticker = row_data.get("ticker","")
    moat_info = moat_db.get(ticker, {})
    bonus   = moat_info.get("bonus", 0)
    grade   = moat_info.get("moat_grade", "")
    types   = moat_info.get("moat_type", [])
    comment = moat_info.get("comment", "")
    risk    = moat_info.get("risk", "")

    total = round(min(quant + bonus, 100), 1)

    # gradeが未設定の場合は定量スコアから自動判定
    if not grade:
        if total >= 75: grade = "wide"
        elif total >= 55: grade = "narrow"
        else: grade = "none"

    return {
        "moat_score": total,
        "moat_quant": quant,
        "moat_bonus": bonus,
        "moat_grade": grade,
        "moat_types": types,
        "moat_comment": comment,
        "moat_risk": risk,
    }

def _jquants_fetch(path, params=""):
    """J-Quants API V2 リクエスト"""
    import ssl as _ssl, json as _json
    url = f"{JQUANTS_BASE}{path}?{params}" if params else f"{JQUANTS_BASE}{path}"
    req = urllib.request.Request(url, headers={"x-api-key": JQUANTS_API_KEY})
    ctx = _ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode    = _ssl.CERT_NONE
    with urllib.request.urlopen(req, timeout=15, context=ctx) as r:
        return _json.loads(r.read())

_jq_cache = {}   # J-Quants財務データキャッシュ

def _eodhd_fetch(path, params=""):
    """EODHD API リクエスト"""
    import ssl as _ssl, json as _json
    url = (f"{EODHD_BASE}/{path}?api_token={EODHD_TOKEN}&fmt=json&{params}"
           if params else
           f"{EODHD_BASE}/{path}?api_token={EODHD_TOKEN}&fmt=json")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    ctx = _ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode    = _ssl.CERT_NONE
    with urllib.request.urlopen(req, timeout=20, context=ctx) as r:
        return _json.loads(r.read())

_eodhd_cache = {}   # EODHDメモリキャッシュ

def _cache_path(ticker, source="eodhd"):
    """キャッシュファイルパスを返す"""
    os.makedirs(CACHE_DIR, exist_ok=True)
    safe = ticker.replace(".", "_").replace("/", "_")
    return os.path.join(CACHE_DIR, f"{source}_{safe}.json")

# ── 株価キャッシュ（日次・全銘柄一括）──
PRICE_CACHE_FILE = os.path.join(SCRIPT_DIR, "via_cache", "prices.json")

def load_price_cache():
    """株価キャッシュを読み込む。{ticker: {"price": x, "ts": timestamp}}"""
    import json as _json
    if os.path.exists(PRICE_CACHE_FILE):
        try:
            with open(PRICE_CACHE_FILE, "r", encoding="utf-8") as f:
                return _json.load(f)
        except Exception:
            pass
    return {}

def save_price_cache(cache):
    """株価キャッシュを保存する"""
    import json as _json
    os.makedirs(os.path.dirname(PRICE_CACHE_FILE), exist_ok=True)
    with open(PRICE_CACHE_FILE, "w", encoding="utf-8") as f:
        _json.dump(cache, f, ensure_ascii=False)

def get_cached_price(ticker, price_cache):
    """
    キャッシュから株価を取得。
    PRICE_CACHE_HOURS以内なら有効。なければNoneを返す。
    """
    entry = price_cache.get(ticker)
    if not entry:
        return None
    age_hours = (time.time() - entry.get("ts", 0)) / 3600
    if age_hours > PRICE_CACHE_HOURS:
        return None
    return entry.get("price")

def _load_cache(ticker, source="eodhd"):
    """
    キャッシュから財務データを読み込む。
    FINS_CACHE_DAYS以内なら有効。
    戻り値: (data, is_fresh) または (None, False)
    """
    import json as _json
    path = _cache_path(ticker, source)
    if not os.path.exists(path):
        return None, False
    age_days = (time.time() - os.path.getmtime(path)) / 86400
    if age_days > FINS_CACHE_DAYS:
        return None, False
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = _json.load(f)
        return data, True
    except Exception:
        return None, False

def _save_cache(ticker, data, source="eodhd"):
    """財務データをキャッシュファイルに保存"""
    import json as _json
    path = _cache_path(ticker, source)
    with open(path, "w", encoding="utf-8") as f:
        _json.dump(data, f, ensure_ascii=False)

def _has_recent_earnings(ticker_us, days=90):
    """
    EODHDのEarnings::Historyで直近N日以内に決算発表があったか確認。
    True → 財務データを再取得すべき
    """
    import json as _json
    from datetime import datetime as _dt, timedelta as _td
    # 決算チェックキャッシュ
    path = _cache_path(ticker_us, "earnings_check")
    if os.path.exists(path):
        age_days = (time.time() - os.path.getmtime(path)) / 86400
        if age_days <= EARNINGS_CACHE_DAYS:
            with open(path, "r") as f:
                return _json.load(f).get("has_recent", False)
    try:
        eh = _eodhd_fetch(f"fundamentals/{ticker_us}", "filter=Earnings::History")
        cutoff = (_dt.now() - _td(days=days)).strftime("%Y-%m-%d")
        has_recent = any(
            d.get("reportDate","") >= cutoff
            for d in (eh.values() if isinstance(eh, dict) else [])
            if d.get("epsActual") is not None
        )
        # 結果をキャッシュ
        _save_cache(ticker_us, {"has_recent": has_recent}, "earnings_check")
        return has_recent
    except Exception:
        return False  # エラー時はスキップ（再取得しない）

def _has_recent_jquants_earnings(code4, days=90):
    """
    J-Quantsの直近DiscDateを確認して、N日以内に決算発表があったか判定。
    True → 財務データを再取得すべき
    """
    import json as _json
    from datetime import datetime as _dt, timedelta as _td

    # 決算チェックキャッシュ確認
    path = _cache_path(code4, "jq_earnings_check")
    if os.path.exists(path):
        age_days = (time.time() - os.path.getmtime(path)) / 86400
        if age_days <= EARNINGS_CACHE_DAYS:
            try:
                with open(path, "r") as f:
                    return _json.load(f).get("has_recent", False)
            except Exception:
                pass

    try:
        data = _jquants_fetch("/fins/summary", f"code={code4}")
        stmts = data.get("data", [])

        cutoff = (_dt.now() - _td(days=days)).strftime("%Y-%m-%d")
        # FY決算（年次）の最新DiscDateを確認
        fy_dates = [
            s.get("DiscDate","")
            for s in stmts
            if "FYFinancialStatements" in s.get("DocType","")
            and s.get("DiscDate","") >= cutoff
        ]
        has_recent = len(fy_dates) > 0

        # 結果をキャッシュ保存
        os.makedirs(CACHE_DIR, exist_ok=True)
        with open(path, "w") as f:
            _json.dump({"has_recent": has_recent,
                        "checked": _dt.now().strftime("%Y-%m-%d")}, f)
        return has_recent

    except Exception:
        return False  # エラー時は再取得しない


def get_us_financials(ticker, force_refresh=False):
    """
    EODHD から米国株の財務データを取得。
    キャッシュ（30日有効）を優先使用。
    force_refresh=True: 直近に決算発表があった場合に強制再取得。
    ticker: "FHI" などのティッカー（.US付きでも可）
    """
    code = ticker.replace(".US","").upper()
    if code in _eodhd_cache:
        return _eodhd_cache[code]

    empty = {k: [] for k in ["eps","gm","ni","rev","gp","eq","ta","td",
                               "ic","ocf","fcf","roe","roa","de","roic"]}

    # ── ファイルキャッシュ確認 ──
    cached, is_fresh = _load_cache(code, "eodhd")
    if is_fresh and cached and not force_refresh:
        _eodhd_cache[code] = cached
        return cached

    # キャッシュ期限切れ or なし → 決算発表チェック
    ticker_us = f"{code}.US"
    if cached and not force_refresh:
        # キャッシュあり（期限切れ）→ 直近に決算発表があった場合のみ再取得
        if not _has_recent_earnings(ticker_us, days=90):
            # 決算発表なし → 古いキャッシュをそのまま使用
            _eodhd_cache[code] = cached
            return cached

    try:

        # ── EPS: Earnings::Annual ──
        ea = _eodhd_fetch(f"fundamentals/{ticker_us}",
                          "filter=Earnings::Annual")
        eps_raw = {}
        if isinstance(ea, dict):
            from datetime import datetime as _dt
            current_year = _dt.now().year
            for date, d in ea.items():
                y = int(date[:4])
                # 現在年のデータは四半期データの可能性が高いためスキップ
                if y >= current_year:
                    continue
                v = d.get("epsActual")
                if v is not None and str(v) not in ("null","None"):
                    try:
                        eps_raw[str(y)] = float(v)
                    except:
                        pass
        eps_l = [eps_raw[y] for y in sorted(eps_raw.keys())]

        # ── 損益計算書: Financials::Income_Statement::yearly ──
        inc = _eodhd_fetch(f"fundamentals/{ticker_us}",
                           "filter=Financials::Income_Statement::yearly")
        inc_years = sorted(inc.keys()) if isinstance(inc, dict) else []

        def to_f(val):
            try:
                v = float(val)
                return v if v != 0 else None
            except:
                return None

        rev_l = [to_f(inc[y].get("totalRevenue"))   for y in inc_years]
        gp_l  = [to_f(inc[y].get("grossProfit"))    for y in inc_years]
        ni_l  = [to_f(inc[y].get("netIncome"))      for y in inc_years]

        # 粗利益率
        gm_l = [
            gp_l[i] / rev_l[i] * 100
            if gp_l[i] is not None and rev_l[i] and rev_l[i] != 0
            else None
            for i in range(min(len(rev_l), len(gp_l)))
        ]

        # ── バランスシート ──
        bs = _eodhd_fetch(f"fundamentals/{ticker_us}",
                          "filter=Financials::Balance_Sheet::yearly")
        bs_years = sorted(bs.keys()) if isinstance(bs, dict) else []

        eq_l  = [to_f(bs[y].get("totalStockholderEquity")) for y in bs_years]
        ta_l  = [to_f(bs[y].get("totalAssets"))            for y in bs_years]
        td_l  = [to_f(bs[y].get("shortLongTermDebt")
                       or bs[y].get("longTermDebt"))        for y in bs_years]
        ic_l  = [to_f(bs[y].get("investedCapital")
                       or bs[y].get("totalCapitalization")) for y in bs_years]

        # ── キャッシュフロー ──
        cf = _eodhd_fetch(f"fundamentals/{ticker_us}",
                          "filter=Financials::Cash_Flow::yearly")
        cf_years = sorted(cf.keys()) if isinstance(cf, dict) else []

        ocf_l = [to_f(cf[y].get("totalCashFromOperatingActivities")) for y in cf_years]
        fcf_l = [to_f(cf[y].get("freeCashFlow")) for y in cf_years]

        # ── 派生指標（年数を合わせて計算）──
        n = min(len(ni_l), len(eq_l), len(ta_l))
        roe_l = [
            round(ni_l[i] / eq_l[i] * 100, 2)
            if ni_l[i] is not None and eq_l[i] and eq_l[i] != 0 else None
            for i in range(n)
        ]
        roa_l = [
            round(ni_l[i] / ta_l[i] * 100, 2)
            if ni_l[i] is not None and ta_l[i] and ta_l[i] != 0 else None
            for i in range(n)
        ]
        de_l = [
            round(abs(td_l[i] / eq_l[i]) * 100, 2)
            if i < len(td_l) and td_l[i] is not None
               and eq_l[i] and eq_l[i] != 0 else None
            for i in range(n)
        ]
        roic_l = [
            round(ni_l[i] / ic_l[i] * 100, 2)
            if i < len(ic_l) and ic_l[i] is not None
               and ni_l[i] is not None and ic_l[i] != 0 else None
            for i in range(min(n, len(ic_l)))
        ]

        result = {
            "eps": eps_l, "gm": gm_l, "ni": ni_l, "rev": rev_l, "gp": gp_l,
            "eq": eq_l, "ta": ta_l, "td": td_l, "ic": ic_l,
            "ocf": ocf_l, "fcf": fcf_l,
            "roe": roe_l, "roa": roa_l, "de": de_l, "roic": roic_l,
        }
        # ── EPS通貨ミスマッチ検出・USD換算 ──
        # yfinanceのtrailingEpsと比較してEPSが現地通貨建ての場合は換算
        try:
            import yfinance as _yf3
            _tk3 = _yf3.Ticker(code)
            _ttm = _tk3.info.get("trailingEps")
            _price3 = _tk3.fast_info.last_price
            if _ttm and _ttm > 0 and _price3 and _price3 > 0 and result["eps"]:
                _latest_eps = [v for v in result["eps"] if v is not None]
                if _latest_eps:
                    _pe_eodhd = _price3 / _latest_eps[-1]
                    # P/Eが異常に低い（<1）または高い（>500）場合は通貨ミスマッチ
                    if _pe_eodhd < 1.0 or _pe_eodhd > 500:
                        _fi3 = _tk3.income_stmt
                        _eps_local = _gs(_fi3, ["Basic EPS","Diluted EPS","EPS"])                                      if _fi3 is not None and not _fi3.empty else []
                        if _eps_local and _eps_local[-1] and _eps_local[-1] != 0:
                            _fx = _ttm / _eps_local[-1]
                            _eps_usd = [round(e * _fx, 4) for e in _eps_local if e is not None]
                            result["eps"] = _eps_usd
        except Exception:
            pass  # 換算失敗時はEODHDのEPSをそのまま使用

        # ファイルキャッシュに保存
        _save_cache(code, result, "eodhd")
        _eodhd_cache[code] = result
        return result

    except Exception:
        # 失敗時は古いキャッシュがあればそれを使用
        if cached:
            _eodhd_cache[code] = cached
            return cached
        _eodhd_cache[code] = empty
        return empty



def get_jp_financials(code4):
    """
    J-Quants /fins/summary から年次決算データを取得。
    code4: 4桁コード（例: "7792"）
    """
    if code4 in _jq_cache:
        return _jq_cache[code4]

    empty = {k: [] for k in ["eps","gm","np","op","eq","ta","eqar",
                               "cfo","cfi","fcf","de","roe","roa","roic",
                               "ocf"]}  # ocfはcfoの別名（後方互換）
    empty.update({"latest_EPS": None, "latest_GM": None,
                  "latest_ROE": None, "latest_ROA": None,
                  "latest_DE": None, "latest_ROIC": None})

    # ファイルキャッシュ確認
    cached_jq, is_fresh_jq = _load_cache(code4, "jquants")
    if is_fresh_jq and cached_jq:
        # キャッシュ有効期限内 → そのまま使用
        _jq_cache[code4] = cached_jq
        return cached_jq

    if cached_jq:
        # キャッシュ期限切れ → 直近90日に年次決算発表があった場合のみ再取得
        if not _has_recent_jquants_earnings(code4, days=90):
            # 決算発表なし → 古いキャッシュを継続使用
            _jq_cache[code4] = cached_jq
            return cached_jq
        # 決算発表あり → 再取得へ続行

    try:
        data = _jquants_fetch("/fins/summary", f"code={code4}")
        stmts = data.get("data", [])
        fy_all = [s for s in stmts if "FYFinancialStatements" in s.get("DocType","")]

        # 同一決算期（CurFYEn）に複数の開示がある場合（訂正報告書等）は
        # 最新のDiscDateのもののみを残す（重複除去）
        fy_by_period = {}
        for s in fy_all:
            period = s.get("CurFYEn", "")
            disc_date = s.get("DiscDate", "")
            if period not in fy_by_period or disc_date > fy_by_period[period].get("DiscDate",""):
                fy_by_period[period] = s
        fy = sorted(fy_by_period.values(), key=lambda x: x.get("DiscDate",""))

        if not fy:
            _jq_cache[code4] = empty
            return empty

        def to_f(val):
            try:
                v = float(val)
                return v if v != 0 else None
            except (TypeError, ValueError):
                return None

        eps_l  = [to_f(s.get("EPS"))   for s in fy]
        np_l   = [to_f(s.get("NP"))    for s in fy]
        op_l   = [to_f(s.get("OP"))    for s in fy]
        eq_l   = [to_f(s.get("Eq"))    for s in fy]
        ta_l   = [to_f(s.get("TA"))    for s in fy]
        eqar_l = [to_f(s.get("EqAR")) for s in fy]
        cfo_l  = [to_f(s.get("CFO"))  for s in fy]
        cfi_l  = [to_f(s.get("CFI"))  for s in fy]

        fcf_l = [
            (cfo_l[i] + cfi_l[i])
            if cfo_l[i] is not None and cfi_l[i] is not None else None
            for i in range(len(cfo_l))
        ]
        # D/E: EqARから全負債比率で計算
        de_l = [
            round((1/v - 1) * 100, 1) if v and 0 < v <= 1 else None
            for v in eqar_l
        ]
        roe_l = [
            round(np_l[i] / eq_l[i] * 100, 2)
            if np_l[i] is not None and eq_l[i] and eq_l[i] != 0 else None
            for i in range(min(len(np_l), len(eq_l)))
        ]
        roa_l = [
            round(np_l[i] / ta_l[i] * 100, 2)
            if np_l[i] is not None and ta_l[i] and ta_l[i] != 0 else None
            for i in range(min(len(np_l), len(ta_l)))
        ]

        def _last(arr):
            vals = [v for v in arr if v is not None]
            try:
                import numpy as _np
                vals = [v for v in vals if not _np.isnan(float(v))]
            except Exception:
                pass
            return round(vals[-1], 2) if vals else None

        result = {
            "eps": eps_l, "gm": [],   # 粗利益率はJ-Quantsでは計算不可
            "np": np_l, "op": op_l, "eq": eq_l, "ta": ta_l,
            "eqar": eqar_l, "cfo": cfo_l, "cfi": cfi_l,
            "ocf": cfo_l,              # ocfはcfoの別名（後方互換）
            "fcf": fcf_l, "de": de_l, "roe": roe_l,
            "roa": roa_l, "roic": roa_l,  # ROICはROAで代替
            "latest_EPS":  _last(eps_l),
            "latest_GM":   None,
            "latest_ROE":  _last(roe_l),
            "latest_ROA":  _last(roa_l),
            "latest_DE":   _last(de_l),
            "latest_ROIC": _last(roa_l),
        }
        # ROIC: yfinanceのInvested Capitalで計算（失敗時はROAで代替）
        try:
            import yfinance as _yf2
            _tk2 = _yf2.Ticker(f"{code4}.T")
            _bs2 = _tk2.balance_sheet
            _fi2 = _tk2.income_stmt
            if _bs2 is not None and not _bs2.empty and "Invested Capital" in _bs2.index:
                _ic2 = _bs2.loc["Invested Capital"].dropna().sort_index(ascending=True).values.tolist()
                _ni2 = _fi2.loc["Net Income"].dropna().sort_index(ascending=True).values.tolist() \
                       if _fi2 is not None and "Net Income" in _fi2.index else []
                _n2  = min(len(_ni2), len(_ic2))
                _roic2 = [
                    round(_ni2[i] / _ic2[i] * 100, 2)
                    if _ic2[i] and _ic2[i] != 0 else None
                    for i in range(_n2)
                ]
                if _roic2:
                    result["roic"] = _roic2
                    result["latest_ROIC"] = _last(_roic2)
        except Exception:
            pass  # 失敗時はROAで代替（result["roic"]=roa_lのまま）
        # GM: yfinanceのGross Profit / Total Revenueで計算
        try:
            import yfinance as _yf3
            _tk3 = _yf3.Ticker(f"{code4}.T")
            _fi3 = _tk3.income_stmt
            if _fi3 is not None and not _fi3.empty:
                _gp3 = _fi3.loc["Gross Profit"].dropna().sort_index(ascending=True).values.tolist() \
                       if "Gross Profit" in _fi3.index else []
                _rv3 = _fi3.loc["Total Revenue"].dropna().sort_index(ascending=True).values.tolist() \
                       if "Total Revenue" in _fi3.index else []
                _n3  = min(len(_gp3), len(_rv3))
                _gm3 = [
                    round(_gp3[i] / _rv3[i] * 100, 2)
                    if _rv3[i] and _rv3[i] != 0 else None
                    for i in range(_n3)
                ]
                if _gm3:
                    result["gm"] = _gm3
                    result["latest_GM"] = _last(_gm3)
        except Exception:
            pass  # 失敗時は空リスト（GMスキップ）のまま
        # ファイルキャッシュに保存
        _save_cache(code4, result, "jquants")
        _jq_cache[code4] = result
        return result

    except Exception:
        cached_jq, _ = _load_cache(code4, "jquants")
        if cached_jq:
            _jq_cache[code4] = cached_jq
            return cached_jq
        _jq_cache[code4] = empty
        return empty


def calc_asset_undervaluation(ticker, market, market_cap):
    """
    VIA通過銘柄（日本株）について、資産面の過小評価要因を計算する。
    NCAV（流動資産-負債総額）+ 投資有価証券（時価）- 税効果額（評価差額金から逆算）
    を時価総額と比較し、過小評価スコアを返す。

    米国株、または_edinetモジュールが無い場合はNoneを返す（スキップ）。
    """
    if market != "JP" or _edinet is None or not market_cap:
        return None
    try:
        tk = yf.Ticker(ticker)
        bs = tk.balance_sheet
        if bs is None or bs.empty:
            return None

        def _latest(keys):
            for k in keys:
                if k in bs.index:
                    vals = bs.loc[k].dropna()
                    if not vals.empty:
                        return float(vals.iloc[0])
            return None

        current_assets = _latest(["Current Assets"])
        total_liab      = _latest(["Total Liabilities Net Minority Interest", "Total Liab"])
        investments     = _latest(["Investmentin Financial Assets",
                                     "Long Term Equity Investment",
                                     "Investments And Advances"]) or 0.0
        # 非支配株主持分（連結子会社の少数株主分）。
        # 親会社株主に帰属しない清算価値のため、NCAVから除外する。
        minority_interest = _latest(["Minority Interest"]) or 0.0

        if current_assets is None or total_liab is None:
            return None

        # 棚卸資産劣化リスク用データ（在庫・売上高の時系列）
        inventory_series = None
        revenue_series = None
        try:
            if "Inventory" in bs.index:
                inventory_series = bs.loc["Inventory"].dropna().tolist()
            fi = tk.income_stmt
            if fi is not None and "Total Revenue" in fi.index:
                revenue_series = fi.loc["Total Revenue"].dropna().tolist()
        except Exception:
            pass

        # 決算日ヒント取得（EDINET検索の高速化用）
        fiscal_year_end = None
        try:
            info = tk.info
            fy_end_ts = info.get("lastFiscalYearEnd")
            if fy_end_ts:
                fy_dt = datetime.fromtimestamp(fy_end_ts)
                fiscal_year_end = fy_dt.strftime("%m-%d")
        except Exception:
            pass

        mapping = _edinet.load_edinet_mapping()
        result = _edinet.calc_asset_undervaluation_score(
            ticker, market_cap, current_assets, total_liab,
            investments=investments, mapping=mapping,
            fiscal_year_end=fiscal_year_end,
            inventory_series=inventory_series, revenue_series=revenue_series,
            minority_interest=minority_interest,
        )
        return result
    except Exception:
        return None


def get_tickers():
    """銘柄リストを全て取得してから返す"""
    print("[銘柄リスト取得]")
    print("-" * 40)

    jp = load_jp_tickers()
    print()
    us = load_us_tickers()

    us = list(dict.fromkeys(us))
    jp = list(dict.fromkeys(jp))

    print()
    print("-" * 40)
    print(f"  取得完了: US {len(us)} 銘柄 + JP {len(jp)} 銘柄 = {len(us)+len(jp)} 銘柄")
    print("-" * 40)
    return us, jp


# ── 判定ロジック ───────────────────────────────────────

def is_uptrend(arr, consec_limit=2):
    vals = [v for v in arr if v is not None and not np.isnan(float(v))]
    if len(vals) < 2: return False
    consec = 0
    for i in range(len(vals) - 1):
        if vals[i+1] < vals[i]:
            consec += 1
            if consec >= consec_limit: return False
        else:
            consec = 0
    return True


def is_eps_uptrend(arr, required_years):
    """
    EPS専用の上昇トレンド判定（is_uptrendとは判定基準が異なる）。

    直近required_years年分（厳格=5年、緩和=4年）のEPSのみを取り出し、
    その固定期間内で判定する（絶対条件）。

    判定基準:
      下落0回 → 合格
      下落1回かつ単年度（2年連続でない）→ 例外として合格（一時的要因と判断）
      下落2回以上 → 不合格
      2年連続の下落が1度でもあれば → 不合格（下落回数に関わらず）
    """
    vals = [v for v in arr if v is not None and not np.isnan(float(v))]
    if len(vals) < required_years:
        return False

    recent = vals[-required_years:]

    decline_count = 0
    consec_decline = 0
    max_consec = 0
    for i in range(len(recent) - 1):
        if recent[i+1] < recent[i]:
            decline_count += 1
            consec_decline += 1
            max_consec = max(max_consec, consec_decline)
        else:
            consec_decline = 0

    if max_consec >= 2:
        return False
    if decline_count >= 2:
        return False

    return True

def is_stable(arr):
    vals = [v for v in arr if v is not None and not np.isnan(float(v))]
    if len(vals) < 2: return False
    for i in range(len(vals) - 1):
        if abs(vals[i+1] - vals[i]) > 5: return False
    return True

def pct_above(arr, thr):
    vals = [v for v in arr if v is not None and not np.isnan(float(v))]
    if not vals: return 0.0
    return sum(1 for v in vals if v >= thr) / len(vals)

def pct_positive(arr): return pct_above(arr, 0.0001)

def pct_below(arr, thr):
    vals = [v for v in arr if v is not None and not np.isnan(float(v))]
    if not vals: return 0.0
    return sum(1 for v in vals if v <= thr) / len(vals)

def gm_judge(gm_arr, cfg):
    vals = [v for v in gm_arr if v is not None and not np.isnan(float(v))]
    if not vals: return None, "データなし"
    high40 = sum(1 for v in vals if v >= 40) / len(vals) >= cfg["pass_ratio"]
    if high40:
        ok = is_uptrend(gm_arr, cfg["consec_drop"])
        return ok, f"40%以上({'右肩上がり' if ok else '右肩上がりNG'})"
    else:
        ok = is_stable(gm_arr) or is_uptrend(gm_arr, cfg["consec_drop"])
        avg = np.nanmean(vals)
        return ok, f"40%未満(avg{avg:.1f}%)→{'OK' if ok else 'NG'}"

def run_criteria(eps, gm_arr, ocf, fcf, roe_arr, roa_arr,
                 de_arr, roic_arr, cfg):
    pr = cfg["pass_ratio"]
    cd = cfg["consec_drop"]

    # 年数スライス処理
    ry  = cfg.get("recent_years")   # 緩和: 直近N年で評価
    mn  = cfg.get("min_years")      # 厳格: 最低N年必要
    mx  = cfg.get("max_years")      # 厳格: 最大N年で評価

    if ry:
        # 緩和基準: 直近recent_years年のみで評価
        def sl(arr): return arr[-ry:] if arr and len(arr) > ry else arr
    elif mx:
        # 厳格基準: 直近max_years年に絞る（min_years未満はgm_judgeなどで自然にFalseになる）
        def sl(arr):
            if not arr: return arr
            return arr[-mx:]    # 直近max_years年に絞るだけ
    else:
        def sl(arr): return arr

    eps     = sl(eps)
    gm_arr  = sl(gm_arr)
    ocf     = sl(ocf)
    fcf     = sl(fcf)
    roe_arr = sl(roe_arr)
    roa_arr = sl(roa_arr)
    de_arr  = sl(de_arr)
    roic_arr= sl(roic_arr)

    gm_ok, gm_note = gm_judge(gm_arr, cfg)
    res = {
        "EPS_positive":   pct_positive(eps)              >= pr if eps      else False,
        "EPS_uptrend":    is_eps_uptrend(eps, cfg.get("min_years", 4)) if eps else False,
        "GM_judge":       gm_ok,
        "OCF_positive":   pct_positive(ocf)              >= pr if ocf      else False,
        "OCF_uptrend":    is_uptrend(ocf, cd)                   if ocf      else False,
        "FCF_positive":   pct_positive(fcf)              >= pr if fcf      else False,
        "FCF_uptrend":    is_uptrend(fcf, cd)                   if fcf      else False,
        "ROE_above_thr":  pct_above(roe_arr,cfg["roe_thr"]) >= pr if roe_arr else False,
        "ROE_uptrend":    is_uptrend(roe_arr, cd)               if roe_arr  else False,
        "ROA_above_thr":  pct_above(roa_arr,cfg["roa_thr"]) >= pr if roa_arr else False,
        "ROA_uptrend":    is_uptrend(roa_arr, cd)               if roa_arr  else False,
        "DE_below_thr":   pct_below(de_arr,cfg["de_thr"])   >= pr if de_arr  else False,
        "ROIC_above_thr": pct_above(roic_arr,cfg["roic_thr"])>= pr if roic_arr else False,
        "ROIC_vs_WACC":   (pct_above(roic_arr, WACC) >= pr
                           if roic_arr else None),
    }
    score_items = {k: v for k, v in res.items() if v is not None}
    passed = sum(1 for v in score_items.values() if v is True)
    total  = len(score_items)
    return res, passed, total, gm_note


# ── DCF ───────────────────────────────────────────────

def calc_dcf(eps_series):
    vals = [v for v in eps_series if v is not None and not np.isnan(float(v))]
    if len(vals) < 2: return None
    n = len(eps_series)
    present_eps = eps_series[n - 1]
    if not present_eps or np.isnan(float(present_eps)) or present_eps <= 0:
        return None
    sel_idx = -1
    for back in range(9, 4, -1):
        idx = n - 1 - back
        if idx < 0: continue
        v = eps_series[idx]
        if not v or np.isnan(float(v)) or v <= 0: continue
        if is_uptrend(eps_series[idx:], 2):
            sel_idx = idx; break
    if sel_idx == -1:
        for back in range(min(9, n-1), 0, -1):
            idx = n - 1 - back
            if idx < 0: continue
            v = eps_series[idx]
            if v and not np.isnan(float(v)) and v > 0:
                sel_idx = idx; break
    if sel_idx == -1: return None
    past_eps   = eps_series[sel_idx]
    years_back = (n - 1) - sel_idx
    try:
        raw_rate = (pow(present_eps / past_eps, 1.0 / years_back) - 1) * 100
    except Exception:
        return None
    adj_rate = min(max(raw_rate / 2.0, 1.0), DCF_ADJ_CAP)
    dr = DCF_DISC_RATE/100; ar = adj_rate/100
    ir = DCF_INFL_RATE/100; cy = DCF_CONT_YEARS; sy = DCF_SURV_YEARS
    g0 = (1+ar)/(1+dr)
    gv = (present_eps*cy if abs(g0-1)<1e-9
          else present_eps*g0*(1-pow(g0,cy))/(1-g0))
    s0 = (1+ir)/(1+dr); s2 = pow(g0,cy)
    sv = (present_eps*s2*sy if abs(s0-1)<1e-9
          else present_eps*s2*s0*(1-pow(s0,sy))/(1-s0))
    intrinsic = gv + sv
    return {
        "dcf_buy_price":   round(intrinsic*(1-DCF_MARGIN_SAFE/100), 2),
        "dcf_intrinsic":   round(intrinsic, 2),
        "dcf_adj_rate":    round(adj_rate,  2),
        "dcf_raw_rate":    round(raw_rate,  2),
        "dcf_years_back":  years_back,
        "dcf_past_eps":    round(past_eps,  2),
        "dcf_present_eps": round(present_eps, 2),
    }


# ── 銘柄評価 ──────────────────────────────────────────

# ── J-Quants API統合（日本株財務データ） ──────────────────
JQUANTS_API_KEY = "TMhn3YeV3FTDnbCaW6pm6irEqvbocWKoooVYGySQPFI"
JQUANTS_BASE    = "https://api.jquants.com/v2"
_jq_cache = {}   # {code: [statements]}

def _jq_fetch(path, params=""):
    import urllib.request, ssl as _ssl
    url = f"{JQUANTS_BASE}{path}?{params}" if params else f"{JQUANTS_BASE}{path}"
    req = urllib.request.Request(url, headers={"x-api-key": JQUANTS_API_KEY})
    ctx = _ssl.create_default_context()
    ctx.check_hostname = False; ctx.verify_mode = _ssl.CERT_NONE
    with urllib.request.urlopen(req, timeout=15, context=ctx) as r:
        import json as _json
        return _json.loads(r.read())

def get_jquants_financials(code_with_t):
    """
    J-Quantsから年次FY決算データを取得して財務配列を返す。
    code_with_t: "7792.T" 形式 → "7792" に変換
    戻り値: dict with keys: eps, sales, cfo, fcf, roe, roa, ni, eq, ta
    """
    code = code_with_t.replace(".T","").replace(".","")
    if code in _jq_cache:
        return _jq_cache[code]

    try:
        data  = _jq_fetch("/fins/summary", f"code={code}")
        stmts = data.get("data", [])

        # 年次FY決算のみ抽出（DocTypeにFYFinancialStatementsを含む）
        fy = [s for s in stmts if "FYFinancial" in s.get("DocType","")]
        if not fy:
            # FYがなければ全件（一部銘柄は形式が異なる）
            fy = [s for s in stmts if s.get("CurPerType","") in ("FY","Annual","通期")]

        if not fy:
            _jq_cache[code] = None
            return None

        # 古い順にソート
        fy = sorted(fy, key=lambda x: x.get("DiscDate",""))

        def to_f(v):
            try: return float(v) if v not in (None,"","null") else None
            except: return None

        eps_arr  = [to_f(s.get("EPS"))  for s in fy]
        sales_arr= [to_f(s.get("Sales"))for s in fy]
        np_arr   = [to_f(s.get("NP"))   for s in fy]
        eq_arr   = [to_f(s.get("Eq"))   for s in fy]
        ta_arr   = [to_f(s.get("TA"))   for s in fy]
        cfo_arr  = [to_f(s.get("CFO"))  for s in fy]
        cfi_arr  = [to_f(s.get("CFI"))  for s in fy]

        # FCF = CFO + CFI
        fcf_arr = [
            cfo_arr[i] + cfi_arr[i]
            if cfo_arr[i] is not None and cfi_arr[i] is not None else None
            for i in range(len(cfo_arr))
        ]

        # ROE = NP / Eq * 100
        roe_arr = [
            np_arr[i] / eq_arr[i] * 100
            if np_arr[i] is not None and eq_arr[i] and eq_arr[i] != 0 else None
            for i in range(min(len(np_arr), len(eq_arr)))
        ]

        # ROA = NP / TA * 100
        roa_arr = [
            np_arr[i] / ta_arr[i] * 100
            if np_arr[i] is not None and ta_arr[i] and ta_arr[i] != 0 else None
            for i in range(min(len(np_arr), len(ta_arr)))
        ]

        result = {
            "eps":   eps_arr,
            "sales": sales_arr,
            "ni":    np_arr,
            "eq":    eq_arr,
            "ta":    ta_arr,
            "cfo":   cfo_arr,
            "fcf":   fcf_arr,
            "roe":   roe_arr,
            "roa":   roa_arr,
            "years": len(fy),
        }
        _jq_cache[code] = result
        return result
    except Exception as e:
        _jq_cache[code] = None
        return None


def _gs(df, keys):
    for k in keys:
        if k in df.index:
            return df.loc[k].dropna().sort_index().values.tolist()
    return []

def process_ticker(ticker, market):
    try:
        # ── 日本株はJ-Quantsを優先使用 ──
        if market == "JP" and JQUANTS_API_KEY:
            code4 = ticker.replace(".T", "")
            def _jq_last(arr):
                vals = [v for v in (arr or []) if v is not None]
                try:
                    import numpy as _np2
                    vals = [v for v in vals if not _np2.isnan(float(v))]
                except Exception: pass
                return round(vals[-1], 2) if vals else None
            jq = get_jp_financials(code4)
            if jq:
                # J-Quantsデータで評価
                gm_arr  = jq.get("gm", [])
                eps     = jq.get("eps", [])
                ocf     = jq.get("cfo") or jq.get("ocf", [])
                fcf     = jq.get("fcf", [])
                roe_arr = jq.get("roe", [])
                roa_arr = jq.get("roa", [])
                de_arr  = jq.get("de", [])
                roic_arr= jq.get("roic", [])

                # 株価・名称はyfinanceから取得
                tk = yf.Ticker(ticker)
                price = None
                try: price = round(tk.fast_info.last_price, 2)
                except: pass
                name, sector = ticker, ""
                try:
                    fi_name = getattr(tk.fast_info, "name", None)
                    if fi_name and fi_name != ticker:
                        name = fi_name
                    else:
                        info = tk.info
                        name = info.get('longName') or info.get('shortName') or ticker
                    sector = getattr(tk.fast_info, "sector", None) or tk.info.get('sector', "")
                except: pass
                # 名称クリーニング
                if ',' in name:
                    name = name.split(',')[0].strip()
                import re as _re
                if name == ticker or _re.match(r'^[\d\.]+$', name):
                    name = ticker

                # 判定実行
                # J-QuantsはROICをROAで代替するため閾値をROAと揃える
                jp_strict  = dict(STRICT,  roic_thr=STRICT["roa_thr"])
                jp_relaxed = dict(RELAXED, roic_thr=RELAXED["roa_thr"])
                s_res, s_pass, s_total, s_gm = run_criteria(
                    eps, gm_arr, ocf, fcf, roe_arr, roa_arr, de_arr, roic_arr, jp_strict)
                r_res, r_pass, r_total, r_gm = run_criteria(
                    eps, gm_arr, ocf, fcf, roe_arr, roa_arr, de_arr, roic_arr, jp_relaxed)

                s_all = (s_pass == s_total and s_total > 0)
                r_all = (r_pass == r_total and r_total > 0)
                grade = "STRICT" if s_all else ("RELAXED" if r_all else "FAIL")
                relaxed_items = [k for k in s_res if s_res[k] is False and r_res.get(k) is True]

                dcf = calc_dcf(eps) if eps else None
                if dcf and price and price > 0:
                    cur_eps_v = [v for v in eps if v is not None and not np.isnan(float(v))]
                    if cur_eps_v and cur_eps_v[-1] > 0:
                        pe = price / cur_eps_v[-1]
                        if pe < 0.1 or pe > 300:
                            dcf = None
                if dcf and price and price > 0:
                    bp_check = dcf["dcf_buy_price"]
                    ratio = bp_check / price
                    if ratio > 3.5:
                        # EPS通貨ミスマッチ → yfinanceのtrailingEpsで換算レートを推定
                        try:
                            tk_yf = yf.Ticker(ticker)
                            eps_ttm = tk_yf.info.get("trailingEps")
                            if eps_ttm and eps_ttm > 0:
                                fi_yf = tk_yf.income_stmt
                                eps_local = _gs(fi_yf, ["Basic EPS","Diluted EPS","EPS"]) \
                                            if fi_yf is not None and not fi_yf.empty else []
                                if eps_local and eps_local[-1] and eps_local[-1] != 0:
                                    fx = eps_ttm / eps_local[-1]
                                    eps_usd = [round(e * fx, 4) for e in eps_local if e is not None]
                                    dcf = calc_dcf(eps_usd)
                                    if dcf and price > 0:
                                        ratio2 = dcf["dcf_buy_price"] / price
                                        if ratio2 < 0.15 or ratio2 > 3.5:
                                            dcf = None
                                else:
                                    dcf = None
                            else:
                                dcf = None
                        except Exception:
                            dcf = None
                    elif ratio < 0.15:
                        dcf = None
                        dcf = None

                bp = dcf["dcf_buy_price"] if dcf else None
                undervalued = (price < bp) if bp and price and price > 0 else None
                margin_pct  = round((bp-price)/bp*100, 1) if bp and price and price > 0 else None

                return {
                    "ticker": ticker, "name": name,
                    "sector": sector, "market": market, "price": price,
                    "s_score": s_pass, "s_total": s_total,
                    "s_pct":   round(s_pass/s_total*100, 1) if s_total else 0,
                    "r_score": r_pass, "r_total": r_total,
                    "r_pct":   round(r_pass/r_total*100, 1) if r_total else 0,
                    "grade": grade,
                    "relaxed_items": ", ".join(relaxed_items),
                    **{f"s_{k}": v for k, v in s_res.items()},
                    **{f"r_{k}": v for k, v in r_res.items()},
                    "s_gm_note": s_gm,
                    **(dcf if dcf else {
                        "dcf_buy_price": None, "dcf_intrinsic": None,
                        "dcf_adj_rate":  None, "dcf_raw_rate":  None,
                        "dcf_years_back":None, "dcf_past_eps":  None,
                        "dcf_present_eps":None,
                    }),
                    "dcf_skip_reason": None,
                    "undervalued": undervalued,
                    "margin_pct":  margin_pct,
                    "latest_EPS":  _jq_last(eps),
                    "latest_GM":   _jq_last(gm_arr),
                    "latest_ROE":  _jq_last(roe_arr),
                    "latest_ROA":  _jq_last(roa_arr),
                    "latest_DE":   _jq_last(de_arr),
                    "latest_ROIC": _jq_last(roic_arr),
                }

        # ── 米国株: EODHDを使用 ──
        if market == "US" and EODHD_TOKEN:
            eo = get_us_financials(ticker)
            if any(eo.get(k) for k in ["eps","ni","roe"]):
                eps      = eo["eps"]
                gm_arr   = eo["gm"]
                ocf      = eo["ocf"]
                fcf      = eo["fcf"]
                roe_arr  = eo["roe"]
                roa_arr  = eo["roa"]
                de_arr   = eo["de"]
                roic_arr = eo["roic"]

                # 株価・名称はyfinanceから取得
                price = None
                name, sector = ticker, ""
                try:
                    tk = yf.Ticker(ticker)
                    price = round(tk.fast_info.last_price, 2)
                    fi_name = getattr(tk.fast_info, "name", None)
                    name = fi_name if (fi_name and fi_name != ticker) else (
                        tk.info.get("longName") or tk.info.get("shortName") or ticker)
                    if "," in name: name = name.split(",")[0].strip()
                    sector = getattr(tk.fast_info, "sector", None) or tk.info.get("sector","")
                except: pass

                s_res, s_pass, s_total, s_gm = run_criteria(
                    eps, gm_arr, ocf, fcf, roe_arr, roa_arr, de_arr, roic_arr, STRICT)
                r_res, r_pass, r_total, r_gm = run_criteria(
                    eps, gm_arr, ocf, fcf, roe_arr, roa_arr, de_arr, roic_arr, RELAXED)

                s_all = (s_pass == s_total and s_total > 0)
                r_all = (r_pass == r_total and r_total > 0)
                grade = "STRICT" if s_all else ("RELAXED" if r_all else "FAIL")
                relaxed_items = [k for k in s_res
                                 if s_res[k] is False and r_res.get(k) is True]

                dcf = calc_dcf(eps) if eps else None
                if dcf and price and price > 0:
                    cur_v = [v for v in eps if v is not None and not np.isnan(float(v))]
                    if cur_v and cur_v[-1] > 0:
                        pe = price / cur_v[-1]
                        if pe < 0.1 or pe > 300: dcf = None
                if dcf and price and price > 0:
                    ratio = dcf["dcf_buy_price"] / price
                    if ratio > 3.5:
                        # EPS通貨ミスマッチ → yfinanceのtrailingEpsで換算レートを推定
                        try:
                            tk_yf = yf.Ticker(ticker)
                            eps_ttm = tk_yf.info.get("trailingEps")
                            if eps_ttm and eps_ttm > 0:
                                fi_yf = tk_yf.income_stmt
                                eps_local = _gs(fi_yf, ["Basic EPS","Diluted EPS","EPS"])                                             if fi_yf is not None and not fi_yf.empty else []
                                if eps_local and eps_local[-1] and eps_local[-1] != 0:
                                    # 換算レート = trailingEps(USD) / 最新income_stmt EPS(現地通貨)
                                    fx = eps_ttm / eps_local[-1]
                                    # 過去のEPSもUSD換算
                                    eps_usd = [round(e * fx, 4) for e in eps_local if e is not None]
                                    dcf = calc_dcf(eps_usd)
                                    if dcf and price > 0:
                                        ratio2 = dcf["dcf_buy_price"] / price
                                        if ratio2 < 0.15 or ratio2 > 3.5:
                                            dcf = None
                                else:
                                    dcf = None
                            else:
                                dcf = None
                        except Exception:
                            dcf = None
                    elif ratio < 0.15:
                        dcf = None

                bp = dcf["dcf_buy_price"] if dcf else None
                uv = (price < bp) if bp and price and price > 0 else None
                mp = round((bp-price)/bp*100,1) if bp and price and price>0 else None

                def last(arr):
                    v=[x for x in arr if x is not None and not np.isnan(float(x))]
                    return round(v[-1],2) if v else None

                return {
                    "ticker": ticker, "name": name,
                    "sector": sector, "market": market, "price": price,
                    "s_score": s_pass, "s_total": s_total,
                    "s_pct": round(s_pass/s_total*100,1) if s_total else 0,
                    "r_score": r_pass, "r_total": r_total,
                    "r_pct": round(r_pass/r_total*100,1) if r_total else 0,
                    "grade": grade,
                    "relaxed_items": ", ".join(relaxed_items),
                    **{f"s_{k}": v for k, v in s_res.items()},
                    **{f"r_{k}": v for k, v in r_res.items()},
                    "s_gm_note": s_gm,
                    **(dcf if dcf else {
                        "dcf_buy_price": None, "dcf_intrinsic": None,
                        "dcf_adj_rate": None, "dcf_raw_rate": None,
                        "dcf_years_back": None, "dcf_past_eps": None,
                        "dcf_present_eps": None,
                    }),
                    "dcf_skip_reason": None,
                    "undervalued": uv, "margin_pct": mp,
                    "latest_EPS": last(eps),
                    "latest_GM": last(gm_arr),
                    "latest_ROE": last(roe_arr),
                    "latest_ROA": last(roa_arr),
                    "latest_DE": last(de_arr),
                    "latest_ROIC": last(roic_arr),
                    "data_source": "eodhd",
                }

        # ── yfinanceフォールバック（EODHDデータなし / 日本株でJ-Quants失敗）──
        tk = yf.Ticker(ticker)
        fi = tk.income_stmt
        cf = tk.cash_flow
        bs = tk.balance_sheet
        if fi is None or fi.empty: return None

        price = None
        try: price = round(tk.fast_info.last_price, 2)
        except: pass
        name, sector = ticker, ""
        try:
            # fast_info.name を最初に試み、なければ info['longName'] を使う
            fi_name = getattr(tk.fast_info, "name", None)
            if fi_name and fi_name != ticker:
                name = fi_name
            else:
                info = tk.info
                name = (info.get('longName')
                        or info.get('shortName')
                        or ticker)
            sector = (getattr(tk.fast_info, "sector", None)
                      or tk.info.get('sector', ""))
        except:
            pass

        # 名称クリーニング:
        # yfinanceが "5617.T,0P0001RW4J,0" のような複数IDを返す場合は
        # カンマ区切りの最初の要素を使い、それがtickerコードなら不明扱いにする
        if ',' in name:
            name = name.split(',')[0].strip()
        # 名称がtickerそのものや数字コードのみの場合は空にする
        import re as _re
        if name == ticker or _re.match(r'^[\d\.]+$', name):
            name = ticker

        eps   = _gs(fi, ["Basic EPS","Diluted EPS","EPS"])
        rev_s = _gs(fi, ["Total Revenue","Revenue"])
        gp_s  = _gs(fi, ["Gross Profit"])
        ni_s  = _gs(fi, ["Net Income"])
        ocf   = _gs(cf, ["Operating Cash Flow",
                          "Cash Flow From Continuing Operating Activities"])
        fcf   = _gs(cf, ["Free Cash Flow"])
        capex = _gs(cf, ["Capital Expenditure"])
        eq_s  = _gs(bs, ["Stockholders Equity",
                          "Total Equity Gross Minority Interest",
                          "Common Stock Equity"])
        ta_s  = _gs(bs, ["Total Assets"])
        td_s  = _gs(bs, ["Total Debt",
                          "Long Term Debt And Capital Lease Obligation",
                          "Long Term Debt",
                          "Short Long Term Debt"])  # 追加
        ic_s  = _gs(bs, ["Invested Capital","Total Capitalization"])

        # ── 日本株: J-Quantsで財務データを上書き ──
        if market == "JP" and JQUANTS_API_KEY:
            jq = get_jquants_financials(ticker)
            if jq:
                if jq.get("eps"):   eps   = jq["eps"]
                if jq.get("ni"):    ni_s  = jq["ni"]
                if jq.get("eq"):    eq_s  = jq["eq"]
                if jq.get("ta"):    ta_s  = jq["ta"]
                if jq.get("cfo"):   ocf   = jq["cfo"]
                if jq.get("fcf"):   fcf   = jq["fcf"]
                if jq.get("roe"):   pass  # roe_arrは後で計算
                if jq.get("roa"):   pass  # roa_arrは後で計算
                # 粗利益率・D/E・ROICはyfinanceのデータを継続使用

        gm_arr = [gp_s[i]/rev_s[i]*100
                  if rev_s[i] and rev_s[i]!=0 else None
                  for i in range(min(len(rev_s),len(gp_s)))]

        def ratio(a, b, scale=100):
            return [a[i]/b[i]*scale
                    if a[i] is not None and b[i] and b[i]!=0 else None
                    for i in range(min(len(a),len(b)))]

        roe_arr  = ratio(ni_s, eq_s)
        roa_arr  = ratio(ni_s, ta_s)
        de_arr   = [abs(td_s[i]/eq_s[i])*100
                    if td_s[i] is not None and eq_s[i] and eq_s[i]!=0 else None
                    for i in range(min(len(td_s),len(eq_s)))]
        roic_arr = ratio(ni_s, ic_s) if ic_s else []

        if not fcf and ocf and capex:
            fcf = [ocf[i]+capex[i]
                   if ocf[i] is not None and capex[i] is not None else None
                   for i in range(min(len(ocf),len(capex)))]

        s_res, s_pass, s_total, s_gm = run_criteria(
            eps, gm_arr, ocf, fcf, roe_arr, roa_arr, de_arr, roic_arr, STRICT)
        r_res, r_pass, r_total, r_gm = run_criteria(
            eps, gm_arr, ocf, fcf, roe_arr, roa_arr, de_arr, roic_arr, RELAXED)

        s_all = (s_pass == s_total and s_total > 0)
        r_all = (r_pass == r_total and r_total > 0)

        if s_all:
            grade = "STRICT"
        elif r_all:
            grade = "RELAXED"
        else:
            grade = "FAIL"

        relaxed_items = [k for k in s_res
                         if s_res[k] is False and r_res.get(k) is True]



        # ── DCF妥当性チェック ──
        # EPS単位ミスマッチ（ADR・外国株など）を検出してDCFをスキップ
        eps_ok = True
        dcf_skip_reason = None
        if eps and price and price > 0:
            latest_eps = [v for v in eps if v is not None and not np.isnan(float(v))]
            if latest_eps:
                cur_eps = latest_eps[-1]
                if cur_eps <= 0:
                    eps_ok = False
                    dcf_skip_reason = "EPS負値"
                else:
                    pe = price / cur_eps
                    if pe < 0.1 or pe > 300:
                        # P/E が 1〜300 の範囲外は単位ミスマッチとみなす
                        eps_ok = False
                        dcf_skip_reason = f"P/E={pe:.1f}（単位ミスマッチ）"

        dcf = calc_dcf(eps) if (eps and eps_ok) else None

        # ── DCF結果の追加チェック ──
        # 買値が株価の 0.2〜3.0倍 の範囲外は単位ミスマッチとみなす
        # 正常な割安/割高銘柄は概ねこの範囲に収まる
        # 範囲外 = ADR・外国株のEPS単位ずれ・データ異常
        if dcf and price and price > 0:
            bp_check = dcf["dcf_buy_price"]
            ratio = bp_check / price
            if ratio < 0.15 or ratio > 3.5:
                dcf = None
                dcf_skip_reason = f"買値/株価={ratio:.2f}（単位ミスマッチ）"

        bp  = dcf["dcf_buy_price"] if dcf else None
        undervalued = (price < bp) if bp and price and price > 0 else None
        margin_pct  = round((bp-price)/bp*100,1) if bp and price and price>0 else None

        def last(arr):
            v = [x for x in arr if x is not None and not np.isnan(float(x))]
            return round(v[-1],2) if v else None

        return {
            "ticker": ticker, "name": name,
            "sector": sector, "market": market, "price": price,
            "s_score": s_pass, "s_total": s_total,
            "s_pct":   round(s_pass/s_total*100,1) if s_total else 0,
            "r_score": r_pass, "r_total": r_total,
            "r_pct":   round(r_pass/r_total*100,1) if r_total else 0,
            "grade":         grade,
            "relaxed_items": ", ".join(relaxed_items),
            **{f"s_{k}": v for k, v in s_res.items()},
            **{f"r_{k}": v for k, v in r_res.items()},
            "s_gm_note": s_gm,
            **(dcf if dcf else {
                "dcf_buy_price": None, "dcf_intrinsic":   None,
                "dcf_adj_rate":  None, "dcf_raw_rate":    None,
                "dcf_years_back":None, "dcf_past_eps":    None,
                "dcf_present_eps":None,
            }),
            "dcf_skip_reason": dcf_skip_reason,
            "undervalued": undervalued,
            "margin_pct":  margin_pct,
            "latest_EPS":  last(eps),
            "latest_GM":   last(gm_arr),
            "latest_ROE":  last(roe_arr),
            "latest_ROA":  last(roa_arr),
            "latest_DE":   last(de_arr),
            "latest_ROIC": last(roic_arr),
        }
        # モートスコアを計算して追加
        moat = calc_moat_score(res, _MOAT_DATA)
        res.update(moat)
        return res
    except Exception as e:
        return {"ticker": ticker, "market": market,
                "error": str(e)[:80], "grade": "ERROR",
                "s_score": None, "s_total": None, "undervalued": None}


# ── HTML生成 ───────────────────────────────────────────

CRITERIA_LABELS = {
    "EPS_positive":   "EPS+",
    "EPS_uptrend":    "EPS↑",
    "GM_judge":       "粗利",
    "OCF_positive":   "営業CF+",
    "OCF_uptrend":    "営業CF↑",
    "FCF_positive":   "FCF+",
    "FCF_uptrend":    "FCF↑",
    "ROE_above_thr":  f"ROE{STRICT['roe_thr']}%+",
    "ROE_uptrend":    "ROE↑",
    "ROA_above_thr":  f"ROA{STRICT['roa_thr']}%+",
    "ROA_uptrend":    "ROA↑",
    "DE_below_thr":   f"D/E{STRICT['de_thr']}%-",
    "ROIC_above_thr": f"ROIC{STRICT['roic_thr']}%+",
    "ROIC_vs_WACC":   "ROIC>WACC",
}

def _bool_cell(sv, rv):
    if sv is True:                return '<td class="p">✓</td>'
    if sv is False and rv is True:return '<td class="rx" title="緩和基準で通過">△</td>'
    if sv is False:               return '<td class="f">✗</td>'
    return '<td class="n">—</td>'

def _badge(grade, relaxed_items):
    if grade == "STRICT":
        return '<span class="badge b-strict">厳格通過</span>'
    return (f'<span class="badge b-relaxed" '
            f'title="緩和で通過した条件: {relaxed_items}">緩和通過</span>')

def _score_color(s, t):
    if not t: return "#888"
    r = s / t
    if r == 1.0: return "#1D9E75"
    if r >= 0.8: return "#5DCAA5"
    if r >= 0.6: return "#BA7517"
    return "#E24B4A"

def _moat_cell(row):
    """モートスコアのHTMLセルを生成"""
    score   = row.get("moat_score")
    grade   = row.get("moat_grade","")
    types   = row.get("moat_types",[])
    comment = str(row.get("moat_comment","")).replace('"',"'")
    risk    = str(row.get("moat_risk","")).replace('"',"'")

    try: score_f = float(score)
    except: return '<td class="n">—</td>'

    # 色分け
    if grade == "wide":
        bg, fg = "#1D3557", "#fff"
        label  = "◎Wide"
    elif grade == "narrow":
        bg, fg = "#457B9D", "#fff"
        label  = "○Narrow"
    else:
        bg, fg = "#e0e0e0", "#888"
        label  = "△None"

    types_str = " ".join(f"[{t}]" for t in (types if isinstance(types,list) else []))
    tooltip   = f"{types_str} | {comment} | リスク:{risk}"[:200]

    return (f'<td style="background:{bg};color:{fg};text-align:center;font-size:11px;'
            f'font-weight:600;white-space:nowrap;cursor:help" title="{tooltip}">'
            f'{label} {int(score_f)}</td>')


def make_html(df_all, total_input, generated):
    df = df_all[df_all["grade"].isin(["STRICT","RELAXED"])].copy()
    n_s    = len(df[df["grade"]=="STRICT"])
    n_r    = len(df[df["grade"]=="RELAXED"])
    n_uv_s = len(df[(df["undervalued"]==True)&(df["grade"]=="STRICT")])
    n_uv_r = len(df[(df["undervalued"]==True)&(df["grade"]=="RELAXED")])
    nc = len(CRITERIA_LABELS)

    th = "".join(f'<th onclick="sortTable({5+i})">{v}</th>'
                 for i,v in enumerate(CRITERIA_LABELS.values()))

    rows = []
    for _, row in df.iterrows():
        g   = row.get("grade","FAIL")
        ss  = row.get("s_score",0) or 0
        st  = row.get("s_total",0) or 0
        rs  = row.get("r_score",0) or 0
        rt  = row.get("r_total",0) or 0
        col = _score_color(ss, st)
        cells = "".join(_bool_cell(row.get(f"s_{k}"), row.get(f"r_{k}"))
                        for k in CRITERIA_LABELS)
        uv  = row.get("undervalued")
        mp  = row.get("margin_pct")
        bp  = row.get("dcf_buy_price")
        itr = row.get("dcf_intrinsic")
        pr  = row.get("price")
        uv_cell = (f'<td class="uv-yes">割安 ({mp:+.1f}%)</td>' if uv is True
                   else f'<td class="uv-no">割高 ({mp:+.1f}%)</td>' if uv is False
                   else '<td class="n">—</td>')
        score_disp = (f'{ss}/{st}'
                      + (f'<span class="r-score"> ({rs}/{rt})</span>'
                         if g=="RELAXED" else ""))
        asset_ratio = row.get("asset_undervaluation_ratio")
        is_asset_uv = str(row.get("is_asset_undervalued","")).strip().lower() == "true"
        vd = row.get("valuation_diff")
        rp_gain = row.get("rental_property_gain")
        try:
            asset_ratio_f = float(asset_ratio)
            if asset_ratio_f != asset_ratio_f:  # NaNチェック(NaN != NaN)
                asset_ratio_f = None
        except Exception:
            asset_ratio_f = None
        if asset_ratio_f is not None:
            try:
                vd_f = float(vd)
            except Exception:
                vd_f = 0.0
            tooltip_parts = [f"評価差額金:{vd_f:,.0f}円"]
            try:
                rp_f = float(rp_gain)
            except Exception:
                rp_f = 0.0
            if rp_f:
                tooltip_parts.append(f"賃貸不動産含み益:{rp_f:,.0f}円")
            tooltip = " | ".join(tooltip_parts)
            asset_cell = (f'<td class="uv-yes" title="{tooltip}">'
                          f'資産{asset_ratio_f:.2f} {"★" if is_asset_uv else ""}</td>')
        else:
            asset_cell = '<td class="n">—</td>'

        rows.append(
            f'<tr data-grade="{g}" data-uv="{"1" if uv else "0"}">'
            f'<td>{row.get("market","")}</td>'
            f'<td><b>{row.get("ticker","")}</b></td>'
            f'<td title="{row.get("s_gm_note","")}">{str(row.get("name",""))[:28]}</td>'
            f'<td>{str(row.get("sector",""))[:20]}</td>'
            f'<td style="color:{col};font-weight:700">{score_disp}</td>'
            f'<td>{_badge(g, str(row.get("relaxed_items","")))}</td>'
            f'<td class="num">{pr or "—"}</td>'
            f'<td class="num">{(str(bp) + " ✓") if bp else ("— " + str(row.get("dcf_skip_reason",""))[:20] if row.get("dcf_skip_reason") else "—")}</td>'
            f'<td class="num">{itr or "—"}</td>'
            f'{uv_cell}'
            f'{asset_cell}'
            f'{_moat_cell(row)}'
            f'<td class="gr" data-rate="{row.get("dcf_raw_rate") or 0}">{row.get("dcf_raw_rate","")}</td>'
            f'<td class="num">{row.get("latest_EPS","")}</td>'
            f'<td class="num">{row.get("latest_ROE","")}</td>'
            f'<td class="num">{row.get("latest_ROA","")}</td>'
            f'<td class="num">{row.get("latest_ROIC","")}</td>'
            f'<td class="num">{row.get("latest_DE","")}</td>'
            f'{cells}'
            f'</tr>'
        )

    s_cfg = (f"通過率{int(STRICT['pass_ratio']*100)}% / D/E<={STRICT['de_thr']}%"
             f" / ROE>={STRICT['roe_thr']}% / ROA>={STRICT['roa_thr']}%"
             f" / ROIC>={STRICT['roic_thr']}% / 連続落込{STRICT['consec_drop']}年"
             f" / 直近{STRICT['max_years']}年で評価")
    r_cfg = (f"通過率{int(RELAXED['pass_ratio']*100)}% / D/E<={RELAXED['de_thr']}%"
             f" / ROE>={RELAXED['roe_thr']}% / ROA>={RELAXED['roa_thr']}%"
             f" / ROIC>={RELAXED['roic_thr']}% / 連続落込{RELAXED['consec_drop']}年 / 直近{RELAXED['recent_years']}年で評価")

    return f"""<!DOCTYPE html>
<html lang="ja"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>VIA スクリーニング結果</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:system-ui,sans-serif;font-size:13px;background:#f5f5f0;color:#1a1a1a;padding:1rem}}
h1{{font-size:18px;font-weight:500;margin-bottom:.4rem}}
.meta{{font-size:12px;color:#888;margin-bottom:.5rem;line-height:1.6}}
.cfg{{font-size:11px;padding:5px 10px;border-radius:6px;margin-bottom:5px;line-height:1.7}}
.cfg-s{{background:#e8f8f2;border-left:3px solid #1D9E75}}
.cfg-r{{background:#fef3e2;border-left:3px solid #BA7517}}
.summary{{display:flex;gap:10px;margin-bottom:.8rem;flex-wrap:wrap}}
.sc{{background:#fff;border:1px solid #e0e0d8;border-radius:8px;padding:8px 14px;min-width:110px}}
.sc .sv{{font-size:22px;font-weight:600}} .sc .sl{{font-size:11px;color:#888}}
.ctrl{{display:flex;gap:8px;margin-bottom:6px;flex-wrap:wrap;align-items:center}}
#search{{padding:5px 10px;border:1px solid #ccc;border-radius:6px;font-size:13px;width:240px}}
.btn{{padding:5px 12px;border:1px solid #ccc;border-radius:6px;background:#fff;
      font-size:12px;cursor:pointer;color:#333}}
.btn:hover{{background:#f0f0e8}}
.btn.active{{background:#185FA5;color:#fff;border-color:#185FA5}}
.leg{{font-size:11px;color:#888;margin-bottom:8px}}
.tbl-wrap{{overflow-x:auto}}
table{{border-collapse:collapse;width:100%;white-space:nowrap;background:#fff}}
th{{background:#f0f0e8;font-weight:500;padding:5px 7px;border-bottom:2px solid #ccc;
    position:sticky;top:0;font-size:12px;cursor:pointer;user-select:none}}
th:hover{{background:#e4e4d8}}
td{{padding:4px 7px;border-bottom:1px solid #eee;font-size:12px;min-width:48px}}
tr:hover td{{background:#f9f9f5}}
tr[data-grade="RELAXED"]{{background:#fffbf0}}
tr[data-grade="RELAXED"]:hover td{{background:#fef5d8}}
td.p{{color:#1D9E75;font-weight:700;text-align:center}}
td.f{{color:#E24B4A;font-weight:700;text-align:center}}
td.rx{{color:#BA7517;font-weight:700;text-align:center;background:#fff8e8}}
td.n{{color:#aaa;text-align:center}}
td.num{{text-align:right;font-variant-numeric:tabular-nums}}
td.uv-yes{{color:#0F6E56;font-weight:700;background:#e8f8f2}}
td.uv-no{{color:#993C1D;font-size:11px}}
.badge{{display:inline-block;font-size:10px;padding:1px 7px;border-radius:10px;font-weight:600}}
.b-strict{{background:#e1f5ee;color:#0F6E56;border:1px solid #1D9E75}}
.b-relaxed{{background:#fef3e2;color:#854F0B;border:1px solid #BA7517}}
.badge-detail{{font-size:10px;color:#888}}
.r-score{{font-size:10px;color:#888}}
</style></head><body>
<h1>VIA スクリーニング結果</h1>
<p class="meta">生成: {generated} ／ WACC: {WACC}% ／ DCF割引率: {DCF_DISC_RATE}% ／ 安全領域: {DCF_MARGIN_SAFE}% ／ 対象: {total_input}銘柄</p>
<div class="cfg cfg-s"><b>厳格基準</b>（緑バッジ）: {s_cfg}</div>
<div class="cfg cfg-r"><b>緩和基準</b>（橙バッジ）: {r_cfg} ／ △=緩和で通過した条件</div>
<div class="summary">
  <div class="sc"><div class="sv" style="color:#1D9E75">{n_s}</div><div class="sl">厳格通過</div></div>
  <div class="sc"><div class="sv" style="color:#BA7517">{n_r}</div><div class="sl">緩和のみ通過</div></div>
  <div class="sc"><div class="sv" style="color:#1D9E75">{n_uv_s}</div><div class="sl">割安×厳格</div></div>
  <div class="sc"><div class="sv" style="color:#BA7517">{n_uv_r}</div><div class="sl">割安×緩和</div></div>
  <div class="sc"><div class="sv">{n_s+n_r}</div><div class="sl">合計通過</div></div>
</div>
<div class="ctrl">
  <input type="text" id="search" placeholder="銘柄名・ティッカーで絞り込み..." oninput="applyFilters()">
  <button class="btn active" id="btn-all"       onclick="setFilter('all')">全通過</button>
  <button class="btn"        id="btn-uv"         onclick="setFilter('uv')">割安</button>
  <button class="btn"        id="btn-uv-strict"  onclick="setFilter('uv-strict')">★割安×厳格</button>
  <button class="btn"        id="btn-uv-relaxed" onclick="setFilter('uv-relaxed')">割安×緩和</button>
  <button class="btn"        id="btn-strict"     onclick="setFilter('strict')">厳格通過のみ</button>
  <button class="btn"        id="btn-relaxed"    onclick="setFilter('relaxed')">緩和のみ通過</button>
</div>
<p class="leg">✓厳格通過 ／ △緩和で通過（橙色） ／ ✗NG ／ スコア: 厳格スコア (緩和スコア) ／ 列ヘッダーで並び替え
／ <b>資産過小評価</b>=NCAV(流動資産-負債総額)+投資有価証券(時価)-税効果額を時価総額と比較した比率（日本株のみ、EDINET有報から評価差額金を取得。1.0未満で★が過小評価）</p>
<div class="tbl-wrap"><table id="tbl">
<thead><tr>
  <th onclick="sortTable(0)">市場</th><th onclick="sortTable(1)">Ticker</th>
  <th onclick="sortTable(2)">名称</th><th onclick="sortTable(3)">セクター</th>
  <th onclick="sortTable(4)">スコア</th><th onclick="sortTable(5)">判定</th>
  <th onclick="sortTable(6)">現在株価</th>
  <th onclick="sortTable(7)">購買ターゲット価格</th>
  <th onclick="sortTable(8)">正味現在価値</th>
  <th onclick="sortTable(9)">割安判定</th>
  <th onclick="sortTable(10)">資産過小評価</th>
  <th onclick="sortTable(11)">モート</th>
  <th onclick="sortTable(12)">成長率%</th>
  <th onclick="sortTable(12)">EPS</th>
  <th onclick="sortTable(13)">ROE%</th>
  <th onclick="sortTable(14)">ROA%</th>
  <th onclick="sortTable(15)">ROIC%</th>
  <th onclick="sortTable(16)">D/E%</th>
  {th}
</tr></thead>
<tbody>{"".join(rows)}</tbody>
</table></div>
<script>
var cur='all',sd={{}};
function applyFilters(){{
  var q=document.getElementById('search').value.toLowerCase();
  document.querySelectorAll('#tbl tbody tr').forEach(function(r){{
    var tm=!q||r.textContent.toLowerCase().includes(q);
    var g=r.dataset.grade,uv=r.dataset.uv==='1',fm=true;
    if(cur==='strict')     fm=g==='STRICT';
    if(cur==='relaxed')    fm=g==='RELAXED';
    if(cur==='uv')         fm=uv;
    if(cur==='uv-strict')  fm=uv&&g==='STRICT';
    if(cur==='uv-relaxed') fm=uv&&g==='RELAXED';
    r.style.display=(tm&&fm)?'':'none';
  }});
}}
function setFilter(f){{
  cur=f;
  ['all','uv','uv-strict','uv-relaxed','strict','relaxed'].forEach(function(id){{
    var b=document.getElementById('btn-'+id);
    if(b)b.className='btn'+(id===f?' active':'');
  }});
  applyFilters();
}}
function sortTable(c){{
  var tb=document.querySelector('#tbl tbody');
  var rows=Array.from(tb.rows);
  sd[c]=!sd[c];
  rows.sort(function(a,b){{
    var av=a.cells[c]?a.cells[c].textContent.trim():'';
    var bv=b.cells[c]?b.cells[c].textContent.trim():'';
    var an=parseFloat(av.replace(/[^0-9.]/g,'')),bn=parseFloat(bv.replace(/[^0-9.]/g,''));
    if(!isNaN(an)&&!isNaN(bn))return sd[c]?an-bn:bn-an;
    return sd[c]?av.localeCompare(bv):bv.localeCompare(av);
  }});
  rows.forEach(function(r){{tb.appendChild(r);}});
  applyFilters();
}}
setFilter('all');
// 成長率セルの色分け
document.querySelectorAll('td.gr').forEach(function(td){{
  var v=parseFloat(td.getAttribute('data-rate')||0);
  if(v>=25)      {{td.style.textAlign='right';td.style.fontWeight='700';td.style.color='#fff';td.style.background='#C0392B';}}
  else if(v>=20) {{td.style.textAlign='right';td.style.fontWeight='700';td.style.color='#fff';td.style.background='#E67E22';}}
  else if(v>=10) {{td.style.textAlign='right';td.style.fontWeight='600';td.style.color='#333';td.style.background='#FFF0B3';}}
  else if(v>0)   {{td.style.textAlign='right';td.style.color='#888';}}
  else           {{td.style.textAlign='right';td.style.color='#ccc';}}
}})
</script></body></html>"""


# ── メイン ─────────────────────────────────────────────

def main():
    # ── 引数処理: python via_screener.py [us|jp|all] ──
    import sys as _sys
    arg = _sys.argv[1].lower() if len(_sys.argv) > 1 else "all"
    run_us = arg in ("all", "us")
    run_jp = arg in ("all", "jp")
    if arg not in ("all", "us", "jp"):
        print(f"引数エラー: '{arg}'")
        print("使い方: python via_screener.py [us|jp|all]")
        print("  us  → 米国株のみ（EODHD）")
        print("  jp  → 日本株のみ（J-Quants）")
        print("  all → 全銘柄（省略可）")
        return

    print("=" * 60)
    print("VIA スクリーナー 起動")
    print(f"対象市場: {'US+JP' if arg=='all' else arg.upper()}")
    print(f"厳格: 通過率{int(STRICT['pass_ratio']*100)}% / "
          f"D/E<={STRICT['de_thr']}% / ROE>={STRICT['roe_thr']}% / "
          f"ROA>={STRICT['roa_thr']}% / ROIC>={STRICT['roic_thr']}% / "
          f"連続落込{STRICT['consec_drop']}年")
    print(f"緩和: 通過率{int(RELAXED['pass_ratio']*100)}% / "
          f"D/E<={RELAXED['de_thr']}% / ROE>={RELAXED['roe_thr']}% / "
          f"ROA>={RELAXED['roa_thr']}% / ROIC>={RELAXED['roic_thr']}% / "
          f"連続落込{RELAXED['consec_drop']}年")
    print("=" * 60)

    # ── STEP1: 銘柄リスト全取得（完了してから次へ） ──
    # モートデータ読み込み
    global _MOAT_DATA
    _MOAT_DATA = load_moat_data()
    print(f"モートデータ: {len(_MOAT_DATA)} 銘柄登録済み")

    # キャッシュフォルダ作成
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache_files = [f for f in os.listdir(CACHE_DIR) if f.endswith(".json")]
    eodhd_cached = len([f for f in cache_files if f.startswith("eodhd_")])
    jq_cached    = len([f for f in cache_files if f.startswith("jquants_")])
    print(f"キャッシュ状況: EODHD={eodhd_cached}銘柄  J-Quants={jq_cached}銘柄")
    print(f"（キャッシュ有効期限: {FINS_CACHE_DAYS}日）")
    print()

    us_tickers, jp_tickers = get_tickers()
    all_tickers = []
    if run_us:
        all_tickers += [(t,"US") for t in us_tickers]
    if run_jp:
        all_tickers += [(t,"JP") for t in jp_tickers]
    total    = len(all_tickers)
    est_min  = round(total * SLEEP_SEC / 60)

    print(f"\n対象: US {len(us_tickers)} + JP {len(jp_tickers)} = {total} 銘柄")
    print(f"推定時間: 約 {est_min} 分")
    print()
    print("スクリーニングを開始します...")
    print("=" * 60)

    # ── STEP2: スクリーニング ──
    rows = []
    for i, (ticker, market) in enumerate(all_tickers):
        pct_done = (i+1)/total*100
        print(f"[{i+1:4d}/{total}] {ticker:<12} ({pct_done:5.1f}%)",
              end=" ", flush=True)
        res = process_ticker(ticker, market)
        if res:
            rows.append(res)
            g  = res.get("grade","?")
            ss = res.get("s_score"); st = res.get("s_total")
            bp = res.get("dcf_buy_price"); pr = res.get("price")
            uv = res.get("undervalued"); ri = res.get("relaxed_items","")
            if ss is not None and st:
                uv_str = " 【割安】" if uv else ""
                rx_str = f" 緩和:{ri}" if g=="RELAXED" and ri else ""
                print(f"→ {g:<7} {ss}/{st}  購買ターゲット:{bp}  株価:{pr}{uv_str}{rx_str}")
            else:
                print(f"→ エラー: {res.get('error','?')[:50]}")
        else:
            print("→ データなし")
        time.sleep(SLEEP_SEC)

    # ── STEP3: 出力 ──
    df = pd.DataFrame(rows)
    df_valid = df[df["s_score"].notna()].copy()
    grade_ord = {"STRICT":0,"RELAXED":1,"FAIL":2,"ERROR":3}
    df_valid["_go"] = df_valid["grade"].map(grade_ord).fillna(3)
    df_valid["_uo"] = df_valid["undervalued"].apply(lambda x: 0 if x else 1)
    df_valid = df_valid.sort_values(
        ["_uo","_go","s_score","s_pct"], ascending=[True,True,False,False]
    ).drop(columns=["_go","_uo"])

    # ── STEP2.5: VIA通過銘柄に資産過小評価スコアを付与（日本株のみ） ──
    passed_mask = df_valid["grade"].isin(["STRICT", "RELAXED"])
    passed_jp = df_valid[passed_mask & (df_valid["market"] == "JP")]
    if _edinet is not None and not passed_jp.empty:
        print(f"\n[資産過小評価スコア計算] 対象: {len(passed_jp)}銘柄（VIA通過の日本株）")
        for idx, row in passed_jp.iterrows():
            ticker = row["ticker"]
            price  = row.get("price")
            # 時価総額を概算（株価のみ取得済みのため、必要なら再取得）
            try:
                tk_tmp = yf.Ticker(ticker)
                mcap = tk_tmp.fast_info.market_cap
            except Exception:
                mcap = None
            print(f"  {ticker:<10}", end=" ", flush=True)
            asset_result = calc_asset_undervaluation(ticker, row["market"], mcap)
            if asset_result:
                def _s(v):
                    return str(v) if v is not None else None
                df_valid.at[idx, "ncav"] = _s(asset_result.get("ncav"))
                df_valid.at[idx, "ncav_plus_tax_adjusted"] = _s(asset_result.get("ncav_plus_tax_adjusted"))
                df_valid.at[idx, "asset_undervaluation_ratio"] = _s(asset_result.get("undervaluation_ratio"))
                df_valid.at[idx, "is_asset_undervalued"] = _s(asset_result.get("is_undervalued"))
                vd_data = asset_result.get("valuation_diff_data") or {}
                rp_data = asset_result.get("rental_property_data") or {}
                df_valid.at[idx, "valuation_diff"] = _s(vd_data.get("valuation_diff"))
                df_valid.at[idx, "rental_property_gain"] = _s(rp_data.get("unrealized_gain"))
                df_valid.at[idx, "rental_property_book_value"] = _s(rp_data.get("book_value"))
                df_valid.at[idx, "rental_property_fair_value"] = _s(rp_data.get("fair_value"))
                inv_risk = asset_result.get("inventory_risk_data") or {}
                df_valid.at[idx, "inventory_growth_gap_pct"] = _s(inv_risk.get("growth_gap_pct"))
                df_valid.at[idx, "is_inventory_risk"] = _s(inv_risk.get("is_inventory_risk"))
                rb_data = asset_result.get("retirement_benefit_data") or {}
                df_valid.at[idx, "retirement_unrealized_diff"] = _s(rb_data.get("unrealized_diff"))
                df_valid.at[idx, "is_retirement_risk"] = _s(asset_result.get("is_retirement_risk"))
                ratio_disp = asset_result.get("undervaluation_ratio")
                rp_str = f" 不動産含み益:{rp_data.get('unrealized_gain',0):,.0f}円" if rp_data.get("unrealized_gain") else ""
                print(f"-> NCAV+調整後比率:{ratio_disp}{rp_str}  "
                      f"{'★資産過小評価' if asset_result.get('is_undervalued') else ''}")
            else:
                print("-> データ取得不可")
            time.sleep(0.5)

    df_valid.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"\nCSV出力: {OUTPUT_CSV}")

    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(make_html(df_valid, total, generated))
    print(f"HTML出力: {OUTPUT_HTML}")

    n_s    = len(df_valid[df_valid["grade"]=="STRICT"])
    n_r    = len(df_valid[df_valid["grade"]=="RELAXED"])
    n_uv_s = len(df_valid[(df_valid["undervalued"]==True)&(df_valid["grade"]=="STRICT")])
    n_uv_r = len(df_valid[(df_valid["undervalued"]==True)&(df_valid["grade"]=="RELAXED")])

    print("\n" + "=" * 60)
    print(f"厳格通過:         {n_s:4d} 銘柄")
    print(f"緩和のみ通過:     {n_r:4d} 銘柄")
    print(f"割安 × 厳格:      {n_uv_s:4d} 銘柄  ← 最優先")
    print(f"割安 × 緩和:      {n_uv_r:4d} 銘柄  ← 参考")
    print("=" * 60)

    for label, mask in [
        ("割安×厳格通過",
         (df_valid["undervalued"]==True)&(df_valid["grade"]=="STRICT")),
        ("割安×緩和のみ通過",
         (df_valid["undervalued"]==True)&(df_valid["grade"]=="RELAXED")),
    ]:
        top = df_valid[mask]
        if not top.empty:
            print(f"\n【{label}】")
            for _, r in top.iterrows():
                ri = r.get("relaxed_items","")
                rx = f"  緩和条件:{ri}" if ri else ""
                print(f"  {r['market']:2s} {r['ticker']:<12} "
                      f"{str(r.get('name',''))[:28]:<30} "
                      f"株価:{r.get('price')}  買値:{r.get('dcf_buy_price')}  "
                      f"乖離:{r.get('margin_pct'):+.1f}%{rx}")

    print(f"\n完了。{OUTPUT_HTML} をブラウザで開いてください。")

if __name__ == "__main__":
    main()
