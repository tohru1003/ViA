"""
VIA EDINET連携モジュール
有報XBRLから資産面の過小評価要因（その他有価証券評価差額金・賃貸等不動産含み益）
を取得し、NCAVと組み合わせて「資産過小評価スコア」を算出する。

via_screener.py から import して使用する。

主な公開関数:
  load_edinet_mapping()        証券コード→EDINETコードのマッピングを取得（キャッシュ付き）
  get_valuation_difference(ticker, mapping, fiscal_year_end)
                                 その他有価証券評価差額金（含み益）を取得
  get_rental_property_gain(ticker, mapping, fiscal_year_end)
                                 賃貸等不動産の含み益を取得
  calc_asset_undervaluation_score(...)
                                 1銘柄の資産過小評価データを取得（NCAV等含む統合スコア）
"""

import os, time, json, urllib.request, ssl, zipfile, re
import pandas as pd
from html import unescape

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EDINET_DIR = os.path.join(SCRIPT_DIR, "edinet_temp")
CACHE_DIR  = os.path.join(SCRIPT_DIR, "via_cache")
MAPPING_CACHE = os.path.join(EDINET_DIR, "edinet_mapping.json")
MAPPING_CACHE_DAYS = 30

EDINET_API_KEY = "f150625f5cee4c829662a23da9700c7f"
EDINET_BASE    = "https://api.edinet-fsa.go.jp/api/v2"
EFFECTIVE_TAX_RATE = 0.30  # 実効税率（評価差額金からの含み益逆算に使用）

os.makedirs(EDINET_DIR, exist_ok=True)

_ctx = ssl.create_default_context()
_ctx.check_hostname = False
_ctx.verify_mode    = ssl.CERT_NONE


def _http_get(url, timeout=20):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=timeout, context=_ctx) as r:
        return r.read()


# ── 証券コード → EDINETコード マッピング ──

def load_edinet_mapping(force_refresh=False):
    """
    証券コード(4桁)→EDINETコードの辞書を返す。
    30日キャッシュ。なければEDINETから一覧CSVをダウンロードする。
    """
    if not force_refresh and os.path.exists(MAPPING_CACHE):
        age_days = (time.time() - os.path.getmtime(MAPPING_CACHE)) / 86400
        if age_days <= MAPPING_CACHE_DAYS:
            try:
                with open(MAPPING_CACHE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass

    try:
        url = "https://disclosure2dl.edinet-fsa.go.jp/searchdocument/codelist/Edinetcode.zip"
        data = _http_get(url, timeout=30)
        zip_path = os.path.join(EDINET_DIR, "Edinetcode.zip")
        with open(zip_path, "wb") as f:
            f.write(data)
        with zipfile.ZipFile(zip_path) as z:
            z.extractall(EDINET_DIR)

        csv_path = os.path.join(EDINET_DIR, "EdinetcodeDlInfo.csv")
        df = pd.read_csv(csv_path, encoding="cp932", skiprows=1, dtype=str)
        df["secCode4"] = df["証券コード"].astype(str).str[:4]
        mapping = dict(zip(df["secCode4"], df["ＥＤＩＮＥＴコード"]))

        with open(MAPPING_CACHE, "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False)
        return mapping
    except Exception as e:
        print(f"  [EDINET] マッピング取得失敗: {e}")
        if os.path.exists(MAPPING_CACHE):
            with open(MAPPING_CACHE, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}


# ── 有報docID検索 ──

def _find_yuho_docid(edinet_code, fiscal_year_end=None, max_back_days=400):
    """
    指定EDINETコードの直近の有報(docTypeCode=120)のdocIDを検索する。

    fiscal_year_end: "MM-DD"形式（例: "03-31"）が分かれば、
                      決算日から45〜150日後の範囲に検索を絞り高速化する。
                      Noneの場合は直近max_back_days日を1日ずつ遡る（低速フォールバック）。
    """
    from datetime import datetime, timedelta
    today = datetime.now()

    search_dates = []
    if fiscal_year_end:
        try:
            mm, dd = map(int, fiscal_year_end.split("-"))
            for year_offset in [0, -1]:
                fy_end = datetime(today.year + year_offset, mm, dd)
                for delta in range(45, 150):
                    candidate = fy_end + timedelta(days=delta)
                    if candidate <= today:
                        search_dates.append(candidate.strftime("%Y-%m-%d"))
        except Exception:
            pass
        search_dates = sorted(set(search_dates), reverse=True)

    if not search_dates:
        search_dates = [(today - timedelta(days=b)).strftime("%Y-%m-%d")
                         for b in range(0, max_back_days)]

    for date in search_dates:
        url = f"{EDINET_BASE}/documents.json?date={date}&type=2&Subscription-Key={EDINET_API_KEY}"
        try:
            data = json.loads(_http_get(url, timeout=15))
        except Exception:
            continue
        for r in (data.get("results") or []):
            if r.get("edinetCode") == edinet_code and r.get("docTypeCode") == "120":
                return r.get("docID"), r.get("periodEnd")
        time.sleep(0.3)
    return None, None


# ── XBRL ZIP取得（共有ヘルパー）──

def _fetch_xbrl_content(doc_id):
    """
    docIDを指定してXBRL ZIPを取得し、PublicDoc内の.xbrlファイルの
    生テキスト（デコード済み）を返す。失敗時はNone。
    """
    try:
        url = f"{EDINET_BASE}/documents/{doc_id}?type=1&Subscription-Key={EDINET_API_KEY}"
        zip_data = _http_get(url, timeout=30)

        zip_path = os.path.join(EDINET_DIR, f"{doc_id}.zip")
        with open(zip_path, "wb") as f:
            f.write(zip_data)

        with zipfile.ZipFile(zip_path) as z:
            xbrl_names = [n for n in z.namelist()
                          if n.startswith("XBRL/PublicDoc/") and n.endswith(".xbrl")]
            if not xbrl_names:
                return None
            content = z.read(xbrl_names[0]).decode("utf-8", errors="ignore")

        try:
            os.remove(zip_path)
        except Exception:
            pass

        return content
    except Exception:
        return None


# ── ① その他有価証券評価差額金の抽出 ──

def _parse_valuation_difference(content):
    """
    XBRL生テキストからその他有価証券評価差額金（連結・当期末）を抽出する。
    戻り値: dict {valuation_diff, note} または None
    """
    result = {}
    pattern = r'<jppfs_cor:ValuationDifferenceOnAvailableForSaleSecurities\s+([^>]*)>([^<]*)</jppfs_cor:ValuationDifferenceOnAvailableForSaleSecurities>'
    found_current = False
    for attrs, value in re.findall(pattern, content):
        ctx_match = re.search(r'contextRef="([^"]*)"', attrs)
        ctx = ctx_match.group(1) if ctx_match else ""
        if ctx == "CurrentYearInstant":
            try:
                result["valuation_diff"] = float(value)
                found_current = True
            except Exception:
                pass
            break

    if not found_current:
        result["valuation_diff"] = 0.0
        result["note"] = "当期分タグなし（評価差額金ゼロまたは非開示と推定）"

    return result if result else None


def _parse_retirement_benefit_remeasurement(content):
    """
    XBRL生テキストから退職給付に係る未認識数理計算上の差異
    （RemeasurementsOfDefinedBenefitPlans、連結・当期末・税効果後）を抽出する。

    この金額はすでに連結BSの「退職給付に係る負債」に反映済みのため、
    NCAVから追加で減算する対象ではない。将来の年金資産・負債の変動リスクを
    示す警告指標として扱う。

    戻り値: dict {remeasurement_after_tax, unrealized_diff, note} または None
    """
    pattern = r'<jppfs_cor:RemeasurementsOfDefinedBenefitPlans\s+([^>]*)>([^<]*)</jppfs_cor:RemeasurementsOfDefinedBenefitPlans>'
    found_current = False
    value = None
    for attrs, v in re.findall(pattern, content):
        ctx_match = re.search(r'contextRef="([^"]*)"', attrs)
        ctx = ctx_match.group(1) if ctx_match else ""
        if ctx == "CurrentYearInstant":
            try:
                value = float(v)
                found_current = True
            except Exception:
                pass
            break

    if not found_current:
        return None  # 退職給付制度の開示自体が無い、または当期分タグなし

    unrealized_diff = value / (1 - EFFECTIVE_TAX_RATE) if value else 0
    return {
        "remeasurement_after_tax": round(value, 0),
        "unrealized_diff": round(unrealized_diff, 0),
    }


# ── ② 賃貸等不動産の含み益の抽出 ──

def _parse_jp_number(s):
    """日本語の三角(△)はマイナス、カンマ除去して数値化（単位:千円のまま）"""
    if s is None:
        return None
    s = str(s).replace(',', '').replace('△', '-').strip()
    try:
        return float(s)
    except Exception:
        return None


def _parse_rental_property(content):
    """
    XBRL生テキストから賃貸等不動産関係の注記（TextBlock）を抽出し、
    HTMLテーブルをパースして「期末残高(BS計上額)」と「期末時価」を取得する。
    複数の不動産カテゴリがある場合は合算する。

    戻り値: dict {book_value, fair_value, unrealized_gain（円単位）} または None
    """
    pattern = r'<jpcrp_cor:NotesRealEstateForLeaseEtc[a-zA-Z]*TextBlock[^>]*>(.*?)</jpcrp_cor:NotesRealEstateForLeaseEtc[a-zA-Z]*TextBlock>'
    match = re.search(pattern, content, re.DOTALL)
    if not match:
        return None  # 賃貸等不動産の開示自体が無い（保有していない）企業

    raw_html = unescape(match.group(1))

    try:
        import io
        tables = pd.read_html(io.StringIO(raw_html))
    except Exception:
        return None

    book_values = []
    fair_values = []

    for t in tables:
        t_str = t.astype(str)
        for _, row in t_str.iterrows():
            row_text = ' '.join(row.values)
            numeric_cells = [_parse_jp_number(v) for v in row.values]
            numeric_cells = [v for v in numeric_cells if v is not None and abs(v) > 100]
            if not numeric_cells:
                continue
            latest_val = numeric_cells[-1]  # 最新期は通常最後の列
            if '期末残高' in row_text:
                book_values.append(latest_val)
            elif '期末時価' in row_text:
                fair_values.append(latest_val)

    if not book_values or not fair_values:
        return None

    # 複数カテゴリ（住宅、オフィス等）があれば合算。単位は千円。
    book_value_total = sum(book_values) * 1000   # 円に変換
    fair_value_total = sum(fair_values) * 1000
    gain = fair_value_total - book_value_total

    return {
        "book_value": round(book_value_total, 0),
        "fair_value": round(fair_value_total, 0),
        "unrealized_gain": round(gain, 0),
    }


# ── キャッシュ ──

def _cache_path(ticker, suffix=""):
    safe = ticker.replace(".", "_")
    return os.path.join(CACHE_DIR, f"edinet_{safe}{suffix}.json")

def _load_cache(ticker, suffix="", max_days=30):
    path = _cache_path(ticker, suffix)
    if not os.path.exists(path):
        return None
    age_days = (time.time() - os.path.getmtime(path)) / 86400
    if age_days > max_days:
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

def _save_cache(ticker, data, suffix=""):
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(_cache_path(ticker, suffix), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)


# ── 公開関数: XBRL取得の統合エントリーポイント ──

def _get_xbrl_data(ticker, mapping=None, fiscal_year_end=None):
    """
    1銘柄のXBRLを取得し、評価差額金・賃貸等不動産の両方を一度に解析する。
    XBRLダウンロードは1回のみ行い、キャッシュは項目ごとに分けて保存する
    （他関数から個別に呼べるようにするため）。
    """
    cached_vd = _load_cache(ticker, "_vd")
    cached_rp = _load_cache(ticker, "_rp")
    cached_rb = _load_cache(ticker, "_rb")
    if cached_vd is not None and cached_rp is not None and cached_rb is not None:
        return cached_vd, cached_rp, cached_rb

    if mapping is None:
        mapping = load_edinet_mapping()

    code4 = ticker.replace(".T", "")
    edinet_code = mapping.get(code4)
    if not edinet_code:
        return None, None, None

    doc_id, period_end = _find_yuho_docid(edinet_code, fiscal_year_end=fiscal_year_end)
    if not doc_id:
        return None, None, None

    content = _fetch_xbrl_content(doc_id)
    if not content:
        return None, None, None

    vd_raw = _parse_valuation_difference(content)
    rp_raw = _parse_rental_property(content)
    rb_raw = _parse_retirement_benefit_remeasurement(content)

    vd_result = None
    if vd_raw:
        vd = vd_raw["valuation_diff"]
        unrealized_gain = vd / (1 - EFFECTIVE_TAX_RATE) if vd else 0
        deferred_tax = unrealized_gain - vd
        vd_result = {
            "valuation_diff": round(vd, 0),
            "unrealized_gain": round(unrealized_gain, 0),
            "deferred_tax": round(deferred_tax, 0),
            "doc_id": doc_id,
            "period_end": period_end,
            "note": vd_raw.get("note", ""),
        }
    _save_cache(ticker, vd_result or {}, "_vd")

    rp_result = None
    if rp_raw:
        rp_result = dict(rp_raw)
        rp_result["doc_id"] = doc_id
        rp_result["period_end"] = period_end
    _save_cache(ticker, rp_result or {}, "_rp")

    rb_result = None
    if rb_raw:
        rb_result = dict(rb_raw)
        rb_result["doc_id"] = doc_id
        rb_result["period_end"] = period_end
    _save_cache(ticker, rb_result or {}, "_rb")

    return vd_result, rp_result, rb_result


def get_valuation_difference(ticker, mapping=None, fiscal_year_end=None):
    """
    その他有価証券評価差額金を取得する（後方互換のための単独関数）。
    内部的には_get_xbrl_dataを呼び、両方の項目を一度に取得・キャッシュする。
    """
    vd, _, _ = _get_xbrl_data(ticker, mapping, fiscal_year_end)
    return vd if vd else None


def get_rental_property_gain(ticker, mapping=None, fiscal_year_end=None):
    """
    賃貸等不動産の含み益を取得する。
    保有していない企業の場合はNoneを返す（過小評価判定では0として扱う）。
    """
    _, rp, _ = _get_xbrl_data(ticker, mapping, fiscal_year_end)
    return rp if rp else None


def calc_inventory_risk(inventory_series, revenue_series):
    """
    棚卸資産の急増リスクを判定する（劣化リスクの代理指標）。
    inventory_series, revenue_series: 新しい順のリスト（yfinanceの.dropna().tolist()形式）

    在庫が事業規模に対して意味のある金額を持つ企業（在庫/売上高比率5%以上）に限定し、
    在庫成長率が売上成長率を大幅に上回る場合（差20pt以上）を危険信号とする。
    在庫が僅少な企業（IT・サービス業等）はノイズが大きいため対象外とする。

    戻り値: dict または None（データ不足・対象外時）
    """
    if not inventory_series or not revenue_series:
        return None
    if len(inventory_series) < 2 or len(revenue_series) < 2:
        return None

    inv_latest, inv_prev = inventory_series[0], inventory_series[1]
    rev_latest, rev_prev = revenue_series[0], revenue_series[1]

    if not inv_prev or not rev_prev or not rev_latest:
        return None

    inv_to_rev_latest = inv_latest / rev_latest
    inv_to_rev_prev = inv_prev / rev_prev if rev_prev else None

    # 在庫が売上高比5%未満の企業は「在庫ビジネスではない」と判断し対象外とする
    # （ノイズの大きい極小在庫の成長率変動を排除）
    MIN_INVENTORY_TO_REVENUE_RATIO = 0.05
    if inv_to_rev_latest < MIN_INVENTORY_TO_REVENUE_RATIO:
        return {
            "inv_growth_pct": None,
            "rev_growth_pct": None,
            "growth_gap_pct": None,
            "inv_to_rev_ratio_latest": round(inv_to_rev_latest, 3),
            "inv_to_rev_ratio_change": None,
            "is_inventory_risk": False,
            "skip_reason": "在庫規模が僅少（対象外）",
        }

    inv_growth = (inv_latest - inv_prev) / inv_prev * 100
    rev_growth = (rev_latest - rev_prev) / rev_prev * 100
    gap = inv_growth - rev_growth
    ratio_change = (inv_to_rev_latest - inv_to_rev_prev) if inv_to_rev_prev else None

    return {
        "inv_growth_pct": round(inv_growth, 1),
        "rev_growth_pct": round(rev_growth, 1),
        "growth_gap_pct": round(gap, 1),
        "inv_to_rev_ratio_latest": round(inv_to_rev_latest, 3),
        "inv_to_rev_ratio_change": round(ratio_change, 3) if ratio_change is not None else None,
        "is_inventory_risk": gap > 20,
    }


def calc_asset_undervaluation_score(ticker, market_cap, current_assets, total_liab,
                                      investments=0, mapping=None, fiscal_year_end=None,
                                      inventory_series=None, revenue_series=None,
                                      minority_interest=0):
    """
    資産過小評価スコアを計算する。
    NCAV + 投資有価証券(時価) + 賃貸等不動産の含み益 - 税効果額 を時価総額と比較し、
    過小評価の度合いをスコア化する。

    戻り値: dict {
        "ncav": NCAV,
        "ncav_plus_tax_adjusted": 税効果・不動産含み益調整後NCAV-Plus,
        "valuation_diff_data": 評価差額金データ,
        "rental_property_data": 賃貸等不動産データ,
        "undervaluation_ratio": 時価総額 / NCAV-Plus（低いほど過小評価）,
        "is_undervalued": NCAV-Plusが時価総額を上回るか,
    }
    """
    if current_assets is None or total_liab is None:
        return None

    # NCAV = 流動資産 - 負債総額 - 非支配株主持分
    # （連結純資産から非支配株主持分を除き、親会社株主に帰属する清算価値のみを算出）
    ncav = current_assets - total_liab - (minority_interest or 0)
    vd_data, rp_data, rb_data = _get_xbrl_data(ticker, mapping, fiscal_year_end)

    deferred_tax = vd_data.get("deferred_tax", 0) if vd_data else 0
    rental_gain  = rp_data.get("unrealized_gain", 0) if rp_data else 0

    # 賃貸等不動産の含み益にも簡易的に税効果（30%）を適用する
    # （その他有価証券同様、売却時には課税されるため保守的に調整）
    rental_gain_after_tax = rental_gain * (1 - EFFECTIVE_TAX_RATE) if rental_gain else 0

    ncav_plus_adj = ncav + investments - deferred_tax + rental_gain_after_tax

    ratio = (market_cap / ncav_plus_adj) if ncav_plus_adj and ncav_plus_adj > 0 else None
    is_undervalued = (ratio is not None and ratio < 1.0)

    inventory_risk = calc_inventory_risk(inventory_series, revenue_series)

    # 退職給付の未認識数理計算上の差異の警告判定
    # （すでにBS負債に反映済みのため減算はしないが、純資産対比で大きい場合は警告フラグとする）
    retirement_risk_flag = False
    if rb_data and rb_data.get("unrealized_diff"):
        diff = rb_data["unrealized_diff"]
        # NCAV（簡易的に純資産規模の代理）の5%以上を「大きい」とみなす
        if ncav and abs(diff) >= ncav * 0.05:
            retirement_risk_flag = True

    return {
        "ncav": round(ncav, 0),
        "minority_interest_excluded": round(minority_interest or 0, 0),
        "ncav_plus_tax_adjusted": round(ncav_plus_adj, 0),
        "valuation_diff_data": vd_data,
        "rental_property_data": rp_data,
        "inventory_risk_data": inventory_risk,
        "retirement_benefit_data": rb_data,
        "is_retirement_risk": retirement_risk_flag,
        "undervaluation_ratio": round(ratio, 3) if ratio is not None else None,
        "is_undervalued": is_undervalued,
    }
