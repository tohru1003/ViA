"""
VIA HTMLリビルドスクリプト
既存のvia_results.csvからスコアを再計算してHTMLを再生成します。
yfinanceによる再取得は不要。約30秒で完了。

使い方:
  python via_rebuild_html.py
"""

import pandas as pd
import numpy as np
import os
from datetime import datetime

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
MOAT_DATA   = {}  # モートデータ（main()で読み込み）

# ── モートスコア計算 ─────────────────────────────────
def calc_moat_score(row, moat_db):
    def to_f(v):
        try: return float(v) if v not in (None,"","nan","None") else 0.0
        except: return 0.0
    roe    = to_f(row.get("latest_ROE"))
    roa    = to_f(row.get("latest_ROA"))
    roic   = to_f(row.get("latest_ROIC"))
    gm     = to_f(row.get("latest_GM"))
    growth = to_f(row.get("dcf_raw_rate"))
    q_roe    = min(roe   / 20 * 20, 20) if roe   > 0 else 0
    q_roa    = min(roa   / 10 * 15, 15) if roa   > 0 else 0
    q_roic   = min(roic  / 15 * 15, 15) if roic  > 0 else 0
    q_gm     = min(gm    / 40 * 15, 15) if gm    > 0 else 0
    q_growth = min(growth/ 15 *  5,  5) if growth > 0 else 0
    quant  = round(q_roe + q_roa + q_roic + q_gm + q_growth, 1)
    ticker = str(row.get("ticker",""))
    info   = moat_db.get(ticker, {})
    bonus  = info.get("bonus", 0)
    grade  = info.get("moat_grade", "")
    types  = info.get("moat_type", [])
    comment= info.get("comment", "")
    risk   = info.get("risk", "")
    # スコアはbonus値のみ（上限なし）
    total = bonus
    if not grade:
        # 未登録銘柄は定量スコアで自動判定
        quant_grade = quant + bonus
        if quant_grade >= 75: grade = "wide"
        elif quant_grade >= 55: grade = "narrow"
        else: grade = "none"
    return {"moat_score": total, "moat_grade": grade,
            "moat_types": types, "moat_comment": comment, "moat_risk": risk}

def _asset_cell(row):
    ratio = row.get("asset_undervaluation_ratio")
    if ratio is None or str(ratio) in ("", "nan", "None"):
        return '<td class="n">—</td>'
    try:
        ratio_f = float(ratio)
    except Exception:
        return '<td class="n">—</td>'
    is_uv = str(row.get("is_asset_undervalued","")).lower() == "true"
    vd = row.get("valuation_diff")
    rp = row.get("rental_property_gain")
    parts = []
    try:
        vd_f = float(vd)
        parts.append(f"評価差額金:{vd_f:,.0f}円")
    except Exception:
        pass
    try:
        rp_f = float(rp)
        if rp_f and rp_f == rp_f:  # NaN チェック（NaN != NaN）
            parts.append(f"賃貸不動産含み益:{rp_f:,.0f}円")
    except Exception:
        pass
    tooltip = " | ".join(parts)
    cls = "uv-yes" if is_uv else ""
    star = " ★" if is_uv else ""
    return f'<td class="{cls}" title="{tooltip}">資産{ratio_f:.2f}{star}</td>'


def _moat_cell(moat):
    score = moat.get("moat_score", 0)
    grade = moat.get("moat_grade", "none")
    types = moat.get("moat_types", [])
    comment = str(moat.get("moat_comment","")).replace('"',"'")
    risk    = str(moat.get("moat_risk","")).replace('"',"'")
    try: score_f = float(score)
    except: return '<td class="n">—</td>'
    if grade == "wide":   bg, fg, label = "#1D3557", "#fff", "◎Wide"
    elif grade == "narrow": bg, fg, label = "#457B9D", "#fff", "○Narrow"
    else:                   bg, fg, label = "#e0e0e0", "#888", "△None"
    types_str = " ".join(f"[{t}]" for t in (types if isinstance(types, list) else []))
    tooltip   = f"{types_str} | {comment} | リスク:{risk}"[:200]
    # bonus値を★5段階で表示
    if score_f >= 25:   stars = "★★★★★"
    elif score_f >= 15: stars = "★★★★"
    elif score_f >= 8:  stars = "★★★"
    elif score_f >= 1:  stars = "★★"
    else:               stars = ""
    star_disp = f" {stars}" if stars else ""
    return (f'<td style="background:{bg};color:{fg};text-align:center;font-size:11px;'
            f'font-weight:600;white-space:nowrap;cursor:help" title="{tooltip}">'
            f'{label}{star_disp}</td>')

CSV_FILE    = os.path.join(SCRIPT_DIR, "via_results.csv")
OUTPUT_HTML = os.path.join(SCRIPT_DIR, "via_results.html")

# ── 基準設定（via_screener.pyと同じ値） ──────────────
WACC        = 8.0
DCF_DISC    = 10
DCF_MARGIN  = 30
DCF_ADJ_CAP = 12

STRICT = dict(
    pass_ratio  = 0.80,
    de_thr      = 50,
    roe_thr     = 15,
    roa_thr     = 7,
    roic_thr    = 15,
    consec_drop = 2,
    min_years   = 5,
    max_years   = 9,
)
RELAXED = dict(
    pass_ratio   = 0.70,
    de_thr       = 100,
    roe_thr      = 12,
    roa_thr      = 5,
    roic_thr     = 10,
    consec_drop  = 3,
    recent_years = 4,
)

# 判定対象の列（s_xxx / r_xxx）
CRITERIA_KEYS = [
    "EPS_positive", "EPS_uptrend", "GM_judge",
    "OCF_positive", "OCF_uptrend", "FCF_positive", "FCF_uptrend",
    "ROE_above_thr", "ROE_uptrend", "ROA_above_thr", "ROA_uptrend",
    "DE_below_thr", "ROIC_above_thr", "ROIC_vs_WACC",
]

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

def parse_bool(val):
    """CSV の True/False/None 文字列をPython boolに変換"""
    if pd.isna(val): return None
    if str(val).strip().lower() in ('true', '1'): return True
    if str(val).strip().lower() in ('false', '0'): return False
    return None

def recalc_scores(df):
    """
    s_xxx列が破損しているため、r_xxx列を使ってスコアを再計算する。
    r_xxx=Trueなら緩和基準通過 = 厳格基準も通過とみなす。
    （緩和閾値 <= 厳格閾値のため、緩和通過なら厳格通過の可能性が高い）
    ただしROE/ROA/D/E/ROICの閾値差があるため、
    s_xxxが壊れている場合はr_xxxをそのまま使う。
    """
    rows = []
    s_broken = df['s_score'].astype(str).str.strip() == '0'
    all_broken = s_broken.all()
    if all_broken:
        print("  ⚠️  s_xxx列が全て0で壊れています。r_xxx列で代替します。")

    for _, row in df.iterrows():
        # 緩和スコア（r_xxx）を計算
        r_vals = {}
        r_pass = 0; r_total = 0
        for k in CRITERIA_KEYS:
            v = parse_bool(row.get(f"r_{k}"))
            r_vals[k] = v
            if v is not None:
                r_total += 1
                if v is True: r_pass += 1

        # 厳格スコア: s_xxxが壊れていればr_xxxで代替
        s_vals = {}
        s_pass = 0; s_total = 0
        for k in CRITERIA_KEYS:
            sv = parse_bool(row.get(f"s_{k}"))
            if all_broken or sv is None:
                sv = r_vals.get(k)  # 壊れている場合はr_xxxで代替
            s_vals[k] = sv
            if sv is not None:
                s_total += 1
                if sv is True: s_pass += 1

        s_all = (s_pass == s_total and s_total > 0)
        r_all = (r_pass == r_total and r_total > 0)

        # データ年数不足チェック（CSVのdcf_years_back列を参照）
        # dcf_years_backはEPS起点までの年数（最大9年前）
        # EPSデータ年数はCSVに直接ないため dcf_present_eps の存在で代替
        eps_years_str = str(row.get("dcf_years_back",""))
        try:
            eps_years = int(float(eps_years_str)) + 1  # 起点年数+1 ≈ データ年数
        except:
            eps_years = 9  # 不明の場合は十分あるとみなす
        data_shortage = (eps_years < STRICT.get("min_years", 5))

        # 緩和で変わった項目
        relaxed_items = [k for k in CRITERIA_KEYS
                         if s_vals.get(k) is False and r_vals.get(k) is True]

        # grade判定（データ不足の場合は厳格→緩和に降格）
        if s_all and data_shortage:
            grade = "RELAXED"
            relaxed_items = [f"データ{eps_years}年（5年未満）"] + relaxed_items
        elif s_all:
            grade = "STRICT"
        elif r_all:
            grade = "RELAXED"
        else:
            grade = "FAIL"

        new_row = dict(row)
        new_row["s_score"]       = s_pass
        new_row["s_total"]       = s_total
        new_row["s_pct"]         = round(s_pass/s_total*100, 1) if s_total else 0
        new_row["r_score"]       = r_pass
        new_row["r_total"]       = r_total
        new_row["grade"]         = grade
        new_row["relaxed_items"] = ", ".join(relaxed_items)
        for k in CRITERIA_KEYS:
            new_row[f"s_{k}"] = s_vals[k]
            new_row[f"r_{k}"] = r_vals[k]
        rows.append(new_row)
    return pd.DataFrame(rows)

def _badge(grade, relaxed_items):
    if grade == "STRICT":
        return '<span class="badge b-strict">厳格通過</span>'
    return (f'<span class="badge b-relaxed" '
            f'title="緩和で通過した条件: {relaxed_items}">緩和通過</span>')

def _bool_cell(sv, rv):
    if sv is True:                return '<td class="p">✓</td>'
    if sv is False and rv is True:return '<td class="rx" title="緩和基準で通過">△</td>'
    if sv is False:               return '<td class="f">✗</td>'
    return '<td class="n">—</td>'

def _score_color(s, t):
    if not t: return "#888"
    r = s / t
    if r == 1.0: return "#1D9E75"
    if r >= 0.8: return "#5DCAA5"
    if r >= 0.6: return "#BA7517"
    return "#E24B4A"

def make_html(df_all, total_input, generated, moat_db=None):
    if moat_db is None: moat_db = {}
    df = df_all[df_all["grade"].isin(["STRICT","RELAXED"])].copy()

    # ソート
    grade_ord = {"STRICT":0,"RELAXED":1,"FAIL":2}
    df["_go"] = df["grade"].map(grade_ord).fillna(2)
    df["_uo"] = df["undervalued"].apply(
        lambda x: 0 if str(x).lower()=='true' else 1)
    df = df.sort_values(
        ["_uo","_go","s_score","s_pct"], ascending=[True,True,False,False]
    ).drop(columns=["_go","_uo"])

    n_s    = len(df[df["grade"]=="STRICT"])
    n_r    = len(df[df["grade"]=="RELAXED"])

    def is_uv(r): return str(r.get("undervalued","")).lower()=="true"
    n_uv_s = len(df[(df["grade"]=="STRICT") & df.apply(is_uv, axis=1)])
    n_uv_r = len(df[(df["grade"]=="RELAXED") & df.apply(is_uv, axis=1)])
    nc = len(CRITERIA_LABELS)

    th = "".join(f'<th onclick="sortTable({5+i})">{v}</th>'
                 for i, v in enumerate(CRITERIA_LABELS.values()))

    rows_html = []
    for _, row in df.iterrows():
        g   = row.get("grade","FAIL")
        ss  = int(row.get("s_score",0) or 0)
        st  = int(row.get("s_total",0) or 0)
        rs  = int(row.get("r_score",0) or 0)
        rt  = int(row.get("r_total",0) or 0)
        col = _score_color(ss, st)
        moat = calc_moat_score(row, moat_db)

        cells = "".join(
            _bool_cell(parse_bool(row.get(f"s_{k}")),
                       parse_bool(row.get(f"r_{k}")))
            for k in CRITERIA_KEYS
        )

        uv_bool = str(row.get("undervalued","")).lower()=="true"
        mp  = row.get("margin_pct")
        bp  = row.get("dcf_buy_price")
        itr = row.get("dcf_intrinsic")
        pr  = row.get("price")
        raw = row.get("dcf_raw_rate","")

        try: mp_f = float(mp)
        except: mp_f = None

        if uv_bool and mp_f is not None:
            uv_cell = f'<td class="uv-yes">割安 ({mp_f:+.1f}%)</td>'
        elif not uv_bool and mp_f is not None:
            uv_cell = f'<td class="uv-no">割高 ({mp_f:+.1f}%)</td>'
        else:
            uv_cell = '<td class="n">—</td>'

        score_disp = (f'{ss}/{st}'
                      + (f'<span class="r-score"> ({rs}/{rt})</span>'
                         if g=="RELAXED" else ""))

        bp_disp = (f'{bp} ✓' if bp and str(bp) not in ('','nan','None')
                   else '—')

        # 成長率色分け
        try: raw_f = float(raw)
        except: raw_f = 0

        rows_html.append(
            f'<tr data-grade="{g}" data-uv="{"1" if uv_bool else "0"}" data-moat="{moat.get("moat_grade","none")}">'
            f'<td>{row.get("market","")}</td>'
            f'<td><b>{row.get("ticker","")}</b></td>'
            f'<td title="{row.get("s_gm_note","")}">{str(row.get("name",""))[:28]}</td>'
            f'<td>{str(row.get("sector",""))[:20]}</td>'
            f'<td style="color:{col};font-weight:700">{score_disp}</td>'
            f'<td>{_badge(g, str(row.get("relaxed_items","")))}</td>'
            f'<td class="num">{pr or "—"}</td>'
            f'<td class="num">{bp_disp}</td>'
            f'<td class="num">{itr or "—"}</td>'
            f'{uv_cell}'
            f'{_asset_cell(row)}'
            f'{_moat_cell(moat)}'
            f'<td class="gr" data-rate="{raw_f}">{raw if raw and str(raw) not in ("","nan","None") else ""}</td>'
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
             f" / {STRICT['min_years']}〜{STRICT['max_years']}年で評価")
    r_cfg = (f"通過率{int(RELAXED['pass_ratio']*100)}% / D/E<={RELAXED['de_thr']}%"
             f" / ROE>={RELAXED['roe_thr']}% / ROA>={RELAXED['roa_thr']}%"
             f" / ROIC>={RELAXED['roic_thr']}% / 連続落込{RELAXED['consec_drop']}年"
             f" / 直近{RELAXED['recent_years']}年で評価")

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
.btn{{padding:5px 12px;border:1px solid #ccc;border-radius:6px;background:#fff;font-size:12px;cursor:pointer;color:#333}}
.btn:hover{{background:#f0f0e8}} .btn.active{{background:#185FA5;color:#fff;border-color:#185FA5}}
.leg{{font-size:11px;color:#888;margin-bottom:8px}}
.tbl-wrap{{overflow-x:auto}}
table{{border-collapse:collapse;width:100%;white-space:nowrap;background:#fff}}
th{{background:#f0f0e8;font-weight:500;padding:5px 7px;border-bottom:2px solid #ccc;position:sticky;top:0;font-size:12px;cursor:pointer;user-select:none}}
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
td.gr{{text-align:right;font-variant-numeric:tabular-nums}}
.badge{{display:inline-block;font-size:10px;padding:1px 7px;border-radius:10px;font-weight:600}}
.b-strict{{background:#e1f5ee;color:#0F6E56;border:1px solid #1D9E75}}
.b-relaxed{{background:#fef3e2;color:#854F0B;border:1px solid #BA7517}}
.r-score{{font-size:10px;color:#888}}
.moat-leg{{font-size:11px;color:#555;margin-bottom:8px;display:flex;flex-wrap:wrap;align-items:center;gap:4px;background:#f8f8f4;border:1px solid #e0e0d8;border-radius:6px;padding:6px 10px}}
.moat-badge{{display:inline-block;font-size:10px;padding:1px 7px;border-radius:10px;font-weight:600;white-space:nowrap}}
.moat-desc{{font-size:10px;color:#888;margin-right:12px}}
</style></head><body>
<h1>VIA スクリーニング結果 <a href="via_netnet_results.html" style="font-size:13px;font-weight:500;color:#185FA5;margin-left:12px;text-decoration:none;border:1px solid #185FA5;border-radius:6px;padding:3px 10px;vertical-align:middle">📊 ネットネット株スクリーニングはこちら</a></h1>
<p class="meta">生成: {generated} ／ WACC: {WACC}% ／ DCF割引率: {DCF_DISC}% ／ 安全領域: {DCF_MARGIN}% ／ 対象: {total_input}銘柄</p>
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
  <span style="color:#ccc;margin:0 4px">｜</span>
  <button class="btn"        id="btn-wide"        onclick="setFilter('wide')">◎Wide</button>
  <button class="btn"        id="btn-narrow"      onclick="setFilter('narrow')">○Narrow</button>
  <button class="btn"        id="btn-none"        onclick="setFilter('none')">△None</button>
</div>
<p class="leg">✓厳格通過 ／ △緩和で通過（橙色） ／ ✗NG ／ スコア: 厳格スコア (緩和スコア) ／ 列ヘッダーで並び替え</p>
<div class="moat-leg">
  <span style="font-weight:600;font-size:11px;color:#555;margin-right:8px">モート:</span>
  <span class="moat-badge" style="background:#1D3557;color:#fff">◎Wide</span>
  <span class="moat-desc">強力な競争優位（ネットワーク効果・業界標準・無形資産など）</span>
  <span class="moat-badge" style="background:#457B9D;color:#fff">○Narrow</span>
  <span class="moat-desc">一定の競争優位（スイッチングコスト・ブランド・規模の経済など）</span>
  <span class="moat-badge" style="background:#e0e0e0;color:#888">△None</span>
  <span class="moat-desc">競争優位が薄い　／　★=モートの強さ（★★★★★強〜★★弱）　／　セルにカーソルでコメント・リスク表示</span>
</div>
<div class="tbl-wrap"><table id="tbl">
<thead><tr>
  <th onclick="sortTable(0)">市場</th><th onclick="sortTable(1)">Ticker</th>
  <th onclick="sortTable(2)">名称</th><th onclick="sortTable(3)">セクター</th>
  <th onclick="sortTable(4)">スコア</th><th onclick="sortTable(5)">判定</th>
  <th onclick="sortTable(6)">現在株価</th><th onclick="sortTable(7)">購買ターゲット価格</th>
  <th onclick="sortTable(8)">正味現在価値</th><th onclick="sortTable(9)">割安判定</th>
  <th onclick="sortTable(10)">モート</th>
  <th onclick="sortTable(11)">成長率%</th>
  <th onclick="sortTable(11)">EPS</th><th onclick="sortTable(12)">ROE%</th>
  <th onclick="sortTable(13)">ROA%</th><th onclick="sortTable(14)">ROIC%</th>
  <th onclick="sortTable(15)">D/E%</th>
  {th}
</tr></thead>
<tbody>{"".join(rows_html)}</tbody>
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
    var mt=r.dataset.moat||'';
    if(cur==='wide')       fm=mt==='wide';
    if(cur==='narrow')     fm=mt==='narrow';
    if(cur==='none')       fm=mt==='none';
    r.style.display=(tm&&fm)?'':'none';
  }});
}}
function setFilter(f){{
  cur=f;
  ['all','uv','uv-strict','uv-relaxed','strict','relaxed','wide','narrow','none'].forEach(function(id){{
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


def main():
    print("VIA HTMLリビルド開始")
    print(f"CSV読み込み: {CSV_FILE}")

    # モートデータ読み込み
    moat_path = os.path.join(SCRIPT_DIR, "via_moat.json")
    MOAT_DATA = {}
    if os.path.exists(moat_path):
        import json as _jm
        with open(moat_path,"r",encoding="utf-8") as f:
            MOAT_DATA = _jm.load(f)
        print(f"モートデータ: {len(MOAT_DATA)}銘柄")

    if not os.path.exists(CSV_FILE):
        print(f"❌ {CSV_FILE} が見つかりません")
        return

    df = pd.read_csv(CSV_FILE, dtype=str, encoding="utf-8-sig")
    print(f"読み込み完了: {len(df)} 行")

    print("スコア再計算中...")
    df = recalc_scores(df)

    strict  = df[df["grade"]=="STRICT"]
    relaxed = df[df["grade"]=="RELAXED"]
    print(f"厳格通過: {len(strict)}")
    print(f"緩和通過: {len(relaxed)}")

    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    total = len(df)
    html = make_html(df, total, generated, MOAT_DATA)

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ HTML再生成完了: {OUTPUT_HTML}")

if __name__ == "__main__":
    main()
