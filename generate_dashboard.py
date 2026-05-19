import re
import json
import pandas as pd
from pathlib import Path

EXCEL_PATH         = Path(__file__).parent / "AS 현황_시디즈.xlsx"
OUTPUT_PATH        = Path(__file__).parent / "index.html"
TEMPLATE_PATH      = Path(__file__).parent / "AS_대시보드_template.html"
NOTIFICATIONS_PATH = Path(__file__).parent / "notifications.json"


# ── Excel readers ──────────────────────────────────────────────────────────────

def read_as_national():
    df = pd.read_excel(EXCEL_PATH, sheet_name="2. 현황_AS", header=None)
    data = df.iloc[2:, :7].copy()
    data.columns = ["조치월", "진행건수", "N차건", "N차율", "1차종결건수", "1차종결율", "평균조치기간"]
    data = data.dropna(subset=["1차종결건수"]).reset_index(drop=True)
    if data.empty:
        return data
    if data["조치월"].isna().all():
        data["조치월"] = data.index.map(lambda i: f"#{i+1}")
    else:
        data = data.dropna(subset=["조치월"])
    data["조치월"] = data["조치월"].astype(str)
    for col in ["진행건수", "N차건", "1차종결건수"]:
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0).astype(int)
    for col in ["N차율", "1차종결율", "평균조치기간"]:
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)
    return data.reset_index(drop=True)


def read_as_seoul():
    df = pd.read_excel(EXCEL_PATH, sheet_name="4. 현황_AS_서울경인", header=None)
    data = df.iloc[2:, :9].copy()
    data.columns = ["조치월", "진행건수", "N차건", "N차율", "1차종결건수", "1차종결율", "평균조치기간", "엔지니어비율", "시공팀비율"]
    data = data.dropna(subset=["1차종결건수"]).reset_index(drop=True)
    if data.empty:
        return data
    if data["조치월"].isna().all():
        data["조치월"] = data.index.map(lambda i: f"#{i+1}")
    else:
        data = data.dropna(subset=["조치월"])
    data["조치월"] = data["조치월"].astype(str)
    for col in ["진행건수", "N차건", "1차종결건수"]:
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0).astype(int)
    for col in ["N차율", "1차종결율", "평균조치기간", "엔지니어비율", "시공팀비율"]:
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)
    return data.reset_index(drop=True)


def read_contact_cost():
    df = pd.read_excel(EXCEL_PATH, sheet_name="AS 페이지, 컨택센터", header=None)
    months = ["1월", "2월", "3월", "4월"]
    items = []
    for row_idx in [4, 5, 6, 7]:
        row = df.iloc[row_idx]
        name = str(row[0])
        counts  = [row[4], row[6], row[8], row[10]]
        amounts = [row[5], row[7], row[9], row[11]]
        def _int(v):
            try:
                return int(float(v)) if pd.notna(v) else 0
            except (ValueError, TypeError):
                return 0
        items.append({
            "name": name,
            "counts":  [_int(v) for v in counts],
            "amounts": [_int(v) for v in amounts],
        })
    return {"months": months, "items": items}


def read_funnel():
    df = pd.read_excel(EXCEL_PATH, sheet_name="AS 페이지, 컨택센터", header=None)
    # 전체 열에서 기간 행을 읽어 데이터 범위 자동 파악
    periods_row = df.iloc[14, 1:].tolist()
    periods = [str(v) for v in periods_row if pd.notna(v)]
    n = len(periods)
    if n == 0:
        return {"periods": [], "제품등록수": [], "메인조회": [], "신청클릭": [],
                "제품선택": [], "증상입력": [], "고객정보입력": [], "AS신청완료": [],
                "1→2유입률": [], "3→4유입률": [], "4→5유입률": []}

    def _ints(row_idx):
        return [int(v) if pd.notna(v) else 0 for v in df.iloc[row_idx, 1:n+1].tolist()]

    reg    = _ints(15)
    view   = _ints(16)
    click  = _ints(17)
    select = _ints(18)
    sym    = _ints(19)
    info   = _ints(20)
    done   = _ints(21)

    # 유입률은 건수로 직접 계산 (Excel 유입률 행 무시 → 과거 데이터 호환)
    conv12 = [round(click[i] / view[i]   * 100, 1) if view[i]   else 0 for i in range(n)]
    conv34 = [round(sym[i]   / select[i] * 100, 1) if select[i] else 0 for i in range(n)]
    conv45 = [round(info[i]  / sym[i]    * 100, 1) if sym[i]    else 0 for i in range(n)]

    return {
        "periods": periods, "제품등록수": reg, "메인조회": view, "신청클릭": click,
        "제품선택": select, "증상입력": sym, "고객정보입력": info, "AS신청완료": done,
        "1→2유입률": conv12, "3→4유입률": conv34, "4→5유입률": conv45,
    }


def read_cs_cost():
    df = pd.read_excel(EXCEL_PATH, sheet_name="CS용역료", header=None)
    months = ["2월", "3월", "4월"]
    labels = ["유선AS", "유선온라인", "모바일", "게시판", "접수AS",
              "오프라인등록", "온라인등록", "오프라인변경", "온라인변경"]
    items = []
    for i, row_idx in enumerate(range(4, 13)):
        row = df.iloc[row_idx]
        unit = float(row[5]) if pd.notna(row[5]) else 0
        cnt  = [int(row[c]) if pd.notna(row[c]) else 0 for c in [6, 8, 10]]
        def _amt(col_idx, c):
            v = row[col_idx] if col_idx < len(row) else None
            return float(v) if (v is not None and pd.notna(v)) else c * unit
        amounts = [_amt(7, cnt[0]), _amt(9, cnt[1]), _amt(11, cnt[2])]
        items.append({"name": labels[i], "counts": cnt, "amounts": amounts})
    totals    = [sum(it["amounts"][m] for it in items) for m in range(3)]
    biz_days  = []
    for col in [6, 8, 10]:
        v = df.iloc[3, col] if col < len(df.columns) else None
        biz_days.append(int(v) if (v is not None and pd.notna(v)) else 1)
    daily_avg = [t / d if d else 0 for t, d in zip(totals, biz_days)]
    return {"months": months, "items": items, "totals": totals, "daily_avg": daily_avg}


# ── Helpers ────────────────────────────────────────────────────────────────────

def pct(v):
    return round(float(v) * 100, 1)


def _trend(last, prev):
    if prev == 0:
        return 0.0
    return round((float(last) - float(prev)) / abs(float(prev)) * 100, 1)


def _build_yoy(nat):
    """Group nat data into year→[12 monthly slots] dict for YoY chart."""
    MONTH_LABELS = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"]
    metrics = {"count": "진행건수", "close": "1차종결율", "n": "N차율", "days": "평균조치기간"}
    result = {k: {} for k in metrics}

    for _, row in nat.iterrows():
        m_str = str(row["조치월"])  # e.g. "2025-12" or "2025.12" or "25-12"
        year, mon = None, None
        for fmt in [r"(\d{4})[.\-](\d{1,2})", r"(\d{2})[.\-](\d{2})"]:
            m = re.match(fmt, m_str)
            if m:
                y, mo = int(m.group(1)), int(m.group(2))
                year = (2000 + y) if y < 100 else y
                mon  = mo
                break
        if year is None or mon is None or not (1 <= mon <= 12):
            continue
        yr = str(year)
        for k, col in metrics.items():
            if yr not in result[k]:
                result[k][yr] = [None] * 12
            v = float(row[col])
            if col == "1차종결율" or col == "N차율":
                v = pct(v)
            else:
                v = round(v, 2) if col == "평균조치기간" else int(v)
            result[k][yr][mon - 1] = v

    return MONTH_LABELS, result


# ── HTML builder ───────────────────────────────────────────────────────────────

def build_html(nat, seoul, contact, funnel, cs):
    # ── Align seoul to nat length ──────────────────────────────────────────────
    n = len(nat)
    s = len(seoul)
    # If seoul has fewer rows than nat, pad with nulls (JS null-safe charts)
    def _align(lst, target_len, fill=None):
        return lst[:target_len] if len(lst) >= target_len else lst + [fill] * (target_len - len(lst))

    nat_months   = nat["조치월"].tolist()
    nat_count    = nat["진행건수"].tolist()
    nat_close    = [pct(v) for v in nat["1차종결율"]]
    nat_n        = [pct(v) for v in nat["N차율"]]
    nat_days     = [round(float(v), 2) for v in nat["평균조치기간"]]

    se_count     = _align(seoul["진행건수"].tolist(),            n)
    se_close     = _align([pct(v) for v in seoul["1차종결율"]], n)
    se_n         = _align([pct(v) for v in seoul["N차율"]],     n)
    se_days      = _align([round(float(v), 2) for v in seoul["평균조치기간"]], n)
    se_eng       = _align([pct(v) for v in seoul["엔지니어비율"]], n)
    se_sig       = _align([pct(v) for v in seoul["시공팀비율"]],  n)

    yoy_months, yoy = _build_yoy(nat)

    # ── cs items: design needs {name, amounts} only ────────────────────────────
    cs_items_for_js = [{"name": it["name"], "amounts": it["amounts"]} for it in cs["items"]]

    # ── funnel mapping ─────────────────────────────────────────────────────────
    f = funnel
    funnel_js = {
        "periods":    f["periods"],
        "productReg": [int(v) for v in f["제품등록수"]],
        "main":       [int(v) for v in f["메인조회"]],
        "click":      [int(v) for v in f["신청클릭"]],
        "pick":       [int(v) for v in f["제품선택"]],
        "sym":        [int(v) for v in f["증상입력"]],
        "info":       [int(v) for v in f["고객정보입력"]],
        "done":       [int(v) for v in f["AS신청완료"]],
        "conv12":     f["1→2유입률"],
        "conv34":     f["3→4유입률"],
        "conv45":     f["4→5유입률"],
    }

    # ── DATA JS block ──────────────────────────────────────────────────────────
    data_obj = {
        "natMonths":    nat_months,
        "natCount":     nat_count,
        "natClose":     nat_close,
        "natN":         nat_n,
        "natDays":      nat_days,
        "seCount":      se_count,
        "seClose":      se_close,
        "seN":          se_n,
        "seDays":       se_days,
        "seEng":        se_eng,
        "seSig":        se_sig,
        "contactMonths": contact["months"],
        "contact":      contact["items"],
        "csMonths":     cs["months"],
        "csTotals":     cs["totals"],
        "csDaily":      cs["daily_avg"],
        "cs":           cs_items_for_js,
        "funnel":       funnel_js,
        "yoyMonths":    yoy_months,
        "yoy":          yoy,
    }
    data_js_block = "const DATA = " + json.dumps(data_obj, ensure_ascii=False, indent=2) + ";"

    # ── KPIS JS block (references DATA fields inline) ──────────────────────────
    _empty = {"진행건수": 0, "N차율": 0, "1차종결율": 0, "평균조치기간": 0,
              "엔지니어비율": 0, "시공팀비율": 0}
    import pandas as _pd
    ln  = nat.iloc[-1]   if n >= 1 else _pd.Series(_empty)
    pn  = nat.iloc[-2]   if n >= 2 else _pd.Series(_empty)
    ls  = seoul.iloc[-1] if s >= 1 else _pd.Series(_empty)
    ps  = seoul.iloc[-2] if s >= 2 else _pd.Series(_empty)

    def _kpi_dir_count(t): return "up"  if t >= 0 else "down"  # higher is better
    def _kpi_dir_low(t):   return "up"  if t <= 0 else "down"  # lower is better

    t_nc  = _trend(ln["진행건수"],    pn["진행건수"])
    t_cl  = _trend(pct(ln["1차종결율"]), pct(pn["1차종결율"]))
    t_n   = _trend(pct(ln["N차율"]),  pct(pn["N차율"]))
    t_d   = _trend(float(ln["평균조치기간"]), float(pn["평균조치기간"]))
    t_sc  = _trend(ls["진행건수"],    ps["진행건수"])
    t_scl = _trend(pct(ls["1차종결율"]), pct(ps["1차종결율"]))
    t_sn  = _trend(pct(ls["N차율"]),  pct(ps["N차율"]))
    t_sd  = _trend(float(ls["평균조치기간"]), float(ps["평균조치기간"]))

    latest_month = nat_months[-1] if nat_months else (seoul["조치월"].tolist()[-1] if s >= 1 else "-")

    kpis_js = f"""const KPIS = [
  {{label: "전국 조치 건수",       value: {int(ln["진행건수"])},  unit: "건",  trend: {t_nc},  dir: "{_kpi_dir_count(t_nc)}", comp: "전월 대비",                   series: DATA.natCount,  cls: "is-accent"}},
  {{label: "전국 1차 종결률",      value: {pct(ln["1차종결율"])}, unit: "%",   trend: {t_cl},  dir: "{_kpi_dir_count(t_cl)}", comp: "전월 대비",                   series: DATA.natClose,  cls: "is-pos"}},
  {{label: "전국 N차율",           value: {pct(ln["N차율"])},     unit: "%",   trend: {t_n},   dir: "{_kpi_dir_low(t_n)}",    comp: "전월 대비 (낮을수록 좋음)",  series: DATA.natN,      cls: "is-warn"}},
  {{label: "전국 평균 조치기간",   value: {round(float(ln["평균조치기간"]),2)}, unit: "일", trend: {t_d}, dir: "{_kpi_dir_low(t_d)}", comp: "전월 대비 (낮을수록 좋음)", series: DATA.natDays,  cls: "is-pos"}},
  {{label: "서울경인 조치 건수",   value: {int(ls["진행건수"])},  unit: "건",  trend: {t_sc},  dir: "{_kpi_dir_count(t_sc)}", comp: "전월 대비",                   series: DATA.seCount,   cls: ""}},
  {{label: "서울경인 1차 종결률",  value: {pct(ls["1차종결율"])}, unit: "%",   trend: {t_scl}, dir: "{_kpi_dir_count(t_scl)}",comp: "전월 대비",                   series: DATA.seClose,   cls: "is-pos"}},
  {{label: "서울경인 N차율",       value: {pct(ls["N차율"])},     unit: "%",   trend: {t_sn},  dir: "{_kpi_dir_low(t_sn)}",   comp: "전월 대비 (낮을수록 좋음)",  series: DATA.seN,       cls: "is-warn"}},
  {{label: "서울경인 평균 조치",   value: {round(float(ls["평균조치기간"]),2)}, unit: "일", trend: {t_sd}, dir: "{_kpi_dir_low(t_sd)}", comp: "전월 대비 (낮을수록 좋음)", series: DATA.seDays, cls: ""}},
];"""

    # ── Notifications ─────────────────────────────────────────────────────────
    notifs = json.loads(NOTIFICATIONS_PATH.read_text(encoding="utf-8")) if NOTIFICATIONS_PATH.exists() else []
    notif_html = ""
    for n in notifs:
        dot_cls = "notif-dot" if not n.get("read") else "notif-dot read"
        notif_html += (
            f'<div class="notif-item">'
            f'<div class="{dot_cls}"></div>'
            f'<div><div class="notif-text">{n["text"]}</div>'
            f'<div class="notif-time">{n["time"]}</div></div>'
            f'</div>\n'
        )
    notif_badge = sum(1 for n in notifs if not n.get("read"))

    # ── Read template and inject ───────────────────────────────────────────────
    tmpl = TEMPLATE_PATH.read_text(encoding="utf-8")

    # Replace the DATA block
    tmpl = re.sub(
        r'const DATA = \{[\s\S]*?\n\};',
        data_js_block,
        tmpl
    )

    # Replace the KPIS block
    tmpl = re.sub(
        r'const KPIS = \[[\s\S]*?\n\];',
        kpis_js,
        tmpl
    )

    # Update the breadcrumb month badge in the sidebar
    tmpl = re.sub(
        r'(<span[^>]*id="crumb-month"[^>]*>)[^<]*(</span>)',
        rf'\g<1>{latest_month}\g<2>',
        tmpl
    )

    # Inject notifications into panel
    tmpl = re.sub(
        r'(<div class="ui-panel-body" id="notif-body">)<!-- NOTIF_PLACEHOLDER -->(</div>)',
        rf'\g<1>\n{notif_html}\g<2>',
        tmpl
    )

    # Update notification badge dot visibility
    if notif_badge > 0:
        tmpl = tmpl.replace(
            '<span class="tb-dot"></span>',
            '<span class="tb-dot" style="background:var(--accent)"></span>'
        )
    else:
        tmpl = tmpl.replace('<span class="tb-dot"></span>', '')

    return tmpl


if __name__ == "__main__":
    print("데이터 읽는 중...")
    nat     = read_as_national()
    seoul   = read_as_seoul()
    contact = read_contact_cost()
    funnel  = read_funnel()
    cs      = read_cs_cost()
    print(f"전국: {len(nat)}개월 / 서울경인: {len(seoul)}개월")
    html = build_html(nat, seoul, contact, funnel, cs)
    OUTPUT_PATH.write_text(html, encoding="utf-8")
    print(f"완료: {OUTPUT_PATH}")
