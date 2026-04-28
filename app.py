# -*- coding: utf-8 -*-
"""
교육 만족도 설문 분석 웹앱 (Streamlit)
실행: streamlit run survey_web.py
"""

import io
import os
import re
import statistics
from collections import Counter
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="교육 만족도 분석", page_icon="📊", layout="wide")

# ── 접속 비밀번호 ────────────────────────────────────────────────
def _get_secret(key, fallback=""):
    try:
        return st.secrets.get(key, fallback)
    except Exception:
        return os.environ.get(key, fallback)

def check_password():
    correct = _get_secret("APP_PASSWORD")
    if not correct:
        return True  # 미설정 시 통과
    if st.session_state.get("authenticated"):
        return True
    st.title("🔐 교육 만족도 분석")
    with st.form("login_form"):
        pw = st.text_input("접속 비밀번호", type="password")
        submitted = st.form_submit_button("확인", type="primary", use_container_width=True)
    if submitted:
        if pw == correct:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    return False

if not check_password():
    st.stop()

# ── API 키: Secrets → 환경변수 → 사용자 입력 ─────────────────────
SERVER_KEY = _get_secret("ANTHROPIC_API_KEY")

with st.sidebar:
    st.header("⚙️ 설정")
    if SERVER_KEY:
        api_key = SERVER_KEY
        st.success("✅ AI 분석 사용 가능")
        use_ai = st.toggle("AI 요약 분석", value=True)
    else:
        st.info("AI 분석을 위해 API Key를 입력하세요.")
        api_key = st.text_input("Anthropic API Key", type="password", placeholder="sk-ant-...")
        use_ai = st.toggle("AI 요약 분석", value=bool(api_key))
    st.divider()
    st.markdown("**지원 파일 형식**\n- Google Forms 내보내기 (.xlsx)\n- 점수 컬럼(1~5점)과 주관식 컬럼을 자동으로 감지합니다.")

# ── 상수 ────────────────────────────────────────────────────────
_EMPTY = {
    "", "없음", "없습니다", "없습니다.", "없어요", "없어요.", "특별히 없습니다",
    "특별히 없어요", "지금으로도 좋습니다", "지금으로도 좋습니다.", "현재로 충분합니다",
    "n/a", "na", "-", ".", "특이사항 없음", "수고하셨습니다", "수고하셨습니다.",
    "고생 하셨습니다.", "고생하셨습니다", "고생 하셨습니다",
}

_KW_STOP = {
    # 조사
    '이', '가', '을', '를', '은', '는', '의', '에', '로', '으로', '와', '과', '도',
    '에서', '이고', '에게', '부터', '까지', '처럼', '만큼', '정도', '이라고', '이라는',
    # 지시/관계
    '이런', '저런', '그런', '같은', '다른', '어떤', '이번', '저번',
    # 동사/형용사 활용형 (불용)
    '있는', '있어', '있어요', '있습니다', '있었으면', '있으면', '있을',
    '하고', '하는', '한', '하여', '해서', '했습니다', '하면',
    '됩니다', '되면', '되어', '되었으면', '되었습니다',
    '같습니다', '같아요', '같았습니다', '같았어요', '것같습니다',
    '좋겠습니다', '좋을', '좋아요', '좋습니다',
    '싶습니다', '싶어요', '싶은',
    '해주세요', '주세요', '주시면', '주시기', '드립니다',
    '생각합니다', '생각해요', '생각됩니다',
    '것', '수', '더', '위한', '통해', '대한', '통한',
    # 교육 관련 일반 단어 (너무 포괄적)
    '시간', '교육', '강의', '과정', '수업', '내용', '부분',
}

_CAT_DEFS = [
    ("교육 내용", 0, 3),
    ("시간 배분", 3, 4),
    ("강사",     4, 8),
    ("교육 효과", 8, 11),
    ("운영 환경", 11, 13),
]

# ── 컬럼 자동 감지 ───────────────────────────────────────────────
_META_KEYWORDS = {
    'timestamp': ['타임스탬프', 'timestamp', '작성일', '제출일', '날짜', '일시', '시각'],
    'name':      ['이름', '성함', '성명', '이름을', '성함을', 'name'],
    'team':      ['팀', '부서', '소속', '팀명', '부서명', 'team', 'dept', '회사'],
}

def detect_columns(headers, data):
    """점수 컬럼과 주관식 컬럼을 데이터 값으로 자동 감지"""
    from datetime import datetime as _dt
    score_cols, subj_cols = [], []
    for i, h in enumerate(headers):
        if h is None:
            continue
        h_l = str(h).lower()

        # 메타데이터 컬럼(타임스탬프/이름/팀) 제외
        if any(k in h_l for keys in _META_KEYWORDS.values() for k in keys):
            continue

        vals = [r[i] for r in data if len(r) > i and r[i] is not None]
        if not vals:
            continue

        # datetime 객체 → 제외
        if any(isinstance(v, _dt) for v in vals):
            continue

        numeric = [v for v in vals if isinstance(v, (int, float))]

        # 순번 컬럼 제외: 모든 값이 고유하고 1부터 연속된 정수
        if numeric and len(numeric) == len(vals):
            sorted_n = sorted(numeric)
            if sorted_n == list(range(1, len(sorted_n) + 1)):
                continue

        # 70% 이상 숫자 + 범위 1-10 → 점수 컬럼 (1~5, 1~7, 1~10 모두 지원)
        if len(numeric) / len(vals) >= 0.7 and numeric and min(numeric) >= 1 and max(numeric) <= 10:
            score_cols.append(i)
            continue

        # 고유값 비율 낮으면 카테고리형 → 제외 (팀명, 직무 등 중복값)
        unique_ratio = len(set(str(v).strip() for v in vals)) / len(vals)
        if unique_ratio < 0.5 and len(vals) > 3:
            continue

        # 평균 길이 8자 이상 텍스트 → 주관식
        texts = [str(v).strip() for v in vals if str(v).strip().lower() not in _EMPTY]
        if texts and sum(len(t) for t in texts) / len(texts) >= 8:
            if any(k in h_l for k in ['좋은', '장점', '긍정', '강점', 'good']):
                label = "👍 좋은 점"
            elif any(k in h_l for k in ['개선', '단점', '부족', '불만', '아쉬', 'improve']):
                label = "🔧 개선 점"
            else:
                label = "💡 기타 의견"
            subj_cols.append((i, label, str(h)))
    return score_cols, subj_cols


def load_excel(file_bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    # 헤더 행 자동 탐지: 문자열 값이 3개 이상인 첫 번째 행
    header_idx = 0
    for i, row in enumerate(rows):
        text_cnt = sum(
            1 for c in row
            if isinstance(c, str) and len(c.strip()) > 1 and not c.startswith('=')
        )
        if text_cnt >= 3:
            header_idx = i
            break

    headers = rows[header_idx]

    # 데이터 행 필터링
    data = []
    for row in rows[header_idx + 1:]:
        # 수식·빈 값 제외 후 유효 셀 수 확인
        valid = [
            c for c in row
            if c is not None
            and str(c).strip() not in ('', )
            and not str(c).startswith('=')
        ]
        if len(valid) < 2:
            continue
        # 척도 표시 행 제외 ("1 ~ 7", "1~5" 형태)
        scale_cnt = sum(
            1 for c in row
            if isinstance(c, str) and re.search(r'\d+\s*[~～]\s*\d+', c)
        )
        if scale_cnt >= 1:
            continue
        data.append(row)

    return headers, data


def get_subj(data, col):
    return [
        str(r[col]).strip() for r in data
        if len(r) > col and r[col]
        and str(r[col]).strip().lower() not in _EMPTY
    ]


def score_stats(data, score_cols):
    results = []
    for ci in score_cols:
        scores = [r[ci] for r in data if len(r) > ci and isinstance(r[ci], (int, float))]
        if scores:
            results.append({
                "col": ci,
                "avg": round(statistics.mean(scores), 2),
                "stdev": round(statistics.stdev(scores), 2) if len(scores) > 1 else 0.0,
                "min": min(scores),
                "max": max(scores),
                "count": len(scores),
                "scores": scores,
            })
    return results


# ── 키워드 추출 + 워드클라우드 ──────────────────────────────────────
_KW_STOP_ENDINGS = (
    '습니다', '었으면', '겠습니다', '이에요', '예요', '아요', '어요',
    '같아요', '같습니다', '같았습니다', '것같습니다',
)

def _get_korean_font():
    """OS별 한국어 폰트 경로"""
    candidates = [
        '/usr/share/fonts/truetype/nanum/NanumGothic.ttf',   # Linux (Streamlit Cloud)
        '/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf',
        'C:/Windows/Fonts/malgun.ttf',                        # Windows 맑은고딕
        'C:/Windows/Fonts/gulim.ttc',                         # Windows 굴림
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def get_word_freq(texts):
    words = []
    for t in texts:
        for w in re.findall(r'[가-힣]{2,}', t):
            if w in _KW_STOP:
                continue
            if any(w.endswith(e) for e in _KW_STOP_ENDINGS):
                continue
            words.append(w)
    return Counter(words)


def highlight_text(text, keyword):
    """응답 텍스트에서 키워드를 노란 형광으로 강조"""
    import re
    pattern = re.compile(f'({re.escape(keyword)})', re.IGNORECASE)
    return pattern.sub(
        r'<mark style="background:#fff176;padding:1px 4px;border-radius:3px;font-weight:bold">\1</mark>',
        text
    )


# ── Claude API ───────────────────────────────────────────────────
def call_claude(course_name, subj_data, api_key):
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    parts = []
    for texts, label, _ in subj_data:
        if texts:
            clean_label = label.split(" ", 1)[1] if " " in label else label
            parts.append(f"[{clean_label}]\n" + "\n".join(f"- {p}" for p in texts))
    if not parts:
        return "주관식 응답이 없어 AI 분석을 생략합니다."

    prompt = f"""다음은 "{course_name}" 교육 과정의 만족도 설문 주관식 응답입니다.

{chr(10).join(parts)}

아래 세 항목을 각각 번호 리스트(2~3개)로 간결하게 작성해주세요.

## 주요 강점
(교육에서 잘된 점, 학습자가 높이 평가한 부분)

## 주요 개선 필요 사항
(반복적으로 언급된 불만, 우선순위 포함)

## 담당자 액션 아이템
(다음 과정 운영 시 바로 적용 가능한 구체적 조치)
"""
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}],
    )
    return msg.content[0].text


# ── 차트 ─────────────────────────────────────────────────────────
def short_q(h, n=28):
    s = h.split(". ", 1)[1] if ". " in str(h) else str(h)
    return s[:n] + "…" if len(s) > n else s


def _relative_color(avg, min_avg, max_avg):
    """데이터 범위 내 상대 위치로 색상 결정"""
    if max_avg == min_avg:
        return "#f39c12"
    pct = (avg - min_avg) / (max_avg - min_avg)
    if pct >= 0.7:
        return "#27ae60"   # 상위 30% → 초록
    elif pct <= 0.3:
        return "#e74c3c"   # 하위 30% → 빨강
    return "#f39c12"       # 중간 → 주황


def make_item_bar(headers, stats_list):
    """문항별 가로 막대 차트"""
    labels, avgs, colors = [], [], []
    all_avgs = [s["avg"] for s in stats_list]
    min_avg, max_avg = min(all_avgs), max(all_avgs)
    max_scale = max(s["max"] for s in stats_list)  # 실제 척도 최대값 (5점/7점 등)
    for s in stats_list:
        h = headers[s["col"]] if s["col"] < len(headers) else f"Q{s['col']}"
        labels.append(short_q(h))
        avgs.append(s["avg"])
        colors.append(_relative_color(s["avg"], min_avg, max_avg))
    fig = go.Figure(go.Bar(
        x=avgs, y=labels, orientation="h",
        marker_color=colors,
        text=[f"{v:.2f}" for v in avgs], textposition="outside",
    ))
    fig.update_layout(
        xaxis=dict(range=[0, max_scale + 0.8], title=f"평균 점수 ({max_scale}점 만점)"),
        yaxis=dict(autorange="reversed"),
        margin=dict(l=20, r=60, t=10, b=30),
        height=max(300, len(labels) * 40),
    )
    return fig


def make_category_bar(stats_list):
    """카테고리별 평균 점수 세로 막대 차트 (방사형 대체)"""
    cats, avgs, colors = [], [], []
    max_scale = max(s["max"] for s in stats_list) if stats_list else 5
    hi_thr = max_scale * 0.9   # 상위: 전체 척도의 90%
    lo_thr = max_scale * 0.7   # 하위: 전체 척도의 70%
    for cat_name, s, e in _CAT_DEFS:
        items = stats_list[s:e]
        if not items:
            continue
        avg = round(statistics.mean([i["avg"] for i in items]), 2)
        cats.append(cat_name)
        avgs.append(avg)
        colors.append("#27ae60" if avg >= hi_thr else ("#f39c12" if avg >= lo_thr else "#e74c3c"))
    if len(cats) < 2:   # 카테고리가 1개 이하면 의미 없음 → 숨김
        return None
    fig = go.Figure(go.Bar(
        x=cats, y=avgs,
        marker_color=colors,
        text=[f"{v:.2f}" for v in avgs], textposition="outside",
        width=0.5,
    ))
    fig.update_layout(
        yaxis=dict(range=[0, max_scale + 0.5], title="평균 점수"),
        margin=dict(l=20, r=20, t=20, b=30),
        height=300,
    )
    return fig


# ── 보고서 텍스트 생성 ───────────────────────────────────────────
def generate_report(course_name, respondents, headers, stats_list, subj_data, ai_summary):
    sep, thin = "=" * 65, "-" * 65
    lines = [sep, "  만족도 분석 보고서",
             f"  과정명 : {course_name}",
             f"  분석일 : {datetime.now().strftime('%Y-%m-%d %H:%M')}",
             f"  응답자 : {respondents}명", sep, "",
             "【 객관식 항목별 평균 점수 (5점 만점) 】", thin]

    all_avgs = []
    for cat, s, e in _CAT_DEFS:
        items = stats_list[s:e]
        if not items:
            continue
        lines.append(f"\n  ▶ {cat}")
        for r in items:
            h = headers[r["col"]] if r["col"] < len(headers) else ""
            lbl = h.split(". ", 1)[1] if ". " in str(h) else str(h)
            bar = "█" * round(r["avg"]) + "░" * (5 - round(r["avg"]))
            lines.append(f"    {r['avg']:.2f}  {bar}  {lbl}")
            all_avgs.append(r["avg"])

    if all_avgs:
        lines += ["", thin, f"  전체 평균: {statistics.mean(all_avgs):.2f} / 5.00", sep, ""]

    lines += ["【 주관식 응답 원문 】", thin]
    for texts, label, _ in subj_data:
        clean = label.split(" ", 1)[1] if " " in label else label
        if texts:
            lines.append(f"\n  [{clean}]")
            lines += [f"    · {p}" for p in texts]

    lines += ["", sep, "【 AI 분석 요약 】", thin, ai_summary, "", sep]
    return "\n".join(lines)


# ── 메인 UI ──────────────────────────────────────────────────────
st.title("📊 교육 만족도 설문 분석")
st.caption("Google Forms Excel 파일을 업로드하면 자동으로 분석 보고서를 생성합니다.")

with st.form("analyze_form"):
    uploaded = st.file_uploader(
        "📂 파일을 드래그하거나 클릭하여 업로드 (.xlsx)",
        type=["xlsx"],
        help="Google Forms에서 내보낸 Excel 파일을 올려주세요.",
    )
    course_name = st.text_input("과정명", placeholder="예: AI 기반 업무방식 전환 실무 (AI-DLC 입문)")
    form_submitted = st.form_submit_button("🔍 분석 시작", type="primary", use_container_width=True)

# ── 폼 제출 시 분석 실행 → session_state 저장 ─────────────────────
if form_submitted:
    if not uploaded:
        st.warning("파일을 먼저 업로드해주세요.")
    elif not course_name:
        st.warning("과정명을 입력해주세요.")
    else:
        with st.spinner("데이터 로드 및 컬럼 감지 중…"):
            headers, data = load_excel(uploaded.read())
            respondents = len(data)
            score_cols, subj_cols = detect_columns(headers, data)

        if not score_cols:
            st.error("점수 컬럼(1~5점 숫자)을 찾지 못했습니다. 파일 형식을 확인해주세요.")
            st.stop()

        stats_list = score_stats(data, score_cols)
        subj_data = [(get_subj(data, col), label, h) for col, label, h in subj_cols]

        ai_summary = "(AI 분석 생략)"
        if use_ai:
            if not api_key:
                ai_summary = "⚠️ API Key를 입력하면 AI 분석을 사용할 수 있습니다."
            else:
                with st.spinner("Claude AI 분석 중… (10~20초 소요)"):
                    try:
                        ai_summary = call_claude(course_name, subj_data, api_key)
                    except Exception as e:
                        ai_summary = f"⚠️ AI 분석 실패: {e}"

        st.session_state['_analysis'] = {
            'course_name': course_name,
            'respondents': respondents,
            'headers': list(headers),
            'stats_list': stats_list,
            'subj_data': subj_data,
            'subj_cols': subj_cols,
            'ai_summary': ai_summary,
        }

# ── 결과 표시 (session_state 기반 → pills 클릭 후에도 유지) ────────
if '_analysis' in st.session_state:
    r = st.session_state['_analysis']
    course_name  = r['course_name']
    respondents  = r['respondents']
    headers      = r['headers']
    stats_list   = r['stats_list']
    subj_data    = r['subj_data']
    subj_cols    = r['subj_cols']
    ai_summary   = r['ai_summary']

    st.success(f"✅ 분석 완료 — 응답자 {respondents}명 / 점수 문항 {len(stats_list)}개 / 주관식 {len(subj_cols)}개")
    st.divider()

    all_avgs = [s["avg"] for s in stats_list]
    overall  = round(statistics.mean(all_avgs), 2) if all_avgs else 0
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("전체 평균",  f"{overall:.2f} / 5.00")
    c2.metric("응답자 수",  f"{respondents}명")
    c3.metric("최고 항목",  f"{max(all_avgs):.2f}" if all_avgs else "-")
    c4.metric("최저 항목",  f"{min(all_avgs):.2f}" if all_avgs else "-")
    st.divider()

    # ── 1. 객관식 점수 ────────────────────────────────────────
    st.subheader("📊 객관식 항목별 평균 점수")
    cat_fig = make_category_bar(stats_list)
    if cat_fig:
        col_item, col_cat = st.columns([3, 2])
        with col_item:
            st.caption("문항별 점수")
            st.plotly_chart(make_item_bar(headers, stats_list), use_container_width=True)
        with col_cat:
            st.caption("카테고리별 평균")
            st.plotly_chart(cat_fig, use_container_width=True)
    else:
        st.plotly_chart(make_item_bar(headers, stats_list), use_container_width=True)

    with st.expander("📋 상세 점수 테이블", expanded=False):
        rows_tbl = [{"문항": headers[s["col"]] if s["col"] < len(headers) else "",
                     "평균": s["avg"], "표준편차": s["stdev"],
                     "최저": s["min"], "최고": s["max"], "응답수": s["count"]}
                    for s in stats_list]
        st.dataframe(pd.DataFrame(rows_tbl), use_container_width=True, hide_index=True)

    st.divider()

    # ── 2. 주관식 응답 + 키워드 ──────────────────────────────
    st.subheader("💬 주관식 응답")
    if not subj_cols:
        st.info("주관식 응답 컬럼을 찾지 못했습니다.")
    else:
        cols = st.columns(len(subj_cols))
        for i, (col, (texts, label, h)) in enumerate(zip(cols, subj_data)):
            with col:
                st.markdown(f"**{label}**")

                freq    = get_word_freq(texts)
                top_kws = [w for w, _ in freq.most_common(8)]
                selected_kw = None
                if top_kws:
                    fmt = {w: f"{w}  {freq[w]}" for w in top_kws}
                    selected_kw = st.pills(
                        "핵심 키워드",
                        options=top_kws,
                        format_func=lambda w, f=fmt: f[w],
                        key=f"pills_{i}",
                    )

                if texts:
                    for p in texts:
                        if selected_kw and selected_kw in p:
                            st.markdown(
                                f'- {highlight_text(p, selected_kw)}',
                                unsafe_allow_html=True,
                            )
                        else:
                            st.markdown(f"- {p}")
                else:
                    st.caption("(응답 없음)")

    st.divider()

    # ── 3. AI 분석 요약 ───────────────────────────────────────
    st.subheader("🤖 AI 분석 요약")
    if ai_summary.startswith("⚠️") or ai_summary.startswith("("):
        st.info(ai_summary)
    else:
        st.markdown(ai_summary)

    st.divider()

    # ── 4. 보고서 다운로드 ────────────────────────────────────
    report_txt = generate_report(course_name, respondents, headers, stats_list, subj_data, ai_summary)
    st.download_button(
        "⬇️ 보고서 다운로드 (.txt)",
        data=report_txt.encode("utf-8"),
        file_name=f"{course_name}_분석보고서_{datetime.now().strftime('%Y%m%d')}.txt",
        mime="text/plain",
    )
