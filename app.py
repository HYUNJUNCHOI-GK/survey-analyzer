# -*- coding: utf-8 -*-
"""
교육 만족도 설문 분석 웹앱
배포: Streamlit Community Cloud
"""

import io
import os
import statistics
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ── 페이지 설정 ──────────────────────────────────────────────────
st.set_page_config(
    page_title="교육 만족도 분석",
    page_icon="📊",
    layout="wide",
)

# ── API 키: Secrets → 환경변수 → 사용자 입력 순으로 탐색 ─────────
def get_server_api_key():
    """플랫폼 Secrets 또는 환경변수에 설정된 키 반환 (없으면 None)"""
    try:
        key = st.secrets.get("ANTHROPIC_API_KEY", "")
        if key:
            return key
    except Exception:
        pass
    return os.environ.get("ANTHROPIC_API_KEY", "") or None


SERVER_KEY = get_server_api_key()
HAS_SERVER_KEY = bool(SERVER_KEY)

# ── 사이드바 ─────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 설정")

    if HAS_SERVER_KEY:
        # 관리자가 서버에 키를 설정한 경우 → 사용자에게 숨김
        api_key = SERVER_KEY
        st.success("✅ AI 분석 사용 가능")
        use_ai = st.toggle("AI 요약 분석 사용", value=True)
    else:
        # 키 미설정 → 사용자가 직접 입력
        st.info("AI 분석을 사용하려면 Anthropic API Key를 입력하세요.")
        api_key_input = st.text_input(
            "Anthropic API Key",
            type="password",
            placeholder="sk-ant-...",
        )
        api_key = api_key_input or None
        use_ai = st.toggle("AI 요약 분석 사용", value=bool(api_key))

    st.divider()
    st.markdown("""
**파일 형식 안내**
- Google Forms 내보내기 (.xlsx)
- 컬럼 구조:
  - A: 타임스탬프
  - B: 성함, C: 팀명
  - D~F: 주관식 응답
  - G열~: 객관식 점수 (1~5점)
""")

# ── 유틸 ─────────────────────────────────────────────────────────
_EMPTY = {
    "", "없음", "없습니다", "없습니다.", "없어요", "없어요.",
    "특별히 없습니다", "특별히 없어요", "지금으로도 좋습니다",
    "지금으로도 좋습니다.", "현재로 충분합니다", "n/a", "na", "-", ".",
    "특이사항 없음", "수고하셨습니다", "수고하셨습니다.",
    "고생 하셨습니다.", "고생하셨습니다", "고생 하셨습니다",
}


def load_excel(file_bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    return rows[0], rows[1:]


def get_subj(data, col):
    return [
        str(r[col]).strip() for r in data
        if len(r) > col and r[col]
        and str(r[col]).strip().lower() not in _EMPTY
    ]


def score_stats(data, start=6):
    if not data:
        return []
    n_cols = max(len(r) for r in data)
    results = []
    for ci in range(start, n_cols):
        scores = [r[ci] for r in data if len(r) > ci and isinstance(r[ci], (int, float))]
        if scores:
            results.append({
                "col": ci,
                "avg": round(statistics.mean(scores), 2),
                "stdev": round(statistics.stdev(scores), 2) if len(scores) > 1 else 0.0,
                "min": min(scores),
                "max": max(scores),
                "scores": scores,
                "count": len(scores),
            })
    return results


def call_claude(course_name, good, improve, other, api_key):
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    parts = []
    if good:
        parts.append("[좋은 점]\n" + "\n".join(f"- {p}" for p in good))
    if improve:
        parts.append("[개선 점]\n" + "\n".join(f"- {p}" for p in improve))
    if other:
        parts.append("[기타 의견]\n" + "\n".join(f"- {p}" for p in other))
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


# ── 차트 함수 ─────────────────────────────────────────────────────
def short_label(h, max_len=22):
    s = h.split(". ", 1)[1] if ". " in h else h
    return s[:max_len] + "…" if len(s) > max_len else s


def make_bar(headers, stats_list):
    labels, avgs, colors = [], [], []
    for s in stats_list:
        h = headers[s["col"]] if s["col"] < len(headers) else f"Q{s['col']}"
        labels.append(short_label(h))
        avgs.append(s["avg"])
        colors.append("#27ae60" if s["avg"] >= 4.5 else ("#f39c12" if s["avg"] >= 3.5 else "#e74c3c"))
    fig = go.Figure(go.Bar(
        x=avgs, y=labels, orientation="h",
        marker_color=colors,
        text=[f"{v:.2f}" for v in avgs], textposition="outside",
    ))
    fig.update_layout(
        xaxis=dict(range=[0, 5.8], title="평균 점수 (5점 만점)"),
        yaxis=dict(autorange="reversed"),
        margin=dict(l=20, r=60, t=10, b=30),
        height=max(300, len(labels) * 40),
    )
    return fig


def make_radar(headers, stats_list):
    labels = [short_label(headers[s["col"]] if s["col"] < len(headers) else f"Q{s['col']}", 18) for s in stats_list]
    vals = [s["avg"] for s in stats_list]
    fig = go.Figure(go.Scatterpolar(
        r=vals + [vals[0]], theta=labels + [labels[0]],
        fill="toself", line_color="#3498db", fillcolor="rgba(52,152,219,0.2)",
    ))
    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 5])),
        showlegend=False,
        margin=dict(l=60, r=60, t=30, b=30),
        height=430,
    )
    return fig


# ── 보고서 생성 ───────────────────────────────────────────────────
def generate_report(course_name, respondents, headers, stats_list, good, improve, other, ai_summary):
    sep = "=" * 65
    thin = "-" * 65
    lines = [sep, "  만족도 분석 보고서",
             f"  과정명 : {course_name}",
             f"  분석일 : {datetime.now().strftime('%Y-%m-%d %H:%M')}",
             f"  응답자 : {respondents}명", sep, "",
             "【 객관식 항목별 평균 점수 (5점 만점) 】", thin]

    cat_defs = [("교육 내용", 0, 3), ("시간 배분", 3, 4),
                ("강사", 4, 8), ("교육 효과", 8, 11), ("운영 환경", 11, 13)]
    all_avgs = []
    for cat, s, e in cat_defs:
        items = stats_list[s:e]
        if not items:
            continue
        lines.append(f"\n  ▶ {cat}")
        for r in items:
            h = headers[r["col"]] if r["col"] < len(headers) else ""
            lbl = h.split(". ", 1)[1] if ". " in h else h
            bar = "█" * round(r["avg"]) + "░" * (5 - round(r["avg"]))
            lines.append(f"    {r['avg']:.2f}  {bar}  {lbl}")
            all_avgs.append(r["avg"])
    if all_avgs:
        lines += ["", thin, f"  전체 평균: {statistics.mean(all_avgs):.2f} / 5.00", sep, ""]

    lines += ["【 주관식 응답 원문 】", thin]
    if good:
        lines.append("\n  [좋은 점]"); lines += [f"    · {p}" for p in good]
    if improve:
        lines.append("\n  [개선 점]"); lines += [f"    · {p}" for p in improve]
    if other:
        lines.append("\n  [기타 의견]"); lines += [f"    · {p}" for p in other]
    lines += ["", sep, "【 AI 분석 요약 】", thin, ai_summary, "", sep]
    return "\n".join(lines)


# ── 메인 UI ──────────────────────────────────────────────────────
st.title("📊 교육 만족도 설문 분석")
st.caption("Google Forms Excel 파일을 업로드하면 자동으로 분석 보고서를 생성합니다.")

uploaded = st.file_uploader("📂 설문 Excel 파일 업로드 (.xlsx)", type=["xlsx"])
course_name = st.text_input("과정명", placeholder="예: AI 기반 업무방식 전환 실무 (AI-DLC 입문)")

if uploaded and course_name:
    if st.button("🔍 분석 시작", type="primary", use_container_width=True):

        with st.spinner("데이터 불러오는 중…"):
            headers, data = load_excel(uploaded.read())
            respondents = len(data)

        stats_list = score_stats(data)
        good = get_subj(data, 3)
        improve = get_subj(data, 4)
        other = get_subj(data, 5)

        ai_summary = "(AI 분석 생략)"
        if use_ai:
            if not api_key:
                ai_summary = "⚠️ API Key를 입력하면 AI 분석을 사용할 수 있습니다."
            else:
                with st.spinner("Claude AI 분석 중… (10~20초 소요)"):
                    try:
                        ai_summary = call_claude(course_name, good, improve, other, api_key)
                    except Exception as e:
                        ai_summary = f"⚠️ AI 분석 실패: {e}"

        # ── 결과 ─────────────────────────────────────────────────
        st.success(f"✅ 분석 완료 — 응답자 {respondents}명")
        st.divider()

        all_avgs = [s["avg"] for s in stats_list]
        overall = round(statistics.mean(all_avgs), 2) if all_avgs else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("전체 평균", f"{overall:.2f} / 5.00")
        c2.metric("응답자 수", f"{respondents}명")
        c3.metric("최고 항목", f"{max(all_avgs):.2f}" if all_avgs else "-")
        c4.metric("최저 항목", f"{min(all_avgs):.2f}" if all_avgs else "-")
        st.divider()

        tab1, tab2, tab3, tab4 = st.tabs(["📊 막대 차트", "🕸️ 방사형 차트", "💬 주관식 응답", "🤖 AI 분석 요약"])

        with tab1:
            st.plotly_chart(make_bar(headers, stats_list), use_container_width=True)
            rows_tbl = []
            for s in stats_list:
                h = headers[s["col"]] if s["col"] < len(headers) else ""
                rows_tbl.append({"문항": h, "평균": s["avg"], "표준편차": s["stdev"],
                                  "최저": s["min"], "최고": s["max"], "응답수": s["count"]})
            st.dataframe(pd.DataFrame(rows_tbl), use_container_width=True, hide_index=True)

        with tab2:
            st.plotly_chart(make_radar(headers, stats_list), use_container_width=True)

        with tab3:
            cols = st.columns(3)
            for col, title, items in zip(cols,
                    ["👍 좋은 점", "🔧 개선 점", "💡 기타 의견"],
                    [good, improve, other]):
                with col:
                    st.subheader(title)
                    if items:
                        for p in items:
                            st.markdown(f"- {p}")
                    else:
                        st.caption("(응답 없음)")

        with tab4:
            if ai_summary.startswith("⚠️") or ai_summary.startswith("("):
                st.info(ai_summary)
            else:
                st.markdown(ai_summary)

        st.divider()
        report_txt = generate_report(course_name, respondents, headers, stats_list,
                                     good, improve, other, ai_summary)
        st.download_button(
            "⬇️ 보고서 다운로드 (.txt)",
            data=report_txt.encode("utf-8"),
            file_name=f"{course_name}_분석보고서_{datetime.now().strftime('%Y%m%d')}.txt",
            mime="text/plain",
        )

elif uploaded and not course_name:
    st.info("과정명을 입력한 뒤 분석을 시작하세요.")
