import streamlit as st
import re, json, time, io
from pptx import Presentation
from pptx.util import Pt
import anthropic

# ── 페이지 설정 ────────────────────────────────────────────
st.set_page_config(
    page_title="KRAFTON PPT Translator",
    page_icon="🏢",
    layout="wide"
)

# ── 상수 ──────────────────────────────────────────────────
MODEL         = "claude-sonnet-4-20250514"
LANG_MAP      = {"English": "English", "Japanese": "Japanese", "Chinese": "Traditional Chinese"}
SKIP_NAMES    = ("slide number", "date", "footer", "page")
KOREAN_FONTS  = {
    "Pretendard", "나눔고딕", "맑은 고딕", "Malgun Gothic",
    "NanumGothic", "KoPubWorldBatang", "굴림", "돋움", "바탕", "나눔명조"
}
EN_FONT        = "Calibri"
ADMIN_PASSWORD = "krafton2026"

# ── 기본 Glossary ──────────────────────────────────────────
BASE_GLOSSARY = {
    # 인명
    "장병규": "BG Chang",      "김창한": "CH Kim",         "배동근": "DK Bae",
    "장태석": "TS Jang",       "오진호": "Jin Oh",          "이강욱": "Kangwook Lee",
    "박혜리": "Maria Park",    "윤상훈": "Richard Yoon",   "손현일": "Sean Sohn",
    "한소영": "Soyoung Han",   "이병욱": "Andy Lee",        "김정연": "Jung Yun Kim",
    "김낙형": "Harns Kim",     "박찬민": "Chanmin Park",    "홍정택": "Jeongtack Hong",
    "박재민": "Jaemin Park",   "강재형": "Jaehyung Kang",  "박상훈": "Albert Park",
    "정용": "Young Chung",     "윤수진": "Sujin Yun",       "최승환": "Seunghwan Choi",
    "이동훈": "Donghun Lee",   "조혁일": "Hyuk Cho",        "유재영": "Jaeyoung Yoo",
    "황동경": "Dongkyung Hwang", "김현아": "Hyuna Kim",     "이석현": "Matthew Lee",
    "이창준": "Changjun Lee",  "이지은": "Jieun Lee",       "현종민": "Jongmin Hyun",
    "김정화": "Junghwa Kim",   "신소희": "Sohee Shin",      "노경원": "Kyoungwon Noh",
    "정지현": "Jeehyun Jung",  "김지영": "Jiyoung Kim",     "박서훈": "Seohoon Park",
    "김인영": "Inyoung Kim",   "김고운": "Kowoon Kim",      "어창선": "Chang Seon Eo",
    # 재무
    "매출": "Sales",
    "영업이익": "Operating Profit (OP)",
    "영업이익률": "OPM",
    "회계 매출": "Reported Revenue",
    "재무계획": "Financial Plan",
    "예산": "Budget",
    "연초 실적 점검": "Review of Early-year Performance",
    "전년 대비": "YoY",
    "누적 실적": "Cumulative Performance",
    "계획 대비": "vs. Plan",
    "인건비": "Labor Costs",
    "마케팅비": "Marketing Expenses",
    "지급수수료": "Service Fees",
    "임차료/상각비": "Rental/Depreciation",
    "비통제 비용": "Uncontrollable Costs",
    "매출연동비": "App Fees / Cost of Sales",
    "예비비": "Reserve Fund",
    "손익": "P/L",
    "범 크래프톤": "KRAFTON Family",
    "크래프톤 Business": "KRAFTON Business",
    "비게임 프로젝트": "Non-gaming Projects",
    "전사 비용": "Company-wide Expenses",
    "예산 집중 관리": "Intensive Budget Control",
    # 전략
    "핵심 서비스": "Core Service",
    "신작": "New IP",
    "신규 IP": "New IP",
    "장기 PLC화": "long-term PLC",
    "직접 서비스": "Direct Service",
    "두 자릿수 성장": "Double-digit Growth",
    "계단식 성장": "Stepwise Growth",
    "파이프라인": "Pipeline",
    "보고": "Report",
    "승인/보고": "Approval/Report",
    "승인사항": "Approval Item",
    "보류": "Hold (Deferred)",
    "서면 보고": "Written Report",
    "사전공유": "Pre-sharing",
    "서비스 종료": "Sunset / Discontinue",
    "퍼블리싱 운영 최적화 방안": "Publishing Operations Optimization Plan",
    "운영 Agility": "Operational Agility",
    "체질 개선": "Operational Transformation",
    "런치패드 프로그램": "Launchpad Program",
    "Tentpole 캠페인": "Tentpole Campaign",
    # 조직/HR
    "이사회": "Board of Directors (BOD)",
    "대표이사": "CEO",
    "이사회 의장": "Board Chair",
    "사외이사": "Outside Director",
    "제작 리더십": "Production Leadership",
    "제작 리더": "Production Lead",
    "제작총괄": "Head of Production",
    "비제작 조직": "Non-development Org",
    "신사업 자회사": "New Business Subsidiaries",
    "전문계약직": "Professional Contractors",
    "사내채용": "Internal Hiring",
    "인력계획": "Workforce Plan",
    "증원": "Headcount Increase",
    "유휴인력": "Idle Employees",
    "低성과자": "Low-performing Employees",
    "자발적 퇴직": "Voluntary Exit",
    "희망퇴직": "Voluntary Retirement",
    "정규직": "Full-time Employee",
    "미등기임원": "Non-executive Officers",
    "책임파트너제도": "Accountability Partner System",
    "별첨": "Separately Attached",
}

# ── Glossary DB 파일 핸들러 ────────────────────────────────
import os
GLOSSARY_DB_FILE = "glossary_db.json"

def load_glossary_db():
    if not os.path.exists(GLOSSARY_DB_FILE):
        return {"approved_glossary": {}, "pending_glossary": []}
    try:
        with open(GLOSSARY_DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"approved_glossary": {}, "pending_glossary": []}

def save_glossary_db(db):
    try:
        with open(GLOSSARY_DB_FILE, "w", encoding="utf-8") as f:
            json.dump(db, f, indent=2, ensure_ascii=False)
    except Exception as e:
        st.error(f"DB 저장 오류: {e}")

# ── Session state 초기화 (로컬변수 연동) ──────────────
if "session_extra_glossary" not in st.session_state:
    st.session_state.session_extra_glossary = {}
if "admin_logged_in"        not in st.session_state:
    st.session_state.admin_logged_in = False


def get_active_glossary():
    """BASE + 승인된 항목(DB) + 이번 세션 추가 = 최종 glossary"""
    db = load_glossary_db()
    merged = dict(BASE_GLOSSARY)
    merged.update(db.get("approved_glossary", {}))
    merged.update(st.session_state.session_extra_glossary)
    return merged


# ══════════════════════════════════════════════════════════
# PPT 처리 함수
# ══════════════════════════════════════════════════════════

def should_skip(shape):
    return any(k in getattr(shape, "name", "").lower() for k in SKIP_NAMES)


def has_korean(text):
    return bool(re.search(r"[\uAC00-\uD7A3]", text))


def iter_paragraphs(shapes):
    """shape_id + para_idx 기반 순회 — 그룹/표/중첩 그룹 모두 지원"""
    for shape in shapes:
        if getattr(shape, "has_text_frame", False) and shape.text_frame:
            for pi, para in enumerate(shape.text_frame.paragraphs):
                yield shape, para, pi
        if getattr(shape, "has_table", False) and shape.table:
            pi_counter = 0
            for row in shape.table.rows:
                for cell in row.cells:
                    if getattr(cell, "text_frame", None):
                        for para in cell.text_frame.paragraphs:
                            yield shape, para, pi_counter
                            pi_counter += 1
        if getattr(shape, "shape_type", None) == 6:
            yield from iter_paragraphs(shape.shapes)


def get_shape_width_pt(shape):
    try:
        w_emu = shape.width
        if w_emu:
            tf       = getattr(shape, "text_frame", None)
            left_in  = tf.margin_left  if (tf and tf.margin_left)  else 91440
            right_in = tf.margin_right if (tf and tf.margin_right) else 91440
            return max((w_emu - left_in - right_in) / 12700, 10)
    except Exception:
        pass
    return None


def get_slide_texts(slide):
    result = []
    for global_idx, (shape, para, pi) in enumerate(iter_paragraphs(slide.shapes)):
        if should_skip(shape):
            continue
        full = para.text
        if not full.strip() or not has_korean(full):
            continue
        font_pt = None
        for run in para.runs:
            if run.font.size:
                font_pt = run.font.size.pt
                break
        result.append({
            "shape_id":     shape.shape_id,
            "para_idx":     pi,
            "text":         full,
            "font_pt":      font_pt,
            "box_width_pt": get_shape_width_pt(shape),
        })
    return result


def detect_slide_type(texts):
    combined = " ".join(t["text"] for t in texts)
    if any(k in combined for k in ["승인사항", "승인 요청", "의결"]):
        return "approval"
    if any(k in combined for k in ["억원", "매출", "손익", "YoY"]):
        return "financial"
    if any(k in combined for k in ["출시", "파이프라인", "Q1", "Q2"]):
        return "timeline"
    if "사전공유" in combined:
        return "pre-disclosure"
    return "strategy"


def translate_slide(client, texts, slide_type, target_lang_str, glossary):
    if not texts:
        return {}

    gstr = "\n".join(f"  {k} → {v}" for k, v in glossary.items())
    type_rules = {
        "financial":      "Keep all numbers/units exact. Translate labels/headers only.",
        "approval":       'Use formal request language: "We request approval for..."',
        "strategy":       "Use concise noun phrases for bullets. Lead with the conclusion.",
        "timeline":       "Translate: 출시예정→Scheduled, 지연→Delayed, 완료→Released",
        "pre-disclosure": 'Add "(Pre-Disclosure)" prefix if not present.',
    }
    input_map = {
        str(i): {"text": t["text"]}
        for i, t in enumerate(texts)
    }
    prompt = f"""Translate the following Korean PowerPoint texts into professional {target_lang_str} for a KRAFTON Board of Directors meeting.

Slide type: {slide_type}
{type_rules.get(slide_type, "")}

## Mandatory glossary (use EXACTLY as shown):
{gstr}

## Do NOT translate or modify:
- Numbers, %, financial figures
- Game titles: PUBG, BGMI, OVERDARE, Black Budget, Valor, inZOI
- Company names: KRAFTON, Unknown Worlds, Neon Giant
- Labels: DRI, SL, □, ■, N/A, As-Is, To-Be

## Translation guidelines:
1. QUALITY FIRST — preserve the original meaning, nuance, and tone completely. This is a formal board meeting document.
2. Use natural, executive-level English that a native speaker would write for a board deck.
3. Do NOT over-shorten. If the original has detailed reasoning, keep it. Cutting meaning is worse than slightly longer text.
4. Preserve \\n line breaks exactly as in the original.
5. Preserve the exact position of (사전공유)→(Pre-sharing) within the sentence. Do not move it.
6. For financial slides: keep all numbers and units exact, translate labels with precision.
7. For approval slides: use formal language e.g. "We hereby request approval for..."

## Input (JSON — index → {{"text"}}):
{json.dumps(input_map, ensure_ascii=False, indent=2)}

## Output format:
Return ONLY a valid JSON object. Same index keys, string values only. No markdown, no explanation.
Example: {{"0": "Translated text", "1": "Another translation"}}"""

    system_prompt = "You are a professional Korean-English translator specializing in corporate board meeting materials for KRAFTON. Your top priority is accurate, natural, high-quality English that preserves the original nuance and intent. Return ONLY valid JSON with string values."

    res = client.messages.create(
        model=MODEL,
        max_tokens=4096,
        system=system_prompt,
        messages=[{"role": "user", "content": prompt}]
    )
    raw   = res.content[0].text.strip()
    match = re.search(r"\{[\s\S]*\}", raw)
    if not match:
        raise ValueError("JSON 파싱 실패")
    matched_str = match.group()
    try:
        return json.loads(matched_str)
    except Exception:
        import ast
        try:
            safe_str = matched_str.replace("null", "None").replace("true", "True").replace("false", "False")
            return ast.literal_eval(safe_str)
        except Exception:
            raise ValueError("JSON 및 보조 파싱 모두 실패")


def replace_para_text(para, new_text, shape=None, min_pt=7):
    # ── 방어: string이 아니면 무조건 skip ─────────────────
    if not new_text or not isinstance(new_text, str):
        return
    new_text = new_text.strip()
    if not new_text:
        return

    # 첫 run에서 서식 추출
    orig_font_size = None
    orig_font_name = None
    for run in para.runs:
        if run.font.size:
            orig_font_size = run.font.size.pt
        if run.font.name:
            orig_font_name = run.font.name
        break

    # run 교체
    for i, run in enumerate(para.runs):
        if i == 0:
            run.text = new_text
            
            # 명시된 폰트 사이즈가 있을 때만 크기 재계산 (상속받은 주석 폰트 등이 12pt로 커지는 현상 방지)
            if orig_font_size:
                final_pt = orig_font_size
                min_allowed = max(orig_font_size - 4, min_pt)
                
                if shape is not None:
                    try:
                        w_emu = shape.width
                        if w_emu:
                            tf = getattr(shape, "text_frame", None)
                            left_in  = tf.margin_left  if (tf and getattr(tf, "margin_left", None))  else 91440
                            right_in = tf.margin_right if (tf and getattr(tf, "margin_right", None)) else 91440
                            box_w = max((w_emu - left_in - right_in) / 12700, 10)
                            lines = new_text.split('\n')
                            max_line_len = max(len(l) for l in lines) if lines else len(new_text)
                            if max_line_len > 0:
                                required_font = box_w / (max_line_len * 0.55)
                                if required_font < final_pt:
                                    final_pt = max(round(required_font, 1), min_allowed)
                    except Exception:
                        pass
                
                run.font.size = Pt(final_pt)

            if orig_font_name and orig_font_name in KOREAN_FONTS:
                run.font.name = EN_FONT
        else:
            run.text = ""

    # run이 없는 경우 새로 생성
    if not para.runs:
        from pptx.oxml.ns import qn
        from lxml import etree
        r   = etree.SubElement(para._p, qn("a:r"))
        etree.SubElement(r, qn("a:rPr"), attrib={"lang": "en-US"})
        t   = etree.SubElement(r, qn("a:t"))
        t.text = new_text


# ══════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════
tab_translate, tab_glossary = st.tabs([
    "🚀 번역", "📖 Glossary"
])


# ──────────────────────────────────────────────────────────
# TAB 1: 번역
# ──────────────────────────────────────────────────────────
with tab_translate:
    st.header("🏢 KRAFTON BOD PPT Translator")
    st.caption("BOD 자료 PPT를 올려주시면 AI가 번역해드립니다 🙌 인명·용어 glossary가 자동 적용되고, 레이아웃도 최대한 원본 그대로 유지해요. 다만 번역 후 텍스트 길이 차이로 포맷이 살짝 틀어질 수 있으니, 다운로드 후 간단히 확인해주세요!")

    col1, col2 = st.columns(2)
    with col1:
        target_lang = st.selectbox("번역 언어", ["English", "Japanese", "Chinese"])
    with col2:
        user_min_pt = st.number_input("최소 허용 폰트 크기 (pt)", min_value=1, max_value=40, value=7)


    # API Key — Streamlit Secrets에서 자동 로드
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        st.error("⚠️ API Key 미설정. Streamlit Cloud → Settings → Secrets에 ANTHROPIC_API_KEY를 추가해주세요.")

    active_glossary = get_active_glossary()
    st.caption(
        f"현재 적용 Glossary: **{len(active_glossary)}개** 항목 "
        f"(기본 {len(BASE_GLOSSARY)} "
        f"+ 승인 {len(load_glossary_db()['approved_glossary'])} "
        f"+ 세션 {len(st.session_state.session_extra_glossary)})"
    )

    uploaded_file = st.file_uploader("PPT 파일 업로드 (.pptx)", type=["pptx"])

    if uploaded_file and api_key:
        st.success(f"✅ **{uploaded_file.name}** 업로드 완료")

        if st.button("🚀 번역 시작", type="primary", use_container_width=True):
            client          = anthropic.Anthropic(api_key=api_key)
            target_lang_str = LANG_MAP.get(target_lang, "English")
            active_glossary = get_active_glossary()

            # 파싱
            pptx_bytes = uploaded_file.read()
            prs        = Presentation(io.BytesIO(pptx_bytes))

            with st.spinner("📊 텍스트 파싱 중..."):
                all_slides_info = []
                total_ko        = 0
                for slide in prs.slides:
                    texts = get_slide_texts(slide)
                    all_slides_info.append(texts)
                    total_ko += len(texts)

            st.info(f"총 {len(prs.slides)}장 · 번역 대상 **{total_ko}개** 단락")

            # 번역
            progress_bar     = st.progress(0)
            status_text      = st.empty()
            log_area         = st.empty()
            all_translations = []
            log_lines        = []

            for si, texts in enumerate(all_slides_info):
                ko_count = len(texts)
                progress_bar.progress((si + 1) / len(all_slides_info))

                if ko_count == 0:
                    all_translations.append({})
                    log_lines.append(f"⏭️  Slide {si+1:2d} — 한국어 없음")
                else:
                    slide_type = detect_slide_type(texts)
                    status_text.text(f"Slide {si+1}/{len(all_slides_info)} [{slide_type}] 번역 중...")
                    for attempt in range(2):
                        try:
                            translated = translate_slide(
                                client, texts, slide_type, target_lang_str, active_glossary
                            )
                            all_translations.append(translated)
                            log_lines.append(f"✅ Slide {si+1:2d} [{slide_type:15s}] — {ko_count}개")
                            break
                        except Exception as e:
                            if attempt == 0:
                                log_lines.append(f"⚠️  Slide {si+1:2d} 재시도... ({e})")
                                time.sleep(2)
                            else:
                                all_translations.append({})
                                log_lines.append(f"❌  Slide {si+1:2d} 실패 — 원본 유지")
                    time.sleep(0.3)

                log_area.code("\n".join(log_lines))

            # PPT 텍스트 교체 (shape_id + para_idx 기반)
            status_text.text("💾 PPT 생성 중...")
            for si, (slide, texts, translated_map) in enumerate(
                zip(prs.slides, all_slides_info, all_translations)
            ):
                if not translated_map:
                    continue

                # shape_id → shape 매핑 테이블 생성
                shape_map = {}
                def _collect_shapes(shapes):
                    for shape in shapes:
                        shape_map[shape.shape_id] = shape
                        if getattr(shape, "shape_type", None) == 6:
                            _collect_shapes(shape.shapes)
                _collect_shapes(slide.shapes)

                for ti, text_info in enumerate(texts):
                    tr = translated_map.get(str(ti))
                    if isinstance(tr, dict):
                        tr = tr.get("text") or tr.get("translation") or ""
                    
                    if not tr or not isinstance(tr, str):
                        continue
                    tr = tr.strip()
                    if not tr:
                        continue

                    shape_id = text_info["shape_id"]
                    para_idx = text_info["para_idx"]
                    shape    = shape_map.get(shape_id)
                    if shape is None:
                        continue

                    # 해당 para 찾기 (텍스트프레임 / 표 셀)
                    para = None
                    if getattr(shape, "has_text_frame", False) and shape.text_frame:
                        paras = shape.text_frame.paragraphs
                        if para_idx < len(paras):
                            para = paras[para_idx]
                    elif getattr(shape, "has_table", False) and shape.table:
                        all_paras = []
                        for row in shape.table.rows:
                            for cell in row.cells:
                                if getattr(cell, "text_frame", None):
                                    all_paras.extend(cell.text_frame.paragraphs)
                        if para_idx < len(all_paras):
                            para = all_paras[para_idx]

                    if para is not None:
                        replace_para_text(para, tr, shape=shape, min_pt=user_min_pt)

            # 저장 & 다운로드
            output = io.BytesIO()
            prs.save(output)
            output.seek(0)
            out_name = uploaded_file.name.rsplit(".", 1)[0] + f"_{target_lang[:2].upper()}.pptx"

            progress_bar.progress(1.0)
            status_text.empty()
            st.success("🎉 번역 완료!")
            st.download_button(
                label=f"⬇️ {out_name} 다운로드",
                data=output,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                type="primary",
            )


# ──────────────────────────────────────────────────────────
# TAB 2: Glossary
# ──────────────────────────────────────────────────────────
with tab_glossary:
    st.header("📖 Glossary")
    st.info("💬 Glossary 추가/수정 요청은 **도지영** (michelle@krafton.com)에게 연락해주세요. [Slack](https://krafton.enterprise.slack.com/team/U02RWGEGJ5B)")

    active   = get_active_glossary()

    search = st.text_input("🔍 검색", placeholder="한국어 또는 영문으로 검색...")
    
    db_data = load_glossary_db()
    approved = db_data.get('approved_glossary', {})
    st.caption(f"총 **{len(active)}개** 항목 (기본 {len(BASE_GLOSSARY)} + 추가 {len(approved)}개)")

    NAME_KEYS = set(list(BASE_GLOSSARY.keys())[:39])

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("#### 👤 인명")
        for ko, en in BASE_GLOSSARY.items():
            if ko not in NAME_KEYS:
                continue
            if search and search.lower() not in ko.lower() and search.lower() not in en.lower():
                continue
            st.markdown(f"`{ko}` → {en}")

    with col_b:
        st.markdown("#### 📝 용어")
        for ko, en in BASE_GLOSSARY.items():
            if ko in NAME_KEYS:
                continue
            if search and search.lower() not in ko.lower() and search.lower() not in en.lower():
                continue
            st.markdown(f"`{ko}` → {en}")

    if approved:
        st.divider()
        st.markdown(f"#### ✅ 추가 등록된 항목 — {len(approved)}개")
        for ko, en in approved.items():
            if search and search.lower() not in ko.lower() and search.lower() not in en.lower():
                continue
            st.markdown(f"`{ko}` → {en}")



# ── 플로팅 피드백 버튼 ─────────────────────────────────────
st.markdown("""
<style>
#feedback-btn {
    position: fixed;
    bottom: 28px;
    right: 28px;
    width: 52px;
    height: 52px;
    border-radius: 50%;
    background: #E8273A;
    color: white;
    font-size: 22px;
    border: none;
    cursor: pointer;
    box-shadow: 0 4px 12px rgba(0,0,0,0.25);
    z-index: 9999;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: transform 0.2s;
}
#feedback-btn:hover { transform: scale(1.1); }

#feedback-modal {
    display: none;
    position: fixed;
    bottom: 90px;
    right: 28px;
    width: 370px;
    height: 500px;
    background: white;
    border-radius: 16px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.18);
    z-index: 9998;
    overflow: hidden;
    flex-direction: column;
}
#feedback-modal.open { display: flex; }

#modal-header {
    background: #E8273A;
    color: white;
    padding: 14px 18px;
    font-weight: 600;
    font-size: 14px;
    display: flex;
    justify-content: space-between;
    align-items: center;
}
#modal-close {
    background: none;
    border: none;
    color: white;
    font-size: 18px;
    cursor: pointer;
    padding: 0;
    line-height: 1;
}
#feedback-modal iframe {
    flex: 1;
    border: none;
    width: 100%;
}
</style>

<button class="st-emotion-cache-1r6slb0" id="feedback-btn" onclick="toggleModal()" title="Glossary 제안 / 의견 남기기">
  💬
</button>

<div id="feedback-modal">
  <div id="modal-header">
    <span>💬 Glossary 제안 / 의견 남기기</span>
    <button id="modal-close" onclick="toggleModal()">✕</button>
  </div>
  <iframe src="https://docs.google.com/forms/d/e/1FAIpQLSezU-H6m0TMt2Ve-QUTZv483JklIdtfAsKi7rvYNW74l5B5lw/viewform?embedded=true"
          loading="lazy">
  </iframe>
</div>

<script>
function toggleModal() {
    const modal = document.getElementById('feedback-modal');
    modal.classList.toggle('open');
}
// 모달 외부 클릭 시 닫기
document.addEventListener('click', function(e) {
    const modal = document.getElementById('feedback-modal');
    const btn = document.getElementById('feedback-btn');
    if (modal && btn && modal.classList.contains('open') && 
        !modal.contains(e.target) && 
        !btn.contains(e.target)) {
        modal.classList.remove('open');
    }
});
</script>
""", unsafe_allow_html=True)
