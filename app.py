import streamlit as st
import re, json, time, io, base64, os
from pptx import Presentation
from pptx.util import Pt
import anthropic
from datetime import datetime

# QC imports
try:
    import fitz  # PyMuPDF
    from PIL import Image
    QC_AVAILABLE = True
except ImportError:
    QC_AVAILABLE = False

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

# QC session state
for _k, _v in {
    "qc_pages_ko": [], "qc_pages_en": [],
    "qc_num_pages": 0, "qc_aspect": 0.5625,
    "qc_current": 0, "qc_mode": "compare",
    "qc_processed": False,
    "qc_status": {}, "qc_notes": {}, "qc_reviews": {},
}.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# Translation Review session state
for _k, _v in {
    "tr_slides": [],          # [{texts: [{text, ...}], translated: {idx: str}}]
    "tr_reviews": {},         # {slide_idx: {verdict, summary, issues}}
    "tr_has_data": False,
    "tr_file_name": "",
}.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


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
c_top1, c_top2 = st.columns([3, 1])
with c_top2:
    st.link_button("💬 Glossary 제안 / 의견 남기기", "https://docs.google.com/forms/d/e/1FAIpQLSezU-H6m0TMt2Ve-QUTZv483JklIdtfAsKi7rvYNW74l5B5lw/viewform", use_container_width=True)

tab_translate, tab_qc, tab_glossary = st.tabs([
    "🚀 번역", "🔍 번역 QC", "📖 Glossary"
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
            start_time = time.time()
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
            elapsed_time = time.time() - start_time
            st.success(f"🎉 번역 완료! (⏱️ 소요시간: {elapsed_time:.1f}초)")
            st.download_button(
                label=f"⬇️ {out_name} 다운로드",
                data=output,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                type="primary",
            )

            # ── 번역 결과를 session state에 저장 (AI 검토용) ──
            tr_data = []
            for si, (texts, translated_map) in enumerate(zip(all_slides_info, all_translations)):
                pairs = []
                for ti, text_info in enumerate(texts):
                    orig = text_info["text"]
                    tr = translated_map.get(str(ti), "")
                    if isinstance(tr, dict):
                        tr = tr.get("text") or tr.get("translation") or ""
                    if not isinstance(tr, str):
                        tr = ""
                    pairs.append({"ko": orig, "en": tr.strip()})
                tr_data.append(pairs)
            st.session_state.tr_slides = tr_data
            st.session_state.tr_has_data = True
            st.session_state.tr_file_name = uploaded_file.name
            st.session_state.tr_reviews = {}

    # ══════════════════════════════════════════════════
    # 번역 결과 AI 검토 (번역 탭 하단)
    # ══════════════════════════════════════════════════
    if st.session_state.tr_has_data:
        st.divider()
        st.subheader("📋 번역 결과 검토")
        st.caption("번역된 텍스트를 슬라이드별로 확인하고, AI가 원문과 비교하여 오류를 검토합니다.")

        tr_data = st.session_state.tr_slides
        tr_reviews = st.session_state.tr_reviews
        has_reviews = len(tr_reviews) > 0

        # AI Review button
        tr_c1, tr_c2, tr_c3 = st.columns([1, 1, 2])
        with tr_c1:
            if api_key and st.button("🤖 전체 AI 검토", type="primary",
                                      use_container_width=True, key="tr_ai_all"):
                glossary = get_active_glossary()
                gstr = "\n".join(f"  {k} → {v}" for k, v in glossary.items())
                client = anthropic.Anthropic(api_key=api_key)
                reviews = {}
                prog = st.progress(0, text="AI 검토 중...")
                for si, pairs in enumerate(tr_data):
                    if not pairs:
                        reviews[si] = {"verdict": "ok", "summary": "텍스트 없음", "issues": []}
                    else:
                        ko_block = "\n".join(f"  [{i+1}] {p['ko']}" for i, p in enumerate(pairs))
                        en_block = "\n".join(f"  [{i+1}] {p['en']}" for i, p in enumerate(pairs))
                        try:
                            resp = client.messages.create(
                                model=MODEL, max_tokens=2048,
                                system=f"""You are a translation QC reviewer for KRAFTON board meeting materials.
Compare Korean originals with English translations TEXT BY TEXT.

## 핵심: 의역은 정상
- 문장 구조 달라도 의미 같으면 OK
- 간결하게 줄인 번역도 핵심 보존되면 OK
- 자연스러운 영어 표현으로 바꾼 건 좋은 번역

## "fix" — 진짜 오류만:
- 숫자/금액 틀림
- 핵심 내용 완전 누락
- 의미가 반대로 번역됨
- 인명/고유명사 오류

## "warn" — 확인 필요:
- 같은 문서 내 동일 용어 번역 불일치
- 번역이 안 된 한국어가 남아있음

## 지적하지 않을 것:
- 의역, 어순 변경, 문체 차이
- Glossary와 일치하는 번역
- 문맥상 Glossary와 다른 번역이 자연스러운 경우

## Glossary (참고용):
{gstr}

대부분의 슬라이드는 "ok"여야 합니다.

Return JSON:
{{"verdict":"ok"|"warn"|"fix","summary":"한줄 한국어 요약","issues":[{{"level":"error"|"warn"|"info","detail":"한국어 설명"}}]}}
ONLY valid JSON.""",
                                messages=[{"role":"user","content":f"슬라이드 {si+1} 검토:\n\n[한국어 원문]\n{ko_block}\n\n[영문 번역]\n{en_block}"}]
                            )
                            raw = resp.content[0].text.strip().replace("```json","").replace("```","").strip()
                            reviews[si] = json.loads(raw)
                        except Exception as e:
                            reviews[si] = {"verdict":"warn","summary":f"검토 실패: {e}","issues":[]}
                    prog.progress((si+1)/len(tr_data), text=f"검토 중... {si+1}/{len(tr_data)}")
                prog.empty()
                st.session_state.tr_reviews = reviews
                st.rerun()

        with tr_c2:
            if has_reviews:
                # TXT export
                lines = ["═"*50, "  번역 QC Report", f"  {datetime.now().strftime('%Y-%m-%d %H:%M')}", "═"*50, ""]
                for si, pairs in enumerate(tr_data):
                    rv = tr_reviews.get(si, {})
                    vd = rv.get("verdict","")
                    sm = rv.get("summary","")
                    iss = rv.get("issues",[])
                    icon = {"ok":"✅","warn":"⚠️","fix":"❌"}.get(vd,"⬜")
                    lines.append("─"*50)
                    lines.append(f"  슬라이드 {si+1}  {icon} {sm}")
                    lines.append("─"*50)
                    for p in pairs:
                        lines.append(f"  KR: {p['ko']}")
                        lines.append(f"  EN: {p['en']}")
                        lines.append("")
                    if iss:
                        lines.append("  이슈:")
                        for x in iss:
                            ic = {"error":"🔴","warn":"🟡","info":"🔵"}.get(x.get("level",""),"ℹ️")
                            lines.append(f"    {ic} {x.get('detail','')}")
                    lines.append("")
                lines += ["═"*50, "  END OF REPORT", "═"*50]
                st.download_button("📝 검토 리포트 (TXT)", "\n".join(lines),
                                   file_name=f"Translation_QC_{datetime.now().strftime('%Y%m%d')}.txt",
                                   mime="text/plain", use_container_width=True, key="tr_dl_txt")

        # Summary
        if has_reviews:
            counts = {}
            for rv in tr_reviews.values():
                v = rv.get("verdict","ok")
                counts[v] = counts.get(v, 0) + 1
            mc1,mc2,mc3,mc4 = st.columns(4)
            with mc1: st.metric("전체", f"{len(tr_data)}장")
            with mc2: st.metric("✅ OK", f"{counts.get('ok',0)}장")
            with mc3: st.metric("⚠️ 확인", f"{counts.get('warn',0)}장")
            with mc4: st.metric("❌ 수정", f"{counts.get('fix',0)}장")

        # Per-slide results
        for si, pairs in enumerate(tr_data):
            if not pairs:
                continue
            rv = tr_reviews.get(si, {}) if has_reviews else {}
            vd = rv.get("verdict", "")
            sm = rv.get("summary", "")
            iss = rv.get("issues", [])

            # Header
            if has_reviews:
                colors = {"ok":("#F0FDF4","#166534","#BBF7D0"),"warn":("#FFFBEB","#92400E","#FDE68A"),"fix":("#FEF2F2","#991B1B","#FECACA")}
                bg,fg,bd = colors.get(vd, ("#F3F4F6","#6B7280","#E5E7EB"))
                icon = {"ok":"✅","warn":"⚠️","fix":"❌"}.get(vd,"⬜")
                st.markdown(
                    f'<div style="display:flex;align-items:center;gap:10px;margin-top:16px;margin-bottom:8px;">'
                    f'<span style="font-size:15px;font-weight:700;color:#111827;">슬라이드 {si+1}</span>'
                    f'<span style="font-size:11px;font-weight:600;padding:3px 12px;border-radius:16px;'
                    f'background:{bg};color:{fg};border:1px solid {bd};">{icon} {sm}</span></div>',
                    unsafe_allow_html=True)
            else:
                st.markdown(f"**슬라이드 {si+1}**")

            # Text comparison table
            table_data = [{"원문 (한국어)": p["ko"], "번역 (English)": p["en"]} for p in pairs]
            st.dataframe(table_data, use_container_width=True, hide_index=True)

            # Issues
            if iss:
                for x in iss:
                    lv = x.get("level","info")
                    dt = x.get("detail","")
                    styles = {"error":"background:#FEF2F2;border-left:3px solid #EF4444;color:#991B1B",
                              "warn":"background:#FFFBEB;border-left:3px solid #F59E0B;color:#92400E",
                              "info":"background:#EFF6FF;border-left:3px solid #3B82F6;color:#1E40AF"}
                    ic = {"error":"🔴","warn":"🟡","info":"🔵"}.get(lv,"ℹ️")
                    st.markdown(f'<div style="padding:6px 12px;margin:3px 0;border-radius:6px;font-size:12px;line-height:1.5;{styles.get(lv,"")}">{ic} {dt}</div>', unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────
# TAB 1 하단: Delta 번역
# ──────────────────────────────────────────────────────────
with tab_translate:
    st.divider()
    st.subheader("🔄 Delta 번역 — 수정분만 재번역")
    st.caption(
        "이전 한글 원문(v1)과 수정된 한글 원문(v2)을 비교해 **달라진 슬라이드만 감지**하여 재번역합니다. "
        "슬라이드가 추가·삭제되어 페이지 번호가 밀려도 제목과 내용을 기반으로 같은 슬라이드를 찾아 매칭합니다. "
        "수정된 슬라이드에는 **UPDATED 배지**, 신규 슬라이드에는 **NEW 배지**가 표시되고, "
        "맨 앞에 변경 전/후 텍스트 비교가 담긴 **변경사항 요약 시트**가 자동으로 추가됩니다."
    )
    st.warning(
        "🚧 **Under Construction** — 현재 테스트 중인 기능입니다. "
        "매칭 정확도는 슬라이드 내 텍스트 양에 따라 달라질 수 있으며, "
        "결과물은 반드시 사람이 검토 후 사용해 주세요. "
        "오작동이나 개선 의견은 **도지영**에게 알려주세요!",
        icon="⚠️"
    )

    dc1, dc2 = st.columns(2)
    with dc1:
        delta_lang    = st.selectbox("번역 언어", ["English", "Japanese", "Chinese"], key="delta_lang")
        delta_min_pt  = st.number_input("최소 허용 폰트 크기 (pt)", min_value=1, max_value=40, value=7, key="delta_min_pt")
    with dc2:
        st.markdown("##### 파일 업로드")
        old_translated_file = st.file_uploader("① 이전 한글 원문 v1 (.pptx)", type=["pptx"], key="old_tr")
        new_original_file   = st.file_uploader("② 수정된 한글 원문 v2 (.pptx)", type=["pptx"], key="new_orig")

    delta_api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not delta_api_key:
        st.error("⚠️ API Key 미설정.")

    if old_translated_file and new_original_file and delta_api_key:
        st.success(f"✅ **{old_translated_file.name}** (v1) + **{new_original_file.name}** (v2) 업로드 완료")

        if st.button("🔍 Delta 감지 후 번역", type="primary", use_container_width=True):
            import time as _time

            start_time = _time.time()
            delta_lang_str  = LANG_MAP.get(delta_lang, "English")
            active_glossary = get_active_glossary()
            client          = anthropic.Anthropic(api_key=delta_api_key)

            prs_old = Presentation(io.BytesIO(old_translated_file.read()))
            prs_new = Presentation(io.BytesIO(new_original_file.read()))

            # ── Delta 감지 (콘텐츠 지문 기반 매칭) ───────────
            with st.spinner("🔍 변경된 슬라이드 감지 중..."):

                def _get_slide_tokens(slide):
                    """슬라이드 텍스트 줄 리스트 — text_frame + table 모두 포함"""
                    lines = []
                    for shape in slide.shapes:
                        # 일반 텍스트 박스
                        if getattr(shape, "has_text_frame", False) and shape.text_frame:
                            for para in shape.text_frame.paragraphs:
                                t = para.text.strip()
                                if t:
                                    lines.append(t)
                        # 표(Table) 안 텍스트
                        if getattr(shape, "has_table", False) and shape.table:
                            for row in shape.table.rows:
                                for cell in row.cells:
                                    if getattr(cell, "text_frame", None):
                                        for para in cell.text_frame.paragraphs:
                                            t = para.text.strip()
                                            if t:
                                                lines.append(t)
                    return lines

                def _get_title(lines):
                    """첫 번째 비어있지 않은 줄을 제목으로 사용"""
                    return lines[0] if lines else ""

                def _similarity(lines_a, lines_b):
                    """제목 가중 Jaccard 유사도
                    - 빈 슬라이드 둘 다: 0.0 반환 (빈 슬라이드끼리 매칭 방지)
                    - 제목이 같으면 +0.3 보너스 (제목 가중치)
                    - 나머지는 단어 단위 Jaccard
                    """
                    if not lines_a or not lines_b:
                        return 0.0   # 빈 슬라이드는 매칭 안 함

                    set_a = set(" ".join(lines_a).split())
                    set_b = set(" ".join(lines_b).split())
                    inter = len(set_a & set_b)
                    union = len(set_a | set_b)
                    jaccard = inter / union if union else 0.0

                    # 제목 일치 보너스
                    title_a = _get_title(lines_a)
                    title_b = _get_title(lines_b)
                    title_bonus = 0.3 if title_a and title_b and title_a == title_b else 0.0

                    return min(jaccard + title_bonus, 1.0)

                def _text_changed(lines_a, lines_b):
                    """실제 텍스트 내용이 바뀌었는지 판단
                    단어 집합이 완전히 동일하면 unchanged, 아니면 changed
                    → 숫자/날짜 한 글자만 바뀌어도 감지"""
                    return " ".join(sorted(lines_a)) != " ".join(sorted(lines_b))

                # v1, v2 슬라이드 텍스트 추출
                v1_slides = [_get_slide_tokens(s) for s in prs_old.slides]
                v2_slides = [_get_slide_tokens(s) for s in prs_new.slides]

                MATCH_THRESHOLD = 0.35   # 이 이상이면 "같은 슬라이드" 후보

                # ── Hungarian-style 최적 매칭 ──────────────────
                # 모든 (v2, v1) 쌍의 유사도를 먼저 계산한 뒤
                # 높은 점수 순으로 탐욕적으로 할당 → Greedy 순서 의존성 해결
                score_pairs = []
                for v2_idx, v2_lines in enumerate(v2_slides):
                    for v1_idx, v1_lines in enumerate(v1_slides):
                        score = _similarity(v1_lines, v2_lines)
                        if score >= MATCH_THRESHOLD:
                            score_pairs.append((score, v2_idx, v1_idx))

                # 유사도 높은 순 정렬 후 1:1 매칭
                score_pairs.sort(reverse=True)
                used_v2 = set()
                used_v1 = set()
                match_map = {v2_idx: None for v2_idx in range(len(v2_slides))}

                for score, v2_idx, v1_idx in score_pairs:
                    if v2_idx in used_v2 or v1_idx in used_v1:
                        continue
                    match_map[v2_idx] = v1_idx
                    used_v2.add(v2_idx)
                    used_v1.add(v1_idx)

                # 변경/신규 슬라이드 분류
                total_new       = len(v2_slides)
                changed_indices = []
                new_indices     = []
                delta_report    = {}

                for v2_idx in range(total_new):
                    v1_idx   = match_map[v2_idx]
                    v2_lines = v2_slides[v2_idx]

                    if v1_idx is None:
                        # 신규 슬라이드
                        new_indices.append(v2_idx)
                        delta_report[v2_idx] = {
                            "before": [], "after": v2_lines,
                            "v1_idx": None, "type": "new"
                        }
                    else:
                        v1_lines = v1_slides[v1_idx]
                        if _text_changed(v1_lines, v2_lines):
                            # 매칭됐는데 내용 다름 → 수정
                            changed_indices.append(v2_idx)
                            delta_report[v2_idx] = {
                                "before": v1_lines, "after": v2_lines,
                                "v1_idx": v1_idx, "type": "modified"
                            }
                        # 완전 동일 → 무시

            all_delta_indices = changed_indices + new_indices
            if not all_delta_indices:
                st.info("✅ 변경된 슬라이드가 없습니다. 번역이 필요하지 않아요!")
            else:
                summary_parts = []
                if changed_indices:
                    summary_parts.append(f"수정 **{len(changed_indices)}개** → Slide {[i+1 for i in changed_indices]}")
                if new_indices:
                    summary_parts.append(f"신규 **{len(new_indices)}개** → Slide {[i+1 for i in new_indices]}")
                st.info(f"🔄 변경 감지: {' · '.join(summary_parts)}")

                # ── 앱 화면: 변경 전/후 텍스트 리포트 ────────
                with st.expander("📋 변경사항 미리보기 (변경 전/후)", expanded=True):
                    for si in all_delta_indices:
                        info   = delta_report[si]
                        v1_ref = f" ← v1 Slide {info['v1_idx']+1}" if info["v1_idx"] is not None else ""
                        badge  = "🆕 신규" if info["type"] == "new" else "🔸 수정"
                        st.markdown(f"**{badge} Slide {si+1}**{v1_ref}")
                        before_lines = info["before"]
                        after_lines  = info["after"]

                        # 내용 기반 diff — 추가/삭제/유지를 텍스트 집합으로 판단
                        before_set = set(before_lines)
                        after_set  = set(after_lines)
                        added      = after_set  - before_set   # v2에만 있는 줄
                        removed    = before_set - after_set    # v1에만 있는 줄
                        kept       = before_set & after_set    # 동일한 줄

                        rows = []
                        # 유지된 줄
                        for t in before_lines:
                            if t in kept:
                                rows.append({"변경": "", "변경 전 (v1)": t, "변경 후 (v2)": t})
                        # 삭제된 줄 (v1에만)
                        for t in before_lines:
                            if t in removed:
                                rows.append({"변경": "🔴 삭제", "변경 전 (v1)": t, "변경 후 (v2)": ""})
                        # 추가된 줄 (v2에만)
                        for t in after_lines:
                            if t in added:
                                rows.append({"변경": "🟢 추가", "변경 전 (v1)": "", "변경 후 (v2)": t})

                        if not rows:
                            rows = [{"변경": "🆕 신규 슬라이드", "변경 전 (v1)": "", "변경 후 (v2)": " / ".join(after_lines[:5])}]

                        st.dataframe(rows, use_container_width=True, hide_index=True)
                        st.write("")

                progress_bar = st.progress(0)
                status_text  = st.empty()
                log_area     = st.empty()
                log_lines    = []

                for step, slide_idx in enumerate(all_delta_indices):
                    slide  = prs_new.slides[slide_idx]
                    texts  = get_slide_texts(slide)
                    progress_bar.progress((step + 1) / len(all_delta_indices))

                    if not texts:
                        log_lines.append(f"⏭️  Slide {slide_idx+1:2d} — 한국어 없음")
                        log_area.code("\n".join(log_lines))
                        continue

                    slide_type = detect_slide_type(texts)
                    status_text.text(f"Slide {slide_idx+1}/{total_new} [{slide_type}] 번역 중...")

                    for attempt in range(2):
                        try:
                            translated_map = translate_slide(
                                client, texts, slide_type, delta_lang_str, active_glossary
                            )
                            break
                        except Exception as e:
                            if attempt == 0:
                                log_lines.append(f"⚠️  Slide {slide_idx+1:2d} 재시도... ({e})")
                                log_area.code("\n".join(log_lines))
                                _time.sleep(2)
                            else:
                                translated_map = {}
                                log_lines.append(f"❌  Slide {slide_idx+1:2d} 실패 — 원본 유지")

                    # 텍스트 교체
                    shape_map = {}
                    def _collect(shapes):
                        for s in shapes:
                            shape_map[s.shape_id] = s
                            if getattr(s, "shape_type", None) == 6:
                                _collect(s.shapes)
                    _collect(slide.shapes)

                    for ti, text_info in enumerate(texts):
                        tr = translated_map.get(str(ti))
                        if isinstance(tr, dict):
                            tr = tr.get("text") or tr.get("translation") or ""
                        if not tr or not isinstance(tr, str):
                            continue
                        tr = tr.strip()
                        if not tr:
                            continue
                        shape = shape_map.get(text_info["shape_id"])
                        if shape is None:
                            continue
                        para = None
                        if getattr(shape, "has_text_frame", False) and shape.text_frame:
                            paras = shape.text_frame.paragraphs
                            if text_info["para_idx"] < len(paras):
                                para = paras[text_info["para_idx"]]
                        elif getattr(shape, "has_table", False) and shape.table:
                            all_paras = []
                            for row in shape.table.rows:
                                for cell in row.cells:
                                    if getattr(cell, "text_frame", None):
                                        all_paras.extend(cell.text_frame.paragraphs)
                            if text_info["para_idx"] < len(all_paras):
                                para = all_paras[text_info["para_idx"]]
                        if para is not None:
                            replace_para_text(para, tr, shape=shape, min_pt=delta_min_pt)

                    # 배지 추가 — 수정/신규 구분
                    from pptx.util import Inches, Pt as _Pt
                    from pptx.dml.color import RGBColor as _RGB
                    from pptx.enum.text import PP_ALIGN as _ALIGN
                    is_new      = delta_report.get(slide_idx, {}).get("type") == "new"
                    badge_text  = "NEW" if is_new else "UPDATED"
                    badge_color = _RGB(0x2E, 0x7D, 0x32) if is_new else _RGB(0xFF, 0xC0, 0x00)
                    border_color= _RGB(0x1B, 0x5E, 0x20) if is_new else _RGB(0xCC, 0x99, 0x00)
                    text_color  = _RGB(0xFF, 0xFF, 0xFF) if is_new else _RGB(0x4A, 0x2D, 0x00)

                    badge_w, badge_h = Inches(1.0), Inches(0.30)
                    badge_l = prs_new.slide_width - badge_w - Inches(0.18)
                    badge_t = Inches(0.12)
                    badge_shape = slide.shapes.add_shape(1, badge_l, badge_t, badge_w, badge_h)
                    badge_shape.fill.solid()
                    badge_shape.fill.fore_color.rgb = badge_color
                    badge_shape.line.color.rgb      = border_color
                    badge_shape.line.width          = _Pt(0.75)
                    tf = badge_shape.text_frame
                    tf.word_wrap = False
                    p  = tf.paragraphs[0]
                    p.alignment = _ALIGN.CENTER
                    run = p.add_run()
                    run.text           = badge_text
                    run.font.bold      = True
                    run.font.size      = _Pt(8.5)
                    run.font.color.rgb = text_color

                    log_lines.append(f"✅ Slide {slide_idx+1:2d} [{slide_type:15s}] — {len(texts)}개 번역")
                    log_area.code("\n".join(log_lines))
                    _time.sleep(0.3)

                # ── 변경사항 요약 시트 (맨 앞 삽입) ──────────
                status_text.text("📋 변경사항 요약 시트 생성 중...")
                blank_layout  = prs_new.slide_layouts[6]
                summary_slide = prs_new.slides.add_slide(blank_layout)
                xml_slides    = prs_new.slides._sldIdLst
                new_entry     = xml_slides[-1]
                xml_slides.remove(new_entry)
                xml_slides.insert(0, new_entry)

                sw = prs_new.slide_width

                def _tb(slide, text, l, t, w, h, fs=13, bold=False,
                        color=_RGB(30,30,30), align=_ALIGN.LEFT):
                    tb = slide.shapes.add_textbox(l, t, w, h)
                    tf = tb.text_frame; tf.word_wrap = True
                    p  = tf.paragraphs[0]; p.alignment = align
                    r  = p.add_run(); r.text = text
                    r.font.size = _Pt(fs); r.font.bold = bold
                    r.font.color.rgb = color

                hdr = summary_slide.shapes.add_shape(1, 0, 0, sw, Inches(1.05))
                hdr.fill.solid(); hdr.fill.fore_color.rgb = _RGB(0x1B, 0x3A, 0x6B)
                hdr.line.fill.background()
                _tb(summary_slide, "Translation Update — Change Summary",
                    Inches(0.4), Inches(0.22), sw - Inches(0.8), Inches(0.65),
                    fs=20, bold=True, color=_RGB(255,255,255))

                stats  = [("Total Slides", total_new), ("수정", len(changed_indices)), ("신규", len(new_indices))]
                c_fill = [_RGB(0x1B,0x3A,0x6B), _RGB(0xFF,0xC0,0x00), _RGB(0x2E,0x7D,0x32)]
                c_text = [_RGB(255,255,255), _RGB(0x33,0x1A,0x00), _RGB(255,255,255)]
                cw, ch = Inches(2.0), Inches(1.0)
                for i, (lbl, val) in enumerate(stats):
                    cl = Inches(0.4) + i * (cw + Inches(0.25))
                    ct = Inches(1.35)
                    card = summary_slide.shapes.add_shape(1, cl, ct, cw, ch)
                    card.fill.solid(); card.fill.fore_color.rgb = c_fill[i]
                    card.line.fill.background()
                    _tb(summary_slide, str(val),
                        cl+Inches(0.1), ct+Inches(0.04), cw-Inches(0.2), Inches(0.55),
                        fs=28, bold=True, color=c_text[i], align=_ALIGN.CENTER)
                    _tb(summary_slide, lbl,
                        cl+Inches(0.1), ct+Inches(0.58), cw-Inches(0.2), Inches(0.35),
                        fs=10, color=c_text[i], align=_ALIGN.CENTER)

                nums_str = ", ".join([f"Slide {i+2}" for i in all_delta_indices])
                _tb(summary_slide, f"변경/신규 slides:  {nums_str}",
                    Inches(0.4), Inches(2.55), sw-Inches(0.8), Inches(0.4), fs=11)
                _tb(summary_slide,
                    "Slides marked with UPDATED badge (top-right) contain revised translations.",
                    Inches(0.4), Inches(3.05), sw-Inches(0.8), Inches(0.4),
                    fs=11, color=_RGB(0x55,0x55,0x55))

                # ── 슬라이드별 변경 전/후 텍스트 블록 ────────
                y_cursor = Inches(3.6)
                for si in all_delta_indices:
                    info         = delta_report[si]
                    before_lines = info["before"]
                    after_lines  = info["after"]
                    label        = "🆕 NEW" if info["type"] == "new" else "🔸 UPDATED"
                    v1_ref       = f" (← v1 Slide {info['v1_idx']+1})" if info["v1_idx"] is not None else ""

                    _tb(summary_slide, f"{label}  Slide {si+2}{v1_ref}",
                        Inches(0.4), y_cursor, Inches(4.0), Inches(0.28),
                        fs=10, bold=True, color=_RGB(0x1B,0x3A,0x6B))
                    y_cursor += Inches(0.3)

                    before_text = " / ".join(before_lines[:6]) if before_lines else "(없음)"
                    _tb(summary_slide, f"[전] {before_text}",
                        Inches(0.5), y_cursor, sw - Inches(1.0), Inches(0.32),
                        fs=9, color=_RGB(0x99,0x99,0x99))
                    y_cursor += Inches(0.33)

                    after_text = " / ".join(after_lines[:6]) if after_lines else "(없음)"
                    _tb(summary_slide, f"[후] {after_text}",
                        Inches(0.5), y_cursor, sw - Inches(1.0), Inches(0.32),
                        fs=9, bold=True, color=_RGB(0x1B,0x3A,0x6B))
                    y_cursor += Inches(0.42)

                    if y_cursor > Inches(6.8):
                        remaining = len(all_delta_indices) - all_delta_indices.index(si) - 1
                        _tb(summary_slide, f"... 외 {remaining}개 슬라이드",
                            Inches(0.5), y_cursor, sw - Inches(1.0), Inches(0.3),
                            fs=9, color=_RGB(0x99,0x99,0x99))
                        break

                # ── 저장 & 다운로드 ────────────────────────
                output = io.BytesIO()
                prs_new.save(output)
                output.seek(0)
                elapsed  = _time.time() - start_time
                out_name = new_original_file.name.rsplit(".", 1)[0] + "_EN_delta.pptx"

                progress_bar.progress(1.0)
                status_text.empty()
                st.success(f"🎉 Delta 번역 완료! — 수정 {len(changed_indices)}개 · 신규 {len(new_indices)}개 재번역 (⏱️ {elapsed:.1f}초)")
                st.download_button(
                    label=f"⬇️ {out_name} 다운로드",
                    data=output,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                    type="primary",
                )

# ══════════════════════════════════════════════════════════
# QC Helper Functions
# ══════════════════════════════════════════════════════════

QC_STATUS_ICONS = {"unchecked": "⬜", "ok": "✅", "warn": "⚠️", "fix": "❌"}
QC_STATUS_LABELS = {"unchecked": "미확인", "ok": "OK", "warn": "확인 필요", "fix": "수정 필요"}
QC_STATUS_COLORS = {
    "ok": ("#F0FDF4", "#166534", "#BBF7D0"),
    "warn": ("#FFFBEB", "#92400E", "#FDE68A"),
    "fix": ("#FEF2F2", "#991B1B", "#FECACA"),
    "unchecked": ("#F3F4F6", "#6B7280", "#E5E7EB"),
}
AI_IMAGE_MAX_WIDTH = 1200

def qc_process_pdf(file_bytes, with_thumbs=False):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    pages = []
    aspect = 0.5625
    for i in range(len(doc)):
        page = doc[i]
        if i == 0:
            aspect = page.rect.height / page.rect.width
        pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=90, optimize=True)
        img_b64 = base64.b64encode(buf.getvalue()).decode()
        thumb_b64 = ""
        if with_thumbs:
            tpix = page.get_pixmap(matrix=fitz.Matrix(0.35, 0.35), alpha=False)
            timg = Image.frombytes("RGB", (tpix.width, tpix.height), tpix.samples)
            tbuf = io.BytesIO()
            timg.save(tbuf, format="JPEG", quality=55, optimize=True)
            thumb_b64 = base64.b64encode(tbuf.getvalue()).decode()
        pages.append({"image_b64": img_b64, "thumb_b64": thumb_b64})
    doc.close()
    return pages, aspect

def qc_resize_for_ai(img_b64, max_w=AI_IMAGE_MAX_WIDTH):
    raw = base64.b64decode(img_b64)
    img = Image.open(io.BytesIO(raw))
    if img.width > max_w:
        ratio = max_w / img.width
        img = img.resize((max_w, int(img.height * ratio)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=80)
    return base64.b64encode(buf.getvalue()).decode()

def qc_ai_review_page(client, ko_b64, en_b64, page_num, glossary_str=""):
    ko_small = qc_resize_for_ai(ko_b64)
    en_small = qc_resize_for_ai(en_b64)

    glossary_section = ""
    if glossary_str:
        glossary_section = f"""
## Glossary (참고용 — 절대 규칙이 아님):
{glossary_str}

Glossary 관련 중요 규칙:
- Glossary는 **참고 자료**입니다. 문맥에 따라 다른 번역이 더 적절할 수 있습니다.
- **번역이 Glossary와 일치하면 절대 지적하지 마세요.** 일치하는 걸 지적하는 건 오탐입니다.
- Glossary 용어가 더 긴 문장의 **일부**로 포함된 경우, 해당 부분이 올바르게 번역되었으면 OK입니다. 예: "연초 실적 점검 및 향후 관리 방안" → "Review of Early-year Performance and Management Plan"은 "연초 실적 점검"이 올바르게 들어가 있으므로 정상입니다.
- 같은 한국어 용어라도 문맥에 따라 다른 영어 번역이 자연스러울 수 있습니다. 예: "신작"이 Glossary에 "New IP"로 되어 있어도, 문맥상 "New Titles"가 더 적절하면 그것도 OK입니다.
- Glossary 불일치를 지적하려면, 해당 번역이 **실제로 의미가 잘못되었거나 혼란을 줄 때**만 "info" 레벨로 참고 사항으로 남기세요. "warn" 이상은 안 됩니다."""

    response = client.messages.create(
        model=MODEL,
        max_tokens=2048,
        system=f"""You are a senior translation QC reviewer for KRAFTON's board of directors (이사회) meeting materials.
You receive two slide images: first is the Korean original, second is the English translation.

## 핵심 원칙: 의역(free translation)은 정상입니다
이사회 자료의 영문 번역은 한국어를 그대로 직역하는 것이 아니라, 영어권 이사가 자연스럽게 읽을 수 있도록 의역하는 것이 올바른 방법입니다.
- 문장 구조가 달라도 **의미가 같으면 OK**
- 한국어를 더 간결하게 줄여서 번역해도 **핵심 내용이 보존되면 OK**
- 영어에 자연스러운 표현으로 바꾼 것은 좋은 번역이지, 오류가 아닙니다
- "~하였습니다" → "achieved" 같은 문체 변환은 당연히 OK

## 진짜 오류만 잡으세요 ("fix" 판정):
- ❌ 숫자/금액이 틀린 경우 (1,200억 → 1,020B 등)
- ❌ 표/차트의 데이터가 원본과 다른 경우
- ❌ 핵심 내용이 완전히 누락된 경우 (문단 통째로 빠짐)
- ❌ 의미가 정반대로 번역된 경우
- ❌ 고유명사/인명이 잘못된 경우

## 경미한 확인 사항 ("warn" 판정):
- ⚠️ 차트 범례/축 라벨이 번역되지 않고 한국어로 남아있는 경우
- ⚠️ 영문 텍스트가 텍스트 박스를 넘쳐 잘린 것으로 보이는 경우
- ⚠️ 같은 문서 내에서 동일 용어가 다른 영어로 번역된 일관성 문제 (확실할 때만)

## 이런 건 절대 지적하지 마세요:
- 문장 구조가 다른 것 (의역)
- 한국어보다 짧거나 긴 것 (자연스러운 영어 표현)
- 어순이 바뀐 것
- 불릿 포인트나 줄바꿈 위치가 다른 것
- 표현 방식의 차이 (능동↔수동, 명사형↔동사형 등)
- Glossary와 일치하는 번역 (일치하는데 지적하면 오탐!)
- Glossary 용어가 더 긴 문구 안에 올바르게 포함된 경우
- 문맥상 Glossary와 다른 번역이 더 자연스러운 경우
{glossary_section}

## 판정 기준:
- "ok": 번역이 적절함 (대부분의 슬라이드는 이 판정이어야 합니다)
- "warn": 사소하지만 확인이 필요한 부분이 있음
- "fix": 명백한 오류가 있어 수정이 필요함

대부분의 슬라이드는 "ok"여야 합니다. 전문 번역가가 작업한 결과물이므로, 진짜 문제가 있을 때만 지적하세요.

Return JSON:
{{"verdict":"ok"|"warn"|"fix","summary":"한국어 한줄 요약","issues":[{{"level":"error"|"warn"|"info","detail":"한국어로 구체적 설명"}}]}}
Return ONLY valid JSON. No markdown fences.""",
        messages=[{"role": "user", "content": [
            {"type": "text", "text": f"슬라이드 {page_num} 검토. 첫 번째=한국어 원본, 두 번째=영문 번역."},
            {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": ko_small}},
            {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": en_small}},
        ]}],
    )
    raw = response.content[0].text.strip().replace("```json", "").replace("```", "").strip()
    return json.loads(raw)

def qc_ai_review_all():
    from concurrent.futures import ThreadPoolExecutor, as_completed

    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    client = anthropic.Anthropic(api_key=api_key)
    reviews = {}
    total = st.session_state.qc_num_pages
    progress = st.progress(0, text="AI 검토 중 (병렬 처리)...")

    glossary = get_active_glossary()
    gstr = "\n".join(f"  {k} → {v}" for k, v in glossary.items()) if glossary else ""

    def _review_one(i):
        ko_img = st.session_state.qc_pages_ko[i]["image_b64"]
        en_img = st.session_state.qc_pages_en[i]["image_b64"]
        try:
            return i, qc_ai_review_page(client, ko_img, en_img, i + 1, gstr)
        except Exception as e:
            return i, {"verdict": "warn", "summary": f"검토 실패: {e}", "issues": []}

    done_count = 0
    with ThreadPoolExecutor(max_workers=5) as pool:
        futures = {pool.submit(_review_one, i): i for i in range(total)}
        for future in as_completed(futures):
            idx, result = future.result()
            reviews[idx] = result
            st.session_state.qc_status[idx] = result.get("verdict", "unchecked")
            done_count += 1
            progress.progress(done_count / total, text=f"검토 중... {done_count}/{total}")

    progress.empty()
    return reviews

def qc_generate_txt():
    total = st.session_state.qc_num_pages
    divider = "─" * 50
    header = "═" * 50
    out = [header, "  BOD Slide QC Report", f"  {datetime.now().strftime('%Y-%m-%d %H:%M')}", header, ""]
    counts = {}
    for i in range(total):
        v = st.session_state.qc_status.get(i, "unchecked")
        counts[v] = counts.get(v, 0) + 1
    out.append(f"  전체: {total}장")
    for key in ["fix", "warn", "ok", "unchecked"]:
        c = counts.get(key, 0)
        if c > 0:
            out.append(f"  {QC_STATUS_ICONS[key]} {QC_STATUS_LABELS[key]}: {c}장")
    out.append("")
    for i in range(total):
        status = st.session_state.qc_status.get(i, "unchecked")
        review = st.session_state.qc_reviews.get(i, {})
        summary = review.get("summary", "")
        issues = review.get("issues", [])
        note = st.session_state.qc_notes.get(i, "")
        out.append(divider)
        out.append(f"  슬라이드 {i+1}  {QC_STATUS_ICONS.get(status,'')} {QC_STATUS_LABELS.get(status,'')}")
        out.append(divider)
        if summary: out.append(f"  AI 요약: {summary}")
        if issues:
            out.append("  이슈:")
            for iss in issues:
                icon = {"error":"🔴","warn":"🟡","info":"🔵"}.get(iss.get("level",""),"ℹ️")
                out.append(f"    {icon} {iss.get('detail','')}")
        elif review.get("verdict") == "ok":
            out.append("  이슈: 없음 ✅")
        if note: out.append(f"  메모: {note}")
        out.append("")
    out += [header, "  END OF REPORT", header]
    return "\n".join(out)

def qc_generate_csv():
    lines = ["슬라이드,상태,AI판정,AI요약,이슈수,이슈상세,메모"]
    for i in range(st.session_state.qc_num_pages):
        status = QC_STATUS_LABELS.get(st.session_state.qc_status.get(i,"unchecked"),"미확인")
        review = st.session_state.qc_reviews.get(i, {})
        verdict = review.get("verdict", "-")
        summary = review.get("summary", "-").replace(",", ";")
        issues = review.get("issues", [])
        details = " | ".join(f"[{x.get('level','')}] {x.get('detail','')}" for x in issues).replace(",", ";")
        note = st.session_state.qc_notes.get(i, "").replace(",", ";")
        lines.append(f"{i+1},{status},{verdict},{summary},{len(issues)},{details},{note}")
    return "\n".join(lines)

def qc_slide_html(img_b64, border="#E8EBF0"):
    return f'<div style="border-radius:6px;overflow:hidden;border:1px solid {border};box-shadow:0 1px 4px rgba(0,0,0,0.06);"><img src="data:image/jpeg;base64,{img_b64}" style="width:100%;display:block;"/></div>'


# ══════════════════════════════════════════════════════════
# TAB 2: 번역 QC
# ══════════════════════════════════════════════════════════
with tab_qc:
    if not QC_AVAILABLE:
        st.error("⚠️ QC 기능에 필요한 패키지가 설치되지 않았습니다. requirements.txt에 PyMuPDF, Pillow를 추가해주세요.")
    elif not st.session_state.qc_processed:
        # ── Upload Screen ──
        st.header("🔍 번역 QC — 한/영 슬라이드 비교 검토")
        st.caption("한국어 원본 PDF와 영문 번역 PDF를 나란히 비교하고, AI가 슬라이드 이미지를 직접 보고 번역 적절성을 검토합니다.")

        qu1, qu2 = st.columns(2, gap="large")
        with qu1:
            st.markdown("**🇰🇷 한국어 원본 PDF**")
            qc_pdf_ko = st.file_uploader("KO", type=["pdf"], key="qc_up_ko", label_visibility="collapsed")
        with qu2:
            st.markdown("**🇺🇸 영문 번역 PDF**")
            qc_pdf_en = st.file_uploader("EN", type=["pdf"], key="qc_up_en", label_visibility="collapsed")

        if qc_pdf_ko and qc_pdf_en:
            if st.button("🔍 비교 시작", type="primary", use_container_width=True, key="qc_start"):
                prog = st.progress(0, text="한국어 PDF 처리 중...")
                ko_pages, aspect = qc_process_pdf(qc_pdf_ko.read(), with_thumbs=True)
                prog.progress(50, text="영문 PDF 처리 중...")
                en_pages, _ = qc_process_pdf(qc_pdf_en.read())
                prog.progress(100); prog.empty()
                if len(ko_pages) != len(en_pages):
                    st.error(f"⚠️ 페이지 수 불일치: KR {len(ko_pages)}p / EN {len(en_pages)}p")
                else:
                    st.session_state.qc_pages_ko = ko_pages
                    st.session_state.qc_pages_en = en_pages
                    st.session_state.qc_num_pages = len(ko_pages)
                    st.session_state.qc_aspect = aspect
                    st.session_state.qc_current = 0
                    st.session_state.qc_mode = "compare"
                    st.session_state.qc_processed = True
                    st.session_state.qc_status = {i: "unchecked" for i in range(len(ko_pages))}
                    st.session_state.qc_notes = {}
                    st.session_state.qc_reviews = {}
                    st.rerun()
        elif qc_pdf_ko or qc_pdf_en:
            st.caption("한국어 PDF와 영문 PDF를 모두 올려주세요.")

    else:
        # ── QC Viewer ──
        _total = st.session_state.qc_num_pages
        _cur = st.session_state.qc_current
        _mode = st.session_state.qc_mode
        _api = st.secrets.get("ANTHROPIC_API_KEY", "")
        if _cur >= _total:
            _cur = 0; st.session_state.qc_current = 0

        # Header buttons
        qh1, qh2 = st.columns([8, 3])
        with qh1:
            qc1, qc2, qc3, qc4, qc5 = st.columns([1,1,1,1,1])
            with qc1:
                if st.button("🔀 비교", use_container_width=True, key="qm_cmp",
                             type="primary" if _mode=="compare" else "secondary"):
                    st.session_state.qc_mode = "compare"; st.rerun()
            with qc2:
                if st.button("🇰🇷 한국어", use_container_width=True, key="qm_ko",
                             type="primary" if _mode=="ko" else "secondary"):
                    st.session_state.qc_mode = "ko"; st.rerun()
            with qc3:
                if st.button("🇺🇸 English", use_container_width=True, key="qm_en",
                             type="primary" if _mode=="en" else "secondary"):
                    st.session_state.qc_mode = "en"; st.rerun()
            with qc4:
                if st.button("📋 검토 결과", use_container_width=True, key="qm_rv",
                             type="primary" if _mode=="review" else "secondary"):
                    st.session_state.qc_mode = "review"; st.rerun()
            with qc5:
                if _api:
                    if st.button("🤖 전체 AI 검토", use_container_width=True, key="qm_ai", type="primary"):
                        st.session_state.qc_reviews = qc_ai_review_all()
                        st.session_state.qc_mode = "review"; st.rerun()

        with qh2:
            if st.session_state.qc_reviews:
                qdl1, qdl2 = st.columns(2)
                with qdl1:
                    st.download_button("📝 TXT", qc_generate_txt(),
                                       file_name=f"BOD_QC_{datetime.now().strftime('%Y%m%d')}.txt",
                                       mime="text/plain", use_container_width=True, key="qc_dl_txt")
                with qdl2:
                    st.download_button("📊 CSV", qc_generate_csv(),
                                       file_name=f"BOD_QC_{datetime.now().strftime('%Y%m%d')}.csv",
                                       mime="text/csv", use_container_width=True, key="qc_dl_csv")

        # ── Review results mode ──
        if _mode == "review":
            if not st.session_state.qc_reviews:
                st.info("🤖 '전체 AI 검토' 버튼을 눌러 검토를 시작하세요.")
            else:
                # ── Summary ──
                _counts = {}
                for s in st.session_state.qc_status.values():
                    _counts[s] = _counts.get(s, 0) + 1
                sc1, sc2, sc3, sc4, sc5 = st.columns([1,1,1,1,2])
                with sc1:
                    st.metric("전체", f"{_total}장")
                with sc2:
                    st.metric("✅ OK", f"{_counts.get('ok',0)}장")
                with sc3:
                    st.metric("⚠️ 확인", f"{_counts.get('warn',0)}장")
                with sc4:
                    st.metric("❌ 수정", f"{_counts.get('fix',0)}장")
                with sc5:
                    _filter = st.radio("필터", ["전체","❌ 수정 필요","⚠️ 확인 필요","✅ OK"],
                                        horizontal=True, label_visibility="collapsed", key="qc_filter")

                _fmap = {"전체":None, "❌ 수정 필요":"fix", "⚠️ 확인 필요":"warn", "✅ OK":"ok"}
                _af = _fmap[_filter]

                # ── Per-slide: images + review stacked vertically ──
                for i in range(_total):
                    _st = st.session_state.qc_status.get(i, "unchecked")
                    if _af and _st != _af:
                        continue
                    _rv = st.session_state.qc_reviews.get(i, {})
                    _vd = _rv.get("verdict", "unchecked")
                    _sm = _rv.get("summary", "")
                    _is = _rv.get("issues", [])
                    _nt = st.session_state.qc_notes.get(i, "")
                    bg, fg, bd = QC_STATUS_COLORS.get(_vd, QC_STATUS_COLORS["unchecked"])

                    st.divider()

                    # Header
                    st.markdown(
                        f'<div style="display:flex;align-items:center;gap:12px;margin-bottom:10px;">'
                        f'<span style="font-size:16px;font-weight:700;color:#111827;">슬라이드 {i+1}</span>'
                        f'<span style="font-size:12px;font-weight:600;padding:3px 14px;border-radius:20px;'
                        f'background:{bg};color:{fg};border:1px solid {bd};">'
                        f'{QC_STATUS_ICONS.get(_vd,"")} {QC_STATUS_LABELS.get(_vd,"")}</span></div>',
                        unsafe_allow_html=True)

                    # KR / EN slide images side by side
                    ko_b64 = st.session_state.qc_pages_ko[i]["image_b64"]
                    en_b64 = st.session_state.qc_pages_en[i]["image_b64"]
                    img_col1, img_col2 = st.columns(2, gap="small")
                    with img_col1:
                        st.caption("🇰🇷 한국어")
                        st.image(base64.b64decode(ko_b64), use_container_width=True)
                    with img_col2:
                        st.caption("🇺🇸 English")
                        st.image(base64.b64decode(en_b64), use_container_width=True)

                    # AI review results
                    if _sm:
                        st.markdown(f"**🤖 AI 검토:** {_sm}")

                    if _is:
                        for iss in _is:
                            lv = iss.get("level", "info")
                            dt = iss.get("detail", "")
                            styles = {
                                "error": "background:#FEF2F2;border-left:3px solid #EF4444;color:#991B1B",
                                "warn": "background:#FFFBEB;border-left:3px solid #F59E0B;color:#92400E",
                                "info": "background:#EFF6FF;border-left:3px solid #3B82F6;color:#1E40AF",
                            }
                            ic = {"error": "🔴", "warn": "🟡", "info": "🔵"}.get(lv, "ℹ️")
                            st.markdown(
                                f'<div style="padding:8px 12px;margin:4px 0;border-radius:6px;'
                                f'font-size:12px;line-height:1.6;{styles.get(lv,"")}">{ic} {dt}</div>',
                                unsafe_allow_html=True)
                    elif _vd == "ok":
                        st.markdown(
                            '<div style="padding:6px 12px;border-radius:6px;font-size:12px;'
                            'background:#F0FDF4;border-left:3px solid #22C55E;color:#166534;">'
                            '✅ 번역 적절</div>', unsafe_allow_html=True)

                    # Note
                    new_note = st.text_input(
                        f"📝 슬라이드 {i+1} 메모", value=_nt, key=f"qcn_{i}",
                        placeholder="메모 입력...", label_visibility="collapsed")
                    if new_note != _nt:
                        st.session_state.qc_notes[i] = new_note

        # ── Slide view modes ──
        elif _mode in ("compare", "ko", "en"):
            # Render slide
            if _mode == "compare":
                ko_img = st.session_state.qc_pages_ko[_cur]["image_b64"]
                en_img = st.session_state.qc_pages_en[_cur]["image_b64"]
                html = f'''<div style="display:flex;gap:14px;width:100%;">
                    <div style="flex:1;min-width:0;"><div style="font-size:12px;font-weight:600;color:#374151;text-align:center;padding:4px 0 8px;">🇰🇷 한국어</div>{qc_slide_html(ko_img)}</div>
                    <div style="flex:1;min-width:0;"><div style="font-size:12px;font-weight:600;color:#4F46E5;text-align:center;padding:4px 0 8px;">🇺🇸 English</div>{qc_slide_html(en_img,"#D1D5F0")}</div>
                </div>'''
                st.components.v1.html(html, height=int(520*st.session_state.qc_aspect)+40, scrolling=False)
            else:
                _pages = st.session_state.qc_pages_ko if _mode=="ko" else st.session_state.qc_pages_en
                _img = _pages[_cur]["image_b64"]
                _bdr = "#E8EBF0" if _mode=="ko" else "#D1D5F0"
                _lbl = "🇰🇷 한국어" if _mode=="ko" else "🇺🇸 English"
                _clr = "#374151" if _mode=="ko" else "#4F46E5"
                html = f'<div style="max-width:1000px;margin:0 auto;"><div style="font-size:12px;font-weight:600;color:{_clr};text-align:center;padding:4px 0 8px;">{_lbl}</div>{qc_slide_html(_img,_bdr)}</div>'
                st.components.v1.html(html, height=int(960*st.session_state.qc_aspect)+40, scrolling=False)

            # Status controls
            _st = st.session_state.qc_status.get(_cur, "unchecked")
            sr1, sr2 = st.columns([5, 6])
            with sr1:
                st.markdown(f"**슬라이드 {_cur+1} 상태**")
                ss1,ss2,ss3,ss4 = st.columns(4)
                with ss1:
                    if st.button("✅ OK", key=f"qs_ok_{_cur}", use_container_width=True,
                                 type="primary" if _st=="ok" else "secondary"):
                        st.session_state.qc_status[_cur]="ok"; st.rerun()
                with ss2:
                    if st.button("⚠️ 확인", key=f"qs_w_{_cur}", use_container_width=True,
                                 type="primary" if _st=="warn" else "secondary"):
                        st.session_state.qc_status[_cur]="warn"; st.rerun()
                with ss3:
                    if st.button("❌ 수정", key=f"qs_f_{_cur}", use_container_width=True,
                                 type="primary" if _st=="fix" else "secondary"):
                        st.session_state.qc_status[_cur]="fix"; st.rerun()
                with ss4:
                    if st.button("⬜ 초기화", key=f"qs_u_{_cur}", use_container_width=True, type="secondary"):
                        st.session_state.qc_status[_cur]="unchecked"; st.rerun()
            with sr2:
                _nt = st.session_state.qc_notes.get(_cur, "")
                _new = st.text_input("📝", value=_nt, key=f"qcsn_{_cur}", placeholder="메모...", label_visibility="collapsed")
                if _new != _nt: st.session_state.qc_notes[_cur] = _new

            # Show AI review if exists
            _rv = st.session_state.qc_reviews.get(_cur)
            if _rv:
                st.markdown(f"**🤖 AI:** {QC_STATUS_ICONS.get(_rv.get('verdict',''),'')} {_rv.get('summary','')}")
                for iss in _rv.get("issues",[]):
                    ic = {"error":"🔴","warn":"🟡","info":"🔵"}.get(iss.get("level",""),"ℹ️")
                    st.caption(f"{ic} {iss.get('detail','')}")

            # Navigation
            qn1, qn2, qn3, qn4, qn5 = st.columns([3,1,1,1,3])
            with qn2:
                if st.button("◀ 이전", disabled=(_cur==0), use_container_width=True, key="qc_prev"):
                    st.session_state.qc_current = _cur-1; st.rerun()
            with qn3:
                st.markdown(f'<div style="text-align:center;padding:8px 0;font-size:14px;color:#6B7280;font-weight:500;">{_cur+1}/{_total}</div>', unsafe_allow_html=True)
            with qn4:
                if st.button("다음 ▶", disabled=(_cur==_total-1), use_container_width=True, key="qc_next"):
                    st.session_state.qc_current = _cur+1; st.rerun()

        # Reset button
        st.divider()
        if st.button("↻ 새 파일로 교체", use_container_width=True, key="qc_reset"):
            for k in list(st.session_state.keys()):
                if k.startswith("qc_"):
                    del st.session_state[k]
            st.rerun()


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
