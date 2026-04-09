import streamlit as st
import re, json, time, io, os
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import anthropic

# ── 페이지 설정 ────────────────────────────────────────────
st.set_page_config(
    page_title="KRAFTON PPT Translator",
    page_icon="🏢",
    layout="centered"
)

st.title("🏢 KRAFTON BOD PPT Translator")
st.caption("PPT 파일을 업로드하면 AI가 자동으로 번역합니다.")

# ── 사이드바: 설정 ─────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 설정")
    api_key = st.text_input("Claude API Key", type="password",
                             help="Anthropic API Key를 입력하세요")
    target_lang = st.selectbox("번역 언어", ["English", "Japanese", "Chinese"])
    min_font_pt = st.slider("최소 폰트 크기 (pt)", min_value=5, max_value=12, value=7)
    st.divider()
    st.caption("© KRAFTON BOD Team")

# ── 상수 ──────────────────────────────────────────────────
MODEL = "claude-sonnet-4-20250514"
LANG_MAP = {"English": "English", "Japanese": "Japanese", "Chinese": "Traditional Chinese"}
SKIP_NAMES = ("slide number", "date", "footer", "page")
KOREAN_FONTS = {
    "Pretendard", "나눔고딕", "맑은 고딕", "Malgun Gothic",
    "NanumGothic", "KoPubWorldBatang", "굴림", "돋움", "바탕", "나눔명조"
}
EN_FONT = "Calibri"

GLOSSARY = {
    # ── 인명 ──────────────────────────────────────────────
    "장병규": "BG Chang", "김창한": "CH Kim", "배동근": "DK Bae",
    "장태석": "TS Jang", "오진호": "Jin Oh", "이강욱": "Kangwook Lee",
    "박혜리": "Maria Park", "윤상훈": "Richard Yoon", "손현일": "Sean Sohn",
    "한소영": "Soyoung Han", "이병욱": "Andy Lee", "김정연": "Jung Yun Kim",
    "김낙형": "Harns Kim", "박찬민": "Chanmin Park", "홍정택": "Jeongtack Hong",
    "박재민": "Jaemin Park", "강재형": "Jaehyung Kang", "박상훈": "Albert Park",
    "정용": "Young Chung", "윤수진": "Sujin Yun", "최승환": "Seunghwan Choi",
    "이동훈": "Donghun Lee", "조혁일": "Hyuk Cho", "유재영": "Jaeyoung Yoo",
    "황동경": "Dongkyung Hwang", "김현아": "Hyuna Kim", "이석현": "Matthew Lee",
    "이창준": "Changjun Lee", "이지은": "Jieun Lee", "현종민": "Jongmin Hyun",
    "김정화": "Junghwa Kim", "신소희": "Sohee Shin", "노경원": "Kyoungwon Noh",
    "정지현": "Jeehyun Jung", "김지영": "Jiyoung Kim", "박서훈": "Seohoon Park",
    "김인영": "Inyoung Kim", "김고운": "Kowoon Kim",
    # ── 재무 용어 ──────────────────────────────────────────
    "매출": "Sales", "영업이익": "Operating Profit (OP)", "영업이익률": "OPM",
    "회계 매출": "Reported Revenue", "재무계획": "Financial Plan", "예산": "Budget",
    "연초 실적 점검": "Review of Early-year Performance", "전년 대비": "YoY",
    "누적 실적": "Cumulative Performance", "계획 대비": "vs. Plan",
    "인건비": "Labor Costs", "마케팅비": "Marketing Expenses",
    "지급수수료": "Service Fees", "임차료/상각비": "Rental/Depreciation",
    "비통제 비용": "Uncontrollable Costs", "매출연동비": "App Fees / Cost of Sales",
    "예비비": "Reserve Fund", "손익": "P/L",
    "범 크래프톤": "KRAFTON Family", "크래프톤 Business": "KRAFTON Business",
    "비게임 프로젝트": "Non-gaming Projects", "전사 비용": "Company-wide Expenses",
    "예산 집중 관리": "Intensive Budget Control",
    # ── 전략/비즈니스 용어 ─────────────────────────────────
    "핵심 서비스": "Core Service", "신작": "New IP", "신규 IP": "New IP",
    "장기 PLC화": "long-term PLC", "직접 서비스": "Direct Service",
    "두 자릿수 성장": "Double-digit Growth", "계단식 성장": "Stepwise Growth",
    "파이프라인": "Pipeline", "보고": "Report", "승인/보고": "Approval/Report",
    "승인사항": "Approval Item", "보류": "Hold (Deferred)",
    "서면 보고": "Written Report", "사전공유": "Pre-sharing",
    "서비스 종료": "Sunset / Discontinue",
    "퍼블리싱 운영 최적화 방안": "Publishing Operations Optimization Plan",
    "운영 Agility": "Operational Agility", "체질 개선": "Operational Transformation",
    "런치패드 프로그램": "Launchpad Program", "Tentpole 캠페인": "Tentpole Campaign",
    # ── 조직/HR ────────────────────────────────────────────
    "이사회": "Board of Directors (BOD)", "대표이사": "CEO",
    "이사회 의장": "Board Chair", "사외이사": "Outside Director",
    "제작 리더십": "Production Leadership", "제작 리더": "Production Lead",
    "제작총괄": "Head of Production", "비제작 조직": "Non-development Org",
    "신사업 자회사": "New Business Subsidiaries", "전문계약직": "Professional Contractors",
    "사내채용": "Internal Hiring", "인력계획": "Workforce Plan",
    "증원": "Headcount Increase", "유휴인력": "Idle Employees",
    "低성과자": "Low-performing Employees", "자발적 퇴직": "Voluntary Exit",
    "희망퇴직": "Voluntary Retirement", "정규직": "Full-time Employee",
}

# ── 핵심 함수들 ────────────────────────────────────────────

def should_skip(shape):
    name = getattr(shape, "name", "").lower()
    return any(k in name for k in SKIP_NAMES)

def has_korean(text):
    return bool(re.search(r"[\uAC00-\uD7A3]", text))

def iter_paragraphs(shapes):
    for shape in shapes:
        if getattr(shape, "has_text_frame", False) and shape.text_frame:
            for para in shape.text_frame.paragraphs:
                yield shape, para
        if getattr(shape, "has_table", False) and shape.table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if getattr(cell, "text_frame", None):
                        for para in cell.text_frame.paragraphs:
                            yield shape, para
        if getattr(shape, "shape_type", None) == 6:
            yield from iter_paragraphs(shape.shapes)

def get_shape_width_pt(shape):
    try:
        w_emu = shape.width
        if w_emu:
            tf = getattr(shape, "text_frame", None)
            left_in  = tf.margin_left  if (tf and tf.margin_left)  else 91440
            right_in = tf.margin_right if (tf and tf.margin_right) else 91440
            return max((w_emu - left_in - right_in) / 12700, 10)
    except:
        pass
    return None

def get_slide_texts(slide):
    result = []
    for global_idx, (shape, para) in enumerate(iter_paragraphs(slide.shapes)):
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
            "global_idx": global_idx,
            "text": full,
            "font_pt": font_pt,
            "box_width_pt": get_shape_width_pt(shape),
        })
    return result

def detect_slide_type(texts):
    combined = " ".join(t["text"] for t in texts)
    if any(k in combined for k in ["승인사항", "승인 요청", "의결"]):
        return "approval"
    if any(k in combined for k in ["억원", "매출", "손익", "YoY", "YTD", "실적"]):
        return "financial"
    if any(k in combined for k in ["출시", "파이프라인", "Q1", "Q2", "Q3", "Q4"]):
        return "timeline"
    if "사전공유" in combined:
        return "pre-disclosure"
    return "strategy"

def estimate_max_chars(t):
    box_w = t.get("box_width_pt")
    font  = t.get("font_pt") or 12
    if box_w and font:
        chars_per_line = int(box_w / (font * 0.55))
        lines = max(1, len(t["text"].split("\n")))
        return chars_per_line * lines
    return int(len(t["text"].replace("\n", "")) * 1.3)

def translate_slide(client, texts, slide_type, target_lang_str):
    if not texts:
        return {}
    gstr = "\n".join(f"  {k} → {v}" for k, v in GLOSSARY.items())
    type_rules = {
        "financial":      "Keep all numbers/units exact. Translate labels/headers only.",
        "approval":       'Use formal request language: "We request approval for..."',
        "strategy":       "Use concise noun phrases for bullets. Lead with the conclusion.",
        "timeline":       "Translate: 출시예정→Scheduled, 지연→Delayed, 완료→Released",
        "pre-disclosure": 'Add "(Pre-Disclosure)" prefix if not present. No vote required.',
    }
    type_hint = type_rules.get(slide_type, "")
    input_map = {
        str(i): {"text": t["text"], "max_chars": estimate_max_chars(t)}
        for i, t in enumerate(texts)
    }
    prompt = f"""Translate Korean PowerPoint slide texts to {target_lang_str}.
Slide type: {slide_type}
{type_hint}

Mandatory terms (use EXACTLY as shown):
{gstr}

Always keep unchanged:
- Numbers, %, financial figures
- Game titles: PUBG, BGMI, OVERDARE, Black Budget, Valor, inZOI
- Company names: KRAFTON, Unknown Worlds, Neon Giant
- Labels: DRI, SL, n/a, As-Is, To-Be, □, ■

Input JSON (index → {{text, max_chars}}):
{json.dumps(input_map, ensure_ascii=False, indent=2)}

Rules:
1. Return ONLY valid JSON with same keys, translated string values
2. Preserve \\n line breaks exactly
3. Formal board meeting English (not casual)
4. Be concise — PowerPoint bullets, not paragraphs
5. Each entry has 'text' (to translate) and 'max_chars' (character budget).
   If over budget, shorten with synonyms or drop filler words. Never omit key meaning.
6. Return JSON: same index keys, string values only. e.g. {{"0": "Translated text"}}
7. No markdown, no explanation. JSON only

Output:"""

    res = client.messages.create(
        model=MODEL,
        max_tokens=3000,
        system="You are a precise JSON-output PPT translator for KRAFTON board of directors. "
               "Return ONLY valid JSON. No markdown. No explanation.",
        messages=[{"role": "user", "content": prompt}]
    )
    raw = res.content[0].text.strip()
    match = re.search(r"\{[\s\S]*\}", raw)
    if not match:
        raise ValueError("JSON 파싱 실패")
    return json.loads(match.group())

def replace_para_text(para, new_text, shape=None, min_pt=7):
    if not new_text:
        return
    orig_font_size = None
    orig_font_name = None
    for run in para.runs:
        if run.font.size:
            orig_font_size = run.font.size.pt
        if run.font.name:
            orig_font_name = run.font.name
        break

    base_pt  = orig_font_size or 12
    final_pt = base_pt
    if shape is not None and base_pt > min_pt:
        try:
            w_emu = shape.width
            if w_emu:
                tf = getattr(shape, "text_frame", None)
                left_in  = tf.margin_left  if (tf and tf.margin_left)  else 91440
                right_in = tf.margin_right if (tf and tf.margin_right) else 91440
                box_w = max((w_emu - left_in - right_in) / 12700, 10)
                lines = new_text.split("\n")
                max_line_len = max(len(l) for l in lines) if lines else len(new_text)
                if max_line_len > 0:
                    required_font = box_w / (max_line_len * 0.55)
                    if required_font < base_pt:
                        final_pt = max(round(required_font, 1), min_pt)
        except:
            pass

    for i, run in enumerate(para.runs):
        if i == 0:
            run.text = new_text
            run.font.size = Pt(final_pt)
            if orig_font_name and orig_font_name in KOREAN_FONTS:
                run.font.name = EN_FONT
        else:
            run.text = ""

    if not para.runs:
        from pptx.oxml.ns import qn
        from lxml import etree
        r   = etree.SubElement(para._p, qn("a:r"))
        rPr = etree.SubElement(r, qn("a:rPr"), attrib={"lang": "en-US"})
        t   = etree.SubElement(r, qn("a:t"))
        t.text = new_text

# ── 메인 UI ────────────────────────────────────────────────

uploaded_file = st.file_uploader("PPT 파일 업로드", type=["pptx"],
                                  help=".pptx 파일만 지원합니다")

if uploaded_file and not api_key:
    st.warning("사이드바에서 Claude API Key를 먼저 입력해주세요.")

if uploaded_file and api_key:
    st.success(f"✅ **{uploaded_file.name}** 업로드 완료")

    if st.button("🚀 번역 시작", type="primary", use_container_width=True):

        client = anthropic.Anthropic(api_key=api_key)
        target_lang_str = LANG_MAP.get(target_lang, "English")

        # PPT 로드
        pptx_bytes = uploaded_file.read()
        prs = Presentation(io.BytesIO(pptx_bytes))

        # 텍스트 파싱
        with st.spinner("📊 텍스트 파싱 중..."):
            all_slides_info = []
            total_ko = 0
            for slide in prs.slides:
                texts = get_slide_texts(slide)
                all_slides_info.append(texts)
                total_ko += len(texts)

        st.info(f"총 {len(prs.slides)}장 · 번역 대상 단락 **{total_ko}개**")

        # 번역 실행
        st.write("🔄 번역 중...")
        progress_bar = st.progress(0)
        status_text  = st.empty()
        log_area     = st.empty()

        all_translations = []
        log_lines = []
        total_slides = len(all_slides_info)

        for si, texts in enumerate(all_slides_info):
            ko_count = len(texts)
            progress_bar.progress((si + 1) / total_slides)

            if ko_count == 0:
                all_translations.append({})
                log_lines.append(f"⏭️ Slide {si+1:2d} — 한국어 없음")
            else:
                slide_type = detect_slide_type(texts)
                status_text.text(f"Slide {si+1}/{total_slides} [{slide_type}] 번역 중...")
                for attempt in range(2):
                    try:
                        translated = translate_slide(client, texts, slide_type, target_lang_str)
                        all_translations.append(translated)
                        log_lines.append(f"✅ Slide {si+1:2d} [{slide_type:15s}] — {ko_count}개")
                        break
                    except Exception as e:
                        if attempt == 0:
                            log_lines.append(f"⚠️ Slide {si+1:2d} 재시도... ({e})")
                            time.sleep(2)
                        else:
                            all_translations.append({})
                            log_lines.append(f"❌ Slide {si+1:2d} 실패 — 원본 유지")
                time.sleep(0.3)

            log_area.code("\n".join(log_lines))

        status_text.text("✅ 번역 완료! PPT 생성 중...")

        # 텍스트 교체
        for si, (slide, texts, translated_map) in enumerate(
            zip(prs.slides, all_slides_info, all_translations)
        ):
            if not translated_map:
                continue
            slide_paras = list(iter_paragraphs(slide.shapes))
            for ti, text_info in enumerate(texts):
                translated_text = translated_map.get(str(ti))
                if not translated_text:
                    continue
                global_idx = text_info["global_idx"]
                if global_idx < len(slide_paras):
                    shape, para = slide_paras[global_idx]
                    replace_para_text(para, translated_text, shape=shape, min_pt=min_font_pt)

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
            type="primary"
        )
