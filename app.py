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
    "김인영": "Inyoung Kim",   "김고운": "Kowoon Kim",
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
}

# ── Session state 초기화 ───────────────────────────────────
if "approved_glossary"      not in st.session_state:
    st.session_state.approved_glossary = {}
if "pending_glossary"       not in st.session_state:
    st.session_state.pending_glossary = []
if "session_extra_glossary" not in st.session_state:
    st.session_state.session_extra_glossary = {}
if "admin_logged_in"        not in st.session_state:
    st.session_state.admin_logged_in = False


def get_active_glossary():
    """BASE + 승인된 항목 + 이번 세션 추가 = 최종 glossary"""
    merged = dict(BASE_GLOSSARY)
    merged.update(st.session_state.approved_glossary)
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
            tf       = getattr(shape, "text_frame", None)
            left_in  = tf.margin_left  if (tf and tf.margin_left)  else 91440
            right_in = tf.margin_right if (tf and tf.margin_right) else 91440
            return max((w_emu - left_in - right_in) / 12700, 10)
    except Exception:
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
            "global_idx":   global_idx,
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


def estimate_max_chars(t):
    box_w = t.get("box_width_pt")
    font  = t.get("font_pt") or 12
    if box_w and font:
        return int(box_w / (font * 0.55)) * max(1, len(t["text"].split("\n")))
    return int(len(t["text"].replace("\n", "")) * 1.3)


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
        str(i): {"text": t["text"], "max_chars": estimate_max_chars(t)}
        for i, t in enumerate(texts)
    }
    prompt = f"""Translate Korean PowerPoint slide texts to {target_lang_str}.
Slide type: {slide_type}
{type_rules.get(slide_type, "")}

Mandatory terms (use EXACTLY as shown):
{gstr}

Always keep unchanged: Numbers, %, game titles (PUBG, BGMI, OVERDARE, Black Budget, Valor, inZOI), company names (KRAFTON), Labels (DRI, SL, □, ■)

Input JSON:
{json.dumps(input_map, ensure_ascii=False, indent=2)}

Rules:
1. Return ONLY valid JSON — same index keys, STRING values only
2. Preserve \\n line breaks exactly
3. Formal board meeting English
4. Be concise — PPT bullets, not paragraphs
5. If translation exceeds max_chars, shorten with synonyms. Never omit key meaning.
6. No markdown, no explanation. JSON only

Output:"""

    res = client.messages.create(
        model=MODEL,
        max_tokens=3000,
        system="Precise JSON-output PPT translator for KRAFTON BOD. Return ONLY valid JSON. String values only.",
        messages=[{"role": "user", "content": prompt}]
    )
    raw   = res.content[0].text.strip()
    match = re.search(r"\{[\s\S]*\}", raw)
    if not match:
        raise ValueError("JSON 파싱 실패")
    return json.loads(match.group())


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

    base_pt  = orig_font_size or 12
    final_pt = base_pt

    # ── 폰트 축소 제외 조건 ────────────────────────────────
    # 타이틀(18pt↑), 짧은 텍스트(DRI·페이지번호), 숫자/영문만(주석) → 원본 폰트 유지
    text_stripped = new_text.strip()
    is_title      = base_pt >= 18
    is_short      = len(text_stripped.replace("\n", "")) <= 30
    is_label_only = bool(re.match(r'^[\d\s\.\)\-\|/A-Za-z:]+$', text_stripped))
    skip_shrink   = is_title or is_short or is_label_only

    if not skip_shrink and shape is not None and base_pt > min_pt:
        try:
            w_emu = shape.width
            if w_emu:
                tf       = getattr(shape, "text_frame", None)
                left_in  = tf.margin_left  if (tf and tf.margin_left)  else 91440
                right_in = tf.margin_right if (tf and tf.margin_right) else 91440
                box_w    = max((w_emu - left_in - right_in) / 12700, 10)
                lines         = new_text.split("\n")
                max_line_len  = max(len(l) for l in lines) if lines else len(new_text)
                if max_line_len > 0:
                    req = box_w / (max_line_len * 0.55)
                    if req < base_pt:
                        final_pt = max(round(req, 1), min_pt)
        except Exception:
            pass

    # run 교체
    for i, run in enumerate(para.runs):
        if i == 0:
            run.text = new_text
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
tab_translate, tab_glossary, tab_admin = st.tabs([
    "🚀 번역", "📖 Glossary", "🔐 관리자"
])


# ──────────────────────────────────────────────────────────
# TAB 1: 번역
# ──────────────────────────────────────────────────────────
with tab_translate:
    st.header("🏢 KRAFTON BOD PPT Translator")
    st.caption("PPT 파일을 업로드하면 AI가 자동으로 번역합니다.")

    col1, col2 = st.columns(2)
    with col1:
        target_lang = st.selectbox("번역 언어", ["English", "Japanese", "Chinese"])
    with col2:
        min_font_pt = st.slider("최소 폰트 (pt)", 5, 12, 7)

    # API Key — Streamlit Secrets에서 자동 로드
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        st.error("⚠️ API Key 미설정. Streamlit Cloud → Settings → Secrets에 ANTHROPIC_API_KEY를 추가해주세요.")

    active_glossary = get_active_glossary()
    st.caption(
        f"현재 적용 Glossary: **{len(active_glossary)}개** 항목 "
        f"(기본 {len(BASE_GLOSSARY)} "
        f"+ 승인 {len(st.session_state.approved_glossary)} "
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

            # PPT 텍스트 교체
            status_text.text("💾 PPT 생성 중...")
            for si, (slide, texts, translated_map) in enumerate(
                zip(prs.slides, all_slides_info, all_translations)
            ):
                if not translated_map:
                    continue
                slide_paras = list(iter_paragraphs(slide.shapes))
                for ti, text_info in enumerate(texts):
                    tr = translated_map.get(str(ti))
                    # string이 아니면 skip
                    if not tr or not isinstance(tr, str):
                        continue
                    tr = tr.strip()
                    if not tr:
                        continue
                    gidx = text_info["global_idx"]
                    if gidx < len(slide_paras):
                        shape, para = slide_paras[gidx]
                        replace_para_text(para, tr, shape=shape, min_pt=min_font_pt)

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

    active   = get_active_glossary()
    approved = st.session_state.approved_glossary
    pending  = st.session_state.pending_glossary

    search = st.text_input("🔍 검색", placeholder="한국어 또는 영문으로 검색...")
    st.caption(f"총 **{len(active)}개** 항목 (기본 {len(BASE_GLOSSARY)} + 승인 {len(approved)}개)")

    NAME_KEYS = set(list(BASE_GLOSSARY.keys())[:38])

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
        st.markdown(f"#### ✅ 승인된 추가 항목 — {len(approved)}개")
        for ko, en in approved.items():
            if search and search.lower() not in ko.lower() and search.lower() not in en.lower():
                continue
            st.markdown(f"`{ko}` → {en}")

    st.divider()
    st.subheader("💡 용어 추가 제안")
    st.caption(
        "번역이 어색하거나 더 좋은 표현을 발견하면 제안해주세요. "
        "**관리자 승인 전까지는 이번 세션에만 적용**되고, 승인 후 전체 반영됩니다."
    )

    with st.form("suggest_form", clear_on_submit=True):
        c1, c2, c3 = st.columns([2, 2, 1])
        with c1:
            new_ko = st.text_input("한국어 원문", placeholder="예: 핵심 지표")
        with c2:
            new_en = st.text_input("영문 번역", placeholder="예: Key Metrics")
        with c3:
            submitter = st.text_input("제안자", placeholder="이름 (선택)")
        submitted = st.form_submit_button("📤 제안 제출", use_container_width=True)

        if submitted:
            if not new_ko.strip() or not new_en.strip():
                st.error("한국어와 영문을 모두 입력해주세요.")
            elif new_ko in BASE_GLOSSARY:
                st.warning(f"⚠️ `{new_ko}`는 이미 기본 Glossary에 있어요: **{BASE_GLOSSARY[new_ko]}**")
            else:
                st.session_state.session_extra_glossary[new_ko] = new_en
                already = any(p["ko"] == new_ko for p in st.session_state.pending_glossary)
                if not already:
                    st.session_state.pending_glossary.append({
                        "ko":   new_ko,
                        "en":   new_en,
                        "by":   submitter.strip() or "익명",
                        "time": time.strftime("%Y-%m-%d %H:%M"),
                    })
                st.success(
                    f"✅ **`{new_ko}` → {new_en}** 제안 완료! "
                    "이번 세션 번역에 즉시 적용됩니다. 관리자 승인 후 전체 반영."
                )

    if pending:
        st.info(f"⏳ 현재 관리자 승인 대기 중: **{len(pending)}개** 항목")


# ──────────────────────────────────────────────────────────
# TAB 3: 관리자
# ──────────────────────────────────────────────────────────
with tab_admin:
    st.header("🔐 관리자 페이지")

    if not st.session_state.admin_logged_in:
        st.info("관리자만 접근 가능합니다.")
        pwd = st.text_input(
            "비밀번호", type="password",
            label_visibility="collapsed",
            placeholder="관리자 비밀번호 입력"
        )
        if st.button("로그인", type="primary"):
            if pwd == ADMIN_PASSWORD:
                st.session_state.admin_logged_in = True
                st.rerun()
            else:
                st.error("❌ 비밀번호가 틀렸습니다.")
    else:
        col_t, col_l = st.columns([4, 1])
        with col_t:
            st.success("✅ 관리자로 로그인됨")
        with col_l:
            if st.button("로그아웃"):
                st.session_state.admin_logged_in = False
                st.rerun()

        pending  = st.session_state.pending_glossary
        approved = st.session_state.approved_glossary

        # 승인 대기
        st.subheader(f"⏳ 승인 대기 — {len(pending)}개")
        if not pending:
            st.info("현재 대기 중인 제안이 없습니다.")
        else:
            for idx, item in enumerate(pending):
                with st.container(border=True):
                    c1, c2, c3, c4, c5 = st.columns([2, 2, 1, 1, 1])
                    with c1:
                        st.markdown(f"**`{item['ko']}`**")
                        st.caption(f"by {item['by']} · {item['time']}")
                    with c2:
                        st.markdown(f"→ **{item['en']}**")
                    with c3:
                        st.write("")
                    with c4:
                        if st.button("✅ 승인", key=f"approve_{idx}", type="primary"):
                            approved[item["ko"]] = item["en"]
                            st.session_state.approved_glossary = approved
                            pending.pop(idx)
                            st.session_state.pending_glossary = pending
                            st.toast(f"✅ `{item['ko']}` 승인 완료!")
                            st.rerun()
                    with c5:
                        if st.button("❌ 거절", key=f"reject_{idx}"):
                            pending.pop(idx)
                            st.session_state.pending_glossary = pending
                            st.toast(f"❌ `{item['ko']}` 거절됨")
                            st.rerun()

        # 승인된 항목 관리
        st.divider()
        st.subheader(f"✅ 승인된 추가 항목 — {len(approved)}개")
        if not approved:
            st.info("아직 승인된 추가 항목이 없습니다.")
        else:
            for ko, en in list(approved.items()):
                with st.container(border=True):
                    c1, c2, c3 = st.columns([3, 3, 1])
                    with c1:
                        st.markdown(f"`{ko}`")
                    with c2:
                        st.markdown(f"→ **{en}**")
                    with c3:
                        if st.button("🗑️", key=f"del_{ko}", help="삭제"):
                            del st.session_state.approved_glossary[ko]
                            st.rerun()

        # 관리자 직접 추가
        st.divider()
        st.subheader("➕ 직접 추가 (승인 없이 즉시 반영)")
        with st.form("admin_add_form", clear_on_submit=True):
            ca, cb = st.columns(2)
            with ca:
                admin_ko = st.text_input("한국어")
            with cb:
                admin_en = st.text_input("영문")
            if st.form_submit_button("바로 추가", type="primary"):
                if admin_ko.strip() and admin_en.strip():
                    st.session_state.approved_glossary[admin_ko.strip()] = admin_en.strip()
                    st.success(f"✅ `{admin_ko}` → **{admin_en}** 추가 완료!")
                    st.rerun()

        # 현황
        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric("기본 Glossary", len(BASE_GLOSSARY))
        m2.metric("승인된 추가",   len(approved))
        m3.metric("전체 합계",     len(BASE_GLOSSARY) + len(approved))
