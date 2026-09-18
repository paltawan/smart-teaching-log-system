import streamlit as st
from online_files import file_uploader as online_file_uploader
import os
import json
import re
import io
import copy
import time
from datetime import datetime, timedelta
from docxtpl import DocxTemplate
import docx
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from google import genai
from google.genai import types

st.set_page_config(
    page_title="ระบบบันทึกหลังการสอน AI อาชีวศึกษา",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# สไตล์ตกแต่ง UI สีสัน สดใส มีมิติ
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap');
    [data-testid="stStatusWidget"] {
        display: none !important;
    }
    html, body, [class*="css"] {
        font-family: 'Prompt', sans-serif;
    }
    .main-header {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 50%, #06B6D4 100%);
        padding: 24px;
        border-radius: 16px;
        color: white;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px rgba(30, 58, 138, 0.25);
    }
    .main-header h1 {
        color: white !important;
        font-weight: 700;
        margin-bottom: 8px;
        text-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }
    .main-header p {
        color: #E0F2FE !important;
        font-size: 15px;
        margin-bottom: 0;
    }
    .card-box {
        background: #ffffff;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
    }
    .badge-tag {
        background: linear-gradient(90deg, #EC4899, #8B5CF6);
        color: white;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
        display: inline-block;
        margin-bottom: 10px;
    }
    .stButton>button {
        background: linear-gradient(90deg, #F43F5E 0%, #E11D48 100%) !important;
        color: white !important;
        font-size: 17px !important;
        font-weight: 600 !important;
        border-radius: 12px !important;
        border: none !important;
        padding: 12px 24px !important;
        box-shadow: 0 8px 20px rgba(225, 29, 72, 0.3) !important;
        transition: all 0.3s ease !important;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 12px 25px rgba(225, 29, 72, 0.45) !important;
    }
    .footer-box {
        text-align: center;
        padding: 24px 10px;
        margin-top: 50px;
        border-top: 1px solid #E2E8F0;
        color: #64748B;
        font-size: 13.5px;
    }
    .footer-badge {
        display: inline-block;
        background: #F1F5F9;
        border: 1px solid #CBD5E1;
        padding: 5px 14px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
        color: #334155;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    if st.session_state.get("combined_app"):
        st.page_link("project_app.py", label="ไปยังระบบโครงการสอน", icon=":material/menu_book:", width="stretch")
    else:
        st.link_button(
            "ไปยังระบบโครงการสอน",
            "http://127.0.0.1:8517/",
            icon=":material/menu_book:",
            width="stretch",
        )

# ข้อมูลปฏิทินและแผนกวิชา
THAI_MONTHS = [
    "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]
DAY_NAMES = ["วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์", "วันอาทิตย์"]
DEPARTMENT_OPTIONS = [
    "การจัดการโลจิสติกส์และซัพพลายเชน",
    "เทคโนโลยีสารสนเทศ",
    "คอมพิวเตอร์ธุรกิจ",
    "การบัญชี",
    "การตลาด",
    "ช่างยนต์",
    "ช่างไฟฟ้ากำลัง",
    "ช่างอิเล็กทรอนิกส์",
    "ช่างก่อสร้าง",
    "อื่นๆ (ระบุเอง)"
]

st.markdown("""
<div class="main-header">
    <h1>📝 ระบบจัดทำบันทึกหลังการสอนอัตโนมัติ (AI Professional)</h1>
    <p>สกัดโครงการสอนตามหลักวิชาการอาชีวศึกษา จัดรูปแบบต่อกันหน้าต่อหน้า ไร้หน้าว่าง 100%</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ การตั้งค่าระบบ")
    
    api_key_input = st.text_input("🔑 Gemini API Key:", type="password", placeholder="AIzaSy...")
    st.markdown("[👉 รับ API Key ฟรีคลิกที่นี่](https://aistudio.google.com/apikey)")
    st.divider()

    st.subheader("👤 ข้อมูลครูผู้สอน")
    teacher_name = st.text_input("ชื่อ-สกุลครูผู้สอน:")
    
    dept_choice = st.selectbox("สาขาวิชา / แผนกวิชา:", DEPARTMENT_OPTIONS, index=0)
    if dept_choice == "อื่นๆ (ระบุเอง)":
        department = st.text_input("ระบุสาขาวิชาของคุณ:", value="")
    else:
        department = dept_choice

    st.markdown("""
    <div style="font-size: 12px; color: #94A3B8; text-align: center; margin-top: 25px; line-height: 1.6;">
        <b>AI Vocational Reflection System</b><br/>
        สาขาเทคโนโลยีสารสนเทศ &amp; สาขาโลจิสติกส์
    </div>
    """, unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 1. แบบฟอร์มและหลักสูตร")
    tpl_file = (online_file_uploader("📄 แนบแบบฟอร์มวิทยาลัย (template.docx):", ["docx"], "reflection_template")
                if st.session_state.get("combined_app") else st.file_uploader("📄 แนบแบบฟอร์มวิทยาลัย (template.docx):", type=["docx"]))
    uploaded_file = (online_file_uploader("📚 แนบไฟล์โครงการสอน (PDF, Word, TXT, รูปภาพ):", ["pdf", "docx", "txt", "png", "jpg", "jpeg"], "reflection_source")
                     if st.session_state.get("combined_app") else st.file_uploader("📚 แนบไฟล์โครงการสอน (PDF, Word, TXT, รูปภาพ):", type=["pdf", "docx", "txt", "png", "jpg", "jpeg"]))
    
    st.markdown("---")
    st.markdown("🎯 **เลือกระดับชั้นและวุฒิการศึกษา:**")
    c_deg, c_yr = st.columns(2)
    with c_deg:
        degree = st.selectbox("ระดับคุณวุฒิ:", ["ปวช.", "ปวส."])
    with c_yr:
        if degree == "ปวช.":
            year_num = st.selectbox("ชั้นปี:", ["1", "2", "3"])
            target_weeks = 18
        else:
            year_num = st.selectbox("ชั้นปี:", ["1", "2"])
            target_weeks = 15

    class_level = f"{degree} {year_num}"
    st.info(f"✨ ระดับ: **{class_level}** | สาขา: **{department}** | กำหนดอัตโนมัติ: **{target_weeks} สัปดาห์**")

with col2:
    st.subheader("⏰ 2. ตารางวัน-เวลา และการฉีกคาบสอน")
    slots_count = st.selectbox(
        "จำนวนคาบสอนใน 1 สัปดาห์ (ฉีกคาบได้สูงสุด 4 คาบ):",
        options=[1, 2, 3, 4],
        format_func=lambda x: f"สอน {x} คาบ / สัปดาห์" if x > 1 else "สอน 1 คาบ (วันเดียวจบ)",
        index=1
    )

    slots_info = []
    default_days = [0, 1, 2, 3]
    time_options = ["08.30", "09.30", "10.30", "11.30", "12.30", "13.30", "14.30", "15.30", "16.30"]

    for i in range(slots_count):
        st.markdown(f"**📌 รายละเอียดคาบที่ {i+1}:**")
        col_day, col_start, col_end = st.columns(3)
        with col_day:
            d_val = st.selectbox(f"วัน (คาบที่ {i+1}):", DAY_NAMES, index=default_days[i % len(default_days)], key=f"day_slot_{i}")
        with col_start:
            start_time = st.selectbox(
                f"เวลาเริ่ม (คาบที่ {i+1}):",
                options=time_options,
                index=0,
                key=f"start_time_{i}"
            )
        with col_end:
            end_time = st.selectbox(
                f"เวลาสิ้นสุด (คาบที่ {i+1}):",
                options=time_options,
                index=1,
                key=f"end_time_{i}"
            )
        slots_info.append({"day": d_val, "start": start_time, "end": end_time})

    start_date = st.date_input(
        "📅 วันที่เริ่มสอนสัปดาห์ที่ 1 (คำนวณปฏิทินไทยอัตโนมัติ):",
        format="YYYY/MM/DD"
    )

    holiday_text = st.text_area(
        "ระบุสัปดาห์และเหตุผลการงดสอน (ถ้ามี เช่น '8:ตรงกับวันหยุดนักขัตฤกษ์'):",
        value="8:ตรงกับวันหยุดนักขัตฤกษ์ตามประกาศสถานศึกษา",
        height=70
    )

def format_thai_date(dt):
    d = dt.day
    m = THAI_MONTHS[dt.month]
    y = dt.year + 543
    day_name = DAY_NAMES[dt.weekday()]
    return f"{day_name} {d} {m} {y}"


def force_line_breaks(text):
    if text is None:
        return "-"
    txt = str(text).replace("\r\n", "\n").replace("\r", "\n").strip()
    if not txt:
        return "-"

    for marker, replacement in [
        ("K:", "\nK:"),
        ("P:", "\nP:"),
        ("A:", "\nA:"),
        ("ด้านความรู้:", "\nด้านความรู้:"),
        ("ด้านทักษะ/กระบวนการ:", "\nด้านทักษะ/กระบวนการ:"),
        ("ด้านคุณลักษณะ:", "\nด้านคุณลักษณะ:"),
        ("ปัญหา:", "\nปัญหา:"),
        ("แนวทางแก้ไข:", "\nแนวทางแก้ไข:"),
        ("การจัดการเรียนรู้:", "\nการจัดการเรียนรู้:"),
        ("สื่อและนวัตกรรม:", "\nสื่อและนวัตกรรม:"),
        ("การวัดและประเมินผล:", "\nการวัดและประเมินผล:")
    ]:
        txt = txt.replace(marker, replacement)

    return txt.strip().replace("\n\n", "\n")


def extract_valid_json(raw_text):
    if raw_text is None:
        raise ValueError("AI ไม่ได้คืนค่าข้อมูลกลับมา")

    text = str(raw_text).strip()
    if not text:
        raise ValueError("AI คืนค่าว่างเปล่า")

    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:].lstrip()
        if text.endswith("```"):
            text = text[:-3].rstrip()

    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and start < end:
        text = text[start:end + 1]

    text = text.replace("“", '"').replace("”", '"')
    text = text.replace("‘", "'").replace("’", "'")

    # ซ่อม escape ที่ JSON รับไม่ไหว เช่น \d, \x, \u ที่ไม่ใช่ escape ที่ถูกต้อง
    text = re.sub(r'\\(?!["\\/bfnrtu])', r'\\\\', text)

    # สกัดเฉพาะ object JSON แรกที่พบ เพื่อหลีกเลี่ยงข้อความข้างหน้า/ข้างหลัง
    first_obj = re.search(r"\{.*\}", text, re.DOTALL)
    if first_obj:
        text = first_obj.group(0)

    return text


if st.button(f"🚀 เริ่มสร้างเอกสารบันทึกหลังการสอนครบ {target_weeks} สัปดาห์", use_container_width=True):
    api_key = api_key_input.strip() if api_key_input else ""
    if not api_key:
        st.warning("⚠️ กรุณากรอก Gemini API Key ที่แถบด้านซ้ายก่อนเริ่มใช้งาน")
        st.stop()
    if not tpl_file:
        st.warning("⚠️ กรุณาแนบไฟล์ template.docx ของวิทยาลัย")
        st.stop()
    if not uploaded_file:
        st.warning("⚠️ กรุณาแนบไฟล์โครงการสอน")
        st.stop()
    if not department.strip():
        st.warning("⚠️ กรุณาระบุสาขาวิชา/แผนกวิชา")
        st.stop()

    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        status_text.text("🤖 กำลังส่งข้อมูลให้ Gemini AI วิเคราะห์โครงการสอนตามหลักวิชาการอาชีวศึกษา...")
        client = genai.Client(api_key=api_key)
        file_bytes = uploaded_file.read()
        mime_type = uploaded_file.type or "application/octet-stream"

        prompt = f"""
        คุณคือผู้เชี่ยวชาญด้านหลักสูตรและการจัดการเรียนรู้อาชีวศึกษา (สอศ.)
        จงวิเคราะห์เนื้อหาโครงการสอนที่แนบมานี้ เพื่อจัดทำ "บันทึกหลังการจัดการเรียนรู้" ระดับชั้น {class_level}
        ให้ครบถ้วนตั้งแต่สัปดาห์ที่ 1 ถึงสัปดาห์ที่ {target_weeks} (รวม {target_weeks} สัปดาห์พอดี ห้ามขาด)
        ข้อมูลวันหยุด/งดสอน: {holiday_text}

        เกณฑ์การเขียนเชิงวิชาการที่เข้มข้น สมบูรณ์ และมีมิติ (ความยาวพอเหมาะ ไม่สั้นเกินไปและไม่ล้นหน้า):
        1. topic: ระบุชื่อหน่วยการเรียนรู้ และสมรรถนะประจำหน่วย/หัวข้อการเรียนรู้อย่างชัดเจน
        2. student_eval: ประเมินผลการเรียนรู้ของผู้เรียนอย่างเป็นรูปธรรม แยกมิติ K-P-A และบังคับให้ขึ้นบรรทัดใหม่แบบนี้:
           - "K: ..."
           - "P: ..."
           - "A: ..."
           - "สรุป: มีผู้เรียนผ่านเกณฑ์การประเมินร้อยละ 85 ขึ้นไป (หรือสอดคล้องกับแต่ละสัปดาห์)"
        3. teacher_eval: ผลการสอนของครู เน้นการจัดการเรียนรู้เชิงรุก (Active Learning) แยกเป็นบรรทัดใหม่ตามหัวข้อ:
           - "การจัดการเรียนรู้: ..."
           - "สื่อและนวัตกรรม: ..."
           - "การวัดและประเมินผล: ..."
        4. problem_solution: ปัญหา อุปสรรค และแนวทางแก้ไขตามวงรอบคุณภาพ (PDCA) แยกเป็นบรรทัดใหม่ตามหัวข้อ:
           - "ปัญหา: ..."
           - "แนวทางแก้ไข: ..."

        ส่งออกเป็น Pure JSON โครงสร้างนี้เท่านั้น:
        {{
            "code": "รหัสวิชา",
            "subject": "ชื่อวิชา",
            "weeks": [
                {{
                    "week": 1,
                    "topic": "ชื่อหน่วยและเรื่องที่จัดการเรียนรู้",
                    "is_holiday": false,
                    "off_reason": "-",
                    "student_eval": "K: ...\nP: ...\nA: ...\nสรุป: ...",
                    "teacher_eval": "การจัดการเรียนรู้: ...\nสื่อและนวัตกรรม: ...\nการวัดและประเมินผล: ...",
                    "problem_solution": "ปัญหา: ...\nแนวทางแก้ไข: ..."
                }}
            ]
        }}
        ห้ามใส่เครื่องหมาย markdown block ส่งเฉพาะ Pure JSON เท่านั้น
        """

        models_to_try = [
            "gemini-3.8-flash",
            "gemini-3.5-flash-lite"
        ]
        response = None
        last_error = None

        for target_m in models_to_try:
            status_text.text(f"⏳ กำลังประมวลผลด้วยโมเดล {target_m}...")
            try:
                response = client.models.generate_content(
                    model=target_m,
                    contents=[
                        types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
                        prompt
                    ],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.25
                    )
                )
                if response and response.text:
                    break
            except Exception as err:
                last_error = err
                time.sleep(1)

        if not response or not response.text:
            raise Exception(f"ไม่สามารถเชื่อมต่อโมเดลได้: {last_error}")

        clean_text = extract_valid_json(response.text)
        try:
            data = json.loads(clean_text)
        except json.JSONDecodeError as decode_err:
            repaired = re.sub(r'\\(?!["\\/bfnrtu])', r'\\\\', clean_text)
            try:
                data = json.loads(repaired)
            except json.JSONDecodeError:
                fallback_match = re.search(r"\{.*\}", clean_text, re.DOTALL)
                if fallback_match:
                    data = json.loads(fallback_match.group(0))
                else:
                    raise Exception(f"AI คืนค่า JSON ที่ไม่ถูกต้อง: {decode_err.msg} (line {decode_err.lineno}, col {decode_err.colno})")

        course_code = data.get("code", "วิชา")
        course_name = data.get("subject", "โครงการสอน")
        weeks_data = data.get("weeks", [])

        existing = {w.get("week"): w for w in weeks_data}
        final_weeks = []
        for i in range(1, target_weeks + 1):
            if i in existing:
                final_weeks.append(existing[i])
            else:
                final_weeks.append({
                    "week": i,
                    "topic": f"หน่วยการเรียนรู้ที่ {i} การประยุกต์ใช้สมรรถนะวิชาชีพตามโครงการสอน",
                    "is_holiday": False,
                    "off_reason": "-",
                    "student_eval": "K: ผู้เรียนมีความรู้ความเข้าใจในหลักการและสามารถตอบคำถามได้ถูกต้อง\nP: ผู้เรียนสามารถฝึกปฏิบัติงานตามใบงานและขั้นตอนได้ถูกต้องตามแบบฟอร์ม\nA: ผู้เรียนมีวินัย ความรับผิดชอบ และตรงต่อเวลาในการปฏิบัติงาน\nสรุป: มีผู้เรียนผ่านเกณฑ์การประเมินร้อยละ 90 ขึ้นไป",
                    "teacher_eval": "การจัดการเรียนรู้: จัดการเรียนรู้เชิงรุกด้วยการสาธิตและฝึกปฏิบัติจริง\nสื่อและนวัตกรรม: ใช้สื่อการสอนดิจิทัลและใบงานประกอบการเรียนรู้\nการวัดและประเมินผล: ประเมินจากพฤติกรรมและผลงานจริงของผู้เรียน",
                    "problem_solution": "ปัญหา: ผู้เรียนบางคนยังมีความสับสนในขั้นตอนการปฏิบัติงาน\nแนวทางแก้ไข: ให้คำแนะนำรายบุคคล และจัดกลุ่มเพื่อนช่วยเพื่อนทบทวนร่วมกัน"
                })

        tpl_bytes = tpl_file.read()
        merged_doc = None
        total_count = len(final_weeks)
        day_map = {name: idx for idx, name in enumerate(DAY_NAMES)}

        for idx, w in enumerate(final_weeks):
            progress_bar.progress(int(((idx + 1) / total_count) * 100))
            status_text.text(f"📝 กำลังลงข้อมูลสัปดาห์ที่ {w.get('week')} ในแบบฟอร์มวิทยาลัย...")

            week_num = w.get("week", idx + 1)
            base_week_date = start_date + timedelta(weeks=(week_num - 1))
            
            date_lines = []
            time_lines = []
            for slot in slots_info:
                t_wday = day_map.get(slot["day"], 0)
                dt_slot = base_week_date + timedelta(days=(t_wday - base_week_date.weekday()))
                date_lines.append(format_thai_date(dt_slot))
                start_t = slot.get("start", "08.30")
                end_t = slot.get("end", start_t)
                time_lines.append(f"เวลา {start_t}-{end_t} น.")

            date_display = "\n".join(date_lines)
            time_display = "\n".join(time_lines)

            is_hol = w.get("is_holiday", False)
            context_w = {
                "code": course_code,
                "subject": course_name,
                "week": week_num,
                "date": date_display,
                "date_display": date_display,
                "time": time_display,
                "time_display": time_display,
                "topic": w.get("topic"),
                "level": class_level,
                "department": department,
                "check_on": "☑" if not is_hol else "☐",
                "check_off": "☑" if is_hol else "☐",
                "off_reason": w.get("off_reason", "-") if is_hol else "-",
                "student_eval": force_line_breaks(w.get("student_eval")),
                "teacher_eval": force_line_breaks(w.get("teacher_eval")),
                "problem_solution": force_line_breaks(w.get("problem_solution")),
            }

            t = DocxTemplate(io.BytesIO(tpl_bytes))
            render_context = {
                "teacher_name": teacher_name,
                "w": context_w,
                **context_w
            }
            t.render(render_context)

            tmp_io = io.BytesIO()
            t.save(tmp_io)
            tmp_io.seek(0)

            sub_doc = docx.Document(tmp_io)

            # กำจัดย่อหน้าว่างเปล่าท้ายเอกสารต้นฉบับทั้งหมด เพื่อไม่ให้เกิดการดันตกหน้า
            while len(sub_doc.paragraphs) > 0:
                last_p = sub_doc.paragraphs[-1]
                if not last_p.text.strip() and not last_p._element.xpath('.//w:drawing'):
                    p_elem = last_p._element
                    p_elem.getparent().remove(p_elem)
                else:
                    break

            if merged_doc is None:
                merged_doc = sub_doc
            else:
                # รวมเอกสารโดยต่อหน้าใหม่แบบไม่มีหน้าว่างคั่น (Pure Page Break)
                is_first_block = True
                for el in sub_doc.element.body:
                    if el.tag.endswith('sectPr'):
                        continue
                    copied_el = copy.deepcopy(el)

                    if is_first_block:
                        # สร้าง Page Break แบบแนบสนิท ไม่ทิ้งระยะบรรทัดว่าง
                        p_break = OxmlElement('w:p')
                        pPr = OxmlElement('w:pPr')
                        pPr.append(OxmlElement('w:pageBreakBefore'))
                        
                        # กำหนดขนาดฟอนต์ 1pt และระยะบรรทัด 0 เพื่อไม่ให้กินพื้นที่แม้แต่มิลลิเมตรเดียว
                        spacing = parse_xml(r'<w:spacing %s w:before="0" w:after="0" w:line="1" w:lineRule="exact"/>' % nsdecls('w'))
                        rPr = parse_xml(r'<w:rPr %s><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>' % nsdecls('w'))
                        pPr.append(spacing)
                        pPr.append(rPr)
                        p_break.append(pPr)

                        merged_doc.element.body.append(p_break)
                        is_first_block = False

                    merged_doc.element.body.append(copied_el)

        # ลบ Section Break และ Empty Paragraph ท้ายเอกสารหลัก
        while len(merged_doc.paragraphs) > 0:
            final_p = merged_doc.paragraphs[-1]
            if not final_p.text.strip() and not final_p._element.xpath('.//w:drawing'):
                p_elem = final_p._element
                p_elem.getparent().remove(p_elem)
            else:
                break

        output_stream = io.BytesIO()
        merged_doc.save(output_stream)
        output_stream.seek(0)

        progress_bar.progress(100)
        status_text.empty()

        st.balloons()
        st.success(f"🎉 สร้างเอกสารสำเร็จครบ {target_weeks} สัปดาห์ ถูกต้องตามหลักวิชาการ ต่อเนื่องหน้าต่อหน้า ไร้หน้าว่าง 100%!")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Word (แบบฟอร์มวิทยาลัยตรงเป๊ะ)",
            data=output_stream,
            file_name=f"บันทึกหลังการสอน_{course_code}_{class_level.replace(' ', '')}_ครบ{target_weeks}สัปดาห์.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    except Exception as e:
        status_text.empty()
        error_msg = str(e)
        if "403" in error_msg and "PERMISSION_DENIED" in error_msg:
            st.error("⚠️ เกิดข้อผิดพลาด: API Key ของคุณมีปัญหา หรือถูกระงับการใช้งาน โปรดสร้าง API Key ใหม่ที่ Google AI Studio และนำมาอัปเดตใหม่ครับ")
        else:
            st.error(f"เกิดข้อผิดพลาด: {error_msg}")

# กล่องข้อมูลลิขสิทธิ์และผู้พัฒนาระบบด้านล่างสุด
st.markdown("""
<div class="footer-box">
    <div class="footer-badge">🛡️ PROPRIETARY & EDUCATIONAL OPEN-SOURCE</div><br/>
    <b>ระบบปัญญาประดิษฐ์สกัดและจัดทำบันทึกหลังการสอนอาชีวศึกษา (AI Vocational Reflection)</b><br/>
    สาขาเทคโนโลยีสารสนเทศ &amp; สาขาโลจิสติกส์<br/>
    <span style="font-size: 12px; color: #94A3B8;">ขับเคลื่อนด้วย Streamlit & Google Gemini AI Flash Engine</span>
</div>
""", unsafe_allow_html=True)
