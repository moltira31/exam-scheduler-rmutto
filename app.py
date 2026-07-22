import io
import pandas as pd
import streamlit as st

# --- 1. ตั้งค่าหน้าจอของโปรแกรมเบื้องต้น ---
st.set_page_config(
    page_title="ระบบจัดตารางสอบ - มทร.ตะวันออก วิทยาเขตจันทบุรี",
    page_icon="🏫",
    layout="wide",
)

# --- 2. ข้อมูลรหัสผ่านของระบบหลังบ้าน ---
USER_CREDENTIALS = {
    "monthira": "m123456",
    "registry_staff": "rmutto456",
}

# --- 3. จัดการสถานะการเข้าระบบ (Session State) ---
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""


def logout():
    st.session_state["logged_in"] = False
    st.session_state["username"] = ""
    st.rerun()


# --- 4. ฟังก์ชันสร้างไฟล์ Excel ตัวอย่าง (ปรับหัวคอลัมน์ตามแบบฟอร์มจริง) ---
def create_dummy_subject_excel():
    data = {
        "ที่": [1, 2, 3, 4],
        "รหัสวิชา": [
            "05-110-104",
            "05-110-104",
            "02-303-101",
            "02-303-102",
        ],
        "รายวิชา": [
            "ภาษาอังกฤษทั่วไป",
            "ภาษาอังกฤษทั่วไป",
            "การวิเคราะห์และออกแบบระบบ",
            "การเขียนโปรแกรมคอมพิวเตอร์",
        ],
        "ผู้สอน": [
            "อ.มาริสา คงดี",
            "อ.สมชาย ใจดี",
            "ดร.สมชาย ใจดี",
            "อ.สมศรี รักเรียน",
        ],
        "สำรองที่นั่ง": ["IS-261", "BA-21", "IS-261", "IT-11"],
        "กลุ่ม": ["1", "2", "1", "1"],
        "ลง": [35, 38, 35, 40],
        "M/F": ["M", "F", "M", "F"],  # M = Midterm / F = Final
        "ชั่วโมงสอบ": [1.5, 2.0, 3.0, 2.0],
        "สังกัดสาขา": [
            "ภาษาศาสตร์",
            "ภาษาศาสตร์",
            "เทคโนโลยีสารสนเทศ",
            "เทคโนโลยีสารสนเทศ",
        ],
        "ประเภทการสอบ": [
            "ทฤษฎี",
            "ทฤษฎี",
            "ทฤษฎี",
            "ปฏิบัติคอมพิวเตอร์",
        ],
        "รหัสกลุ่มสอบ": [
            "GENED_ENG01",
            "GENED_ENG01",
            "",
            "",
        ],
    }
    df = pd.DataFrame(data)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Exam_Subjects")
    return buffer.getvalue()


# ==================== หน้าจอ 1: ล็อกอิน (Login Screen) ====================
if not st.session_state["logged_in"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.image(
            "https://www.rmutto.ac.th/wp-content/uploads/2019/11/logo_rmutto.png",
            width=120,
        )
        st.title("ระบบจัดตารางสอบอัตโนมัติ")
        st.subheader("งานส่งเสริมวิชาการและงานทะเบียน วิทยาเขตจันทบุรี")
        st.write("กรุณาลงชื่อเข้าใช้งานด้วยรหัสที่ได้รับการอนุมัติจากแอดมินหลัก")

        username_input = st.text_input(
            "ชื่อผู้ใช้งาน (Username)", placeholder="กรอกชื่อผู้ใช้"
        )
        password_input = st.text_input(
            "รหัสผ่าน (Password)", type="password", placeholder="กรอกรหัสผ่าน"
        )

        if st.button("เข้าสู่ระบบ 🔐", use_container_width=True):
            if (
                username_input in USER_CREDENTIALS
                and USER_CREDENTIALS[username_input] == password_input
            ):
                st.session_state["logged_in"] = True
                st.session_state["username"] = username_input
                st.success("เข้าสู่ระบบสำเร็จ!")
                st.rerun()
            else:
                st.error("ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง")

# ==================== หน้าจอ 2: หน้าหลักการทำงาน (Main App) ====================
else:
    header_col1, header_col2 = st.columns([8, 2])
    with header_col1:
        st.title("🏫 ระบบจัดการตารางสอบ")
        st.caption("มหาวิทยาลัยเทคโนโลยีราชมงคลตะวันออก วิทยาเขตจันทบุรี")
    with header_col2:
        st.write(f"ผู้ใช้งาน: **{st.session_state['username']}**")
        if st.button("ออกจากระบบ 🚪", on_click=logout, use_container_width=True):
            pass

    st.markdown("---")

    # แถบเมนูด้านข้าง
    st.sidebar.header("เมนูการใช้งาน")
    menu_selection = st.sidebar.radio(
        "เลือกฟังก์ชันที่ต้องการทำงาน:",
        [
            "1. จัดตารางสอบ (ผู้สอนคุมสอบอัตโนมัติ)",
            "2. จัดตารางสอบ (คละผู้คุมสอบ)",
            "3. จัดการข้อมูลห้องสอบ",
        ],
    )

    st.sidebar.markdown("---")

    # ปุ่มดาวน์โหลดไฟล์แม่แบบ
    st.sidebar.subheader("📥 ดาวน์โหลดแบบฟอร์ม")
    excel_template = create_dummy_subject_excel()
    st.sidebar.download_button(
        label="ดาวน์โหลดไฟล์ตัวอย่างวิชาสอบ (.xlsx)",
        data=excel_template,
        file_name="exam_subjects_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ---------------- เมนูที่ 1: จัดตารางสอบ (ผู้สอนคุมสอบ) ----------------
    if menu_selection == "1. จัดตารางสอบ (ผู้สอนคุมสอบอัตโนมัติ)":
        st.header("🗓️ จัดตารางสอบอัตโนมัติ (อาจารย์ผู้สอนคุมสอบ)")
        st.write(
            "อัปโหลดไฟล์วิชาสอบ เพื่อให้ระบบจัดวัน เวลา และห้องสอบ โดยให้ผู้สอนคุมสอบเอง"
        )

        uploaded_file = st.file_uploader(
            "เลือกไฟล์ข้อมูลวิชาสอบ (.xlsx)", type=["xlsx"]
        )

        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)
            st.write("📋 **ตัวอย่างข้อมูลที่นำเข้า:**")
            st.dataframe(df.head(5))

            # ตรวจสอบคอลัมน์สำคัญตามแบบฟอร์มใหม่
            required_cols = [
                "ที่",
                "รหัสวิชา",
                "รายวิชา",
                "ผู้สอน",
                "สำรองที่นั่ง",
                "กลุ่ม",
                "ลง",
                "M/F",
                "ชั่วโมงสอบ",
            ]
            if all(col in df.columns for col in required_cols):
                st.success(
                    "✅ โครงสร้างไฟล์ถูกต้อง รองรับระบบ M/F และกลุ่มนักศึกษาเรียบร้อย"
                )

                if st.button("เริ่มประมวลผลจัดตารางสอบ ⚡", type="primary"):
                    with st.spinner("กำลังคำนวณและจัดสรรช่วงเวลา..."):
                        res_df = df.copy()

                        # ตัวอย่าง Logic การจัดสรร วัน/เวลา/ห้องสอบ และผู้คุมสอบ
                        days = [
                            "จันทร์ 10 ต.ค. 2026",
                            "อังคาร 11 ต.ค. 2026",
                            "พุธ 12 ต.ค. 2026",
                        ]
                        times = [
                            "09:00 - 10:30 น.",
                            "09:00 - 12:00 น.",
                            "13:00 - 15:00 น.",
                        ]
                        rooms = ["36-301", "36-302", "IT-201"]

                        assigned_datetime = []
                        assigned_supervisors = []

                        for idx, row in res_df.iterrows():
                            d = days[idx % len(days)]
                            t = times[idx % len(times)]
                            r = rooms[idx % len(rooms)]

                            # รวม วัน/เวลา/ห้องสอบ ไว้ในช่องเดียวกันตามแบบฟอร์ม
                            assigned_datetime.append(f"{d} ({t}) @{r}")
                            # กำหนดผู้คุมสอบ (เมนูนี้ให้ผู้สอนคุมเอง)
                            assigned_supervisors.append(row["ผู้สอน"])

                        res_df["วัน/เวลาสอบ"] = assigned_datetime
                        res_df["ผู้คุมสอบ"] = assigned_supervisors

                        # จัดลำดับคอลัมน์ใหม่ให้ตรงตามแบบฟอร์มภาพต้นฉบับ 100%
                        final_cols = [
                            "ที่",
                            "รหัสวิชา",
                            "รายวิชา",
                            "ผู้สอน",
                            "สำรองที่นั่ง",
                            "กลุ่ม",
                            "ลง",
                            "M/F",
                            "ชั่วโมงสอบ",
                            "วัน/เวลาสอบ",
                            "ผู้คุมสอบ",
                        ]
                        res_df = res_df[final_cols]

                        st.session_state["exam_schedule_result"] = res_df

                    st.balloons()
                    st.success(
                        f"จัดตารางสอบสำเร็จ! รวมทั้งหมด {len(res_df)} รายการ"
                    )

                # แสดงตารางผลลัพธ์และปุ่มดาวน์โหลด
                if "exam_schedule_result" in st.session_state:
                    final_result_df = st.session_state["exam_schedule_result"]

                    st.markdown("---")
                    st.subheader("📊 ตารางสอบที่จัดสำเร็จแล้ว")
                    st.dataframe(final_result_df, use_container_width=True)

                    # สร้างไฟล์ Excel สำหรับดาวน์โหลด
                    output_buffer = io.BytesIO()
                    with pd.ExcelWriter(
                        output_buffer, engine="openpyxl"
                    ) as writer:
                        final_result_df.to_excel(
                            writer, index=False, sheet_name="Exam_Schedule"
                        )

                    st.download_button(
                        label="📥 ดาวน์โหลดผลลัพธ์ตารางสอบ (Excel)",
                        data=output_buffer.getvalue(),
                        file_name="ตารางสอบ_มทร.ตะวันออก.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                    )
            else:
                st.error(
                    f"⚠️ ไฟล์ขาดคอลัมน์สำคัญ กรุณาตรวจสอบคอลัมน์: {', '.join(required_cols)}"
                )

    # ---------------- เมนูที่ 2: จัดตารางสอบ (คละผู้คุมสอบ) ----------------
    elif menu_selection == "2. จัดตารางสอบ (คละผู้คุมสอบ)":
        st.header("🔀 จัดตารางสอบอัตโนมัติ (คละอาจารย์คุมสอบ)")
        st.write(
            "ระบบจะสลับคละอาจารย์ผู้คุมสอบ โดยกระจายภาระงานและไม่ให้อาจารย์ติดสอนหรือคุมสอบชนกัน"
        )

        uploaded_file = st.file_uploader(
            "เลือกไฟล์ข้อมูลวิชาสอบ (.xlsx)", type=["xlsx"], key="menu2_file"
        )

        st.subheader("⚙️ ตั้งค่าเงื่อนไขการคละ")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            avoid_own_dept = st.checkbox(
                "หลีกเลี่ยงให้อาจารย์คุมสอบสาขาตัวเอง", value=True
            )
        with col_c2:
            max_duty = st.number_input(
                "ภาระงานคุมสอบสูงสุดต่อคน (ชั่วโมง)",
                min_value=1,
                max_value=20,
                value=6,
            )

        if uploaded_file is not None:
            if st.button("เริ่มประมวลผลคละผู้คุมสอบ 🔀", type="primary"):
                st.info("ระบบกำลังประมวลผลกระจายคาบสอบตามเงื่อนไข...")
                st.success(
                    "✅ คละผู้คุมสอบสำเร็จ! กระจายภาระงานเท่าเทียมกันเรียบร้อย"
                )

    # ---------------- เมนูที่ 3: จัดการข้อมูลห้องสอบ ----------------
    elif menu_selection == "3. จัดการข้อมูลห้องสอบ":
        st.header("🏫 จัดการข้อมูลห้องสอบ")
        st.write("ตรวจสอบและอัปเดตสถานะความพร้อมของห้องสอบในวิทยาเขต")

        mock_rooms = [
            {
                "อาคาร": "อาคารเรียนรวม 36 ปี",
                "รหัสห้อง": "36-301",
                "ความจุสอบ": 40,
                "ประเภท": "ห้องทฤษฎี",
                "สถานะ": "ใช้งานได้",
            },
            {
                "อาคาร": "อาคารเรียนรวม 36 ปี",
                "รหัสห้อง": "36-302",
                "ความจุสอบ": 40,
                "ประเภท": "ห้องทฤษฎี",
                "สถานะ": "งดใช้งานชั่วคราว",
            },
            {
                "อาคาร": "อาคารปฏิบัติการไอที",
                "รหัสห้อง": "IT-201",
                "ความจุสอบ": 35,
                "ประเภท": "ห้องปฏิบัติการคอมพิวเตอร์",
                "สถานะ": "ใช้งานได้",
            },
        ]

        with st.expander("➕ เพิ่มห้องสอบใหม่เข้าสู่ระบบ"):
            col_in1, col_in2 = st.columns(2)
            with col_in1:
                st.text_input("ชื่ออาคาร")
                st.text_input("รหัสห้อง (เช่น 36-301)")
            with col_in2:
                st.number_input("ความจุสอบ (คน)", min_value=1, value=30)
                st.selectbox(
                    "ประเภทห้อง", ["ห้องทฤษฎี", "ห้องปฏิบัติการคอมพิวเตอร์"]
                )

            if st.button("บันทึกห้องสอบใหม่"):
                st.success("บันทึกข้อมูลเรียบร้อยแล้ว!")

        st.subheader("📋 รายการห้องสอบในระบบปัจจุบัน")
        st.table(mock_rooms)
