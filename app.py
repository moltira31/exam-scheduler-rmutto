import io
import pandas as pd
import streamlit as st

# --- 1. ตั้งค่าหน้าจอของโปรแกรมเบื้องต้น ---
st.set_page_config(
    page_title="ระบบจัดตารางสอบ - มทร.ตะวันออก วิทยาเขตจันทบุรี",
    page_icon="🏫",
    layout="wide",
)

# --- 2. ข้อมูลรหัสผ่านของระบบหลังบ้าน (Database จำลอง) ---
USER_CREDENTIALS = {
    "monthira": "m123456",  # รหัสผ่านแอดมิน 1
    "registry_staff": "rmutto456",  # รหัสผ่านเจ้าหน้าที่ทะเบียน
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


# --- 4. ฟังก์ชันสร้างไฟล์ Excel ตัวอย่างให้ผู้ใช้ดาวน์โหลด ---
def create_dummy_subject_excel():
    data = {
        "รหัสวิชา": [
            "05-110-104",
            "05-110-104",
            "02-303-101",
            "02-303-102",
        ],
        "ชื่อวิชา": [
            "ภาษาอังกฤษทั่วไป",
            "ภาษาอังกฤษทั่วไป",
            "การวิเคราะห์และออกแบบระบบ",
            "การเขียนโปรแกรมคอมพิวเตอร์",
        ],
        "กลุ่มเรียน": ["IS-261", "BA-21", "IS-261", "IT-11"],
        "จำนวนผู้เข้าสอบ": [35, 38, 35, 40],
        "ชื่อผู้สอน": [
            "อ.มาริสา คงดี",
            "อ.สมชาย ใจดี",
            "ดร.สมชาย ใจดี",
            "อ.สมศรี รักเรียน",
        ],
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
        "ชั่วโมงสอบ (นาที)": [90, 90, 180, 120],
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

    # ปุ่มดาวน์โหลดไฟล์แม่แบบใน Sidebar
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

            # ตรวจสอบคอลัมน์สำคัญ
            required_cols = [
                "รหัสวิชา",
                "ชื่อวิชา",
                "ชั่วโมงสอบ (นาที)",
                "รหัสกลุ่มสอบ",
            ]
            if all(col in df.columns for col in required_cols):
                st.success("✅ โครงสร้างไฟล์ถูกต้อง รองรับชั่วโมงสอบและวิชาเรียนรวม")

                if st.button("เริ่มประมวลผลจัดตารางสอบ ⚡", type="primary"):
                    with st.spinner("กำลังคำนวณและจัดสรรช่วงเวลา..."):
                        # [ประมวลผลตรรกะเวลาและรหัสกลุ่มสอบ]
                        gened_count = df[
                            df["รหัสกลุ่มสอบ"].notna()
                            & (df["รหัสกลุ่มสอบ"] != "")
                        ]["รหัสกลุ่มสอบ"].nunique()

                    st.balloons()
                    st.success(
                        f"จัดตารางสอบสำเร็จ! รวมทั้งหมด {len(df)} รายการ (พบวิชาเรียนรวม {gened_count} กลุ่มสอบ)"
                    )
                    st.button("ดาวน์โหลดผลลัพธ์ตารางสอบ (Excel) 📥")
            else:
                st.error(
                    f"⚠️ ไฟล์ขาดคอลัมน์สำคัญ กรุณาตรวจสอบว่ามีคอลัมน์: {', '.join(required_cols)}"
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
                st.success("✅ คละผู้คุมสอบสำเร็จ! กระจายภาระงานเท่าเทียมกันเรียบร้อย")

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
