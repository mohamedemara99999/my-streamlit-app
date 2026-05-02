import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from io import BytesIO

# ================== تسجيل الدخول ==================
USERS = {
    "admin": "m7md3mara2025",
    "user3": "2468"
}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "current_user" not in st.session_state:
    st.session_state.current_user = None

if not st.session_state.logged_in:
    st.title("🔐 تسجيل الدخول")
    username = st.text_input("اسم المستخدم")
    password = st.text_input("كلمة المرور", type="password")

    if st.button("دخول"):
        if username in USERS and USERS[username] == password:
            st.session_state.logged_in = True
            st.session_state.current_user = username
            st.rerun()
        else:
            st.error("❌ بيانات غير صحيحة")

# ================== بعد الدخول ==================
if st.session_state.logged_in:

    st.sidebar.success(f"مرحباً {st.session_state.current_user}")

    if st.sidebar.button("تسجيل خروج"):
        st.session_state.logged_in = False
        st.session_state.current_user = None
        st.rerun()

    st.title("Excel Analyzer Tool")

    uploaded_file = st.file_uploader("رفع ملف Excel", type=["xlsx", "xls"])
    current_df = None
    original_df = None

    if uploaded_file:
        current_df = pd.read_excel(uploaded_file, engine="openpyxl")
        original_df = current_df.copy()

        current_df.columns = current_df.columns.str.strip()
        current_df = current_df.dropna(how="all", axis=1)

        st.success("تم تحميل الملف")
        st.dataframe(current_df)

# ================== Excel Format ==================
def format_excel(output):
    output.seek(0)
    wb = load_workbook(output)

    header_fill = PatternFill("solid", fgColor="006400")
    header_font = Font(bold=True, color="FFFFFF")

    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

    final = BytesIO()
    wb.save(final)
    final.seek(0)
    return final

# ================== ETISALAT ==================
def generate_etisalat_report(df, original_df):
    df = df.copy()

    df['Originating_Number'] = df['Originating_Number'].fillna('').astype(str)
    df['Terminating_Number'] = df['Terminating_Number'].fillna('').astype(str)

    numbers = pd.concat([df['Originating_Number'], df['Terminating_Number']])
    numbers = numbers[numbers != '']

    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number', 'Count']

    df_final = freq.copy()

    df_final['SMS'] = 0

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='calls', index=False)
        original_df.to_excel(writer, sheet_name='cheet', index=False)

    return format_excel(output)

# ================== VODAFONE ==================
def generate_vodafone_report(df):
    df = df.copy()

    df['B_NUMBER'] = df['B_NUMBER'].astype(str)

    freq = df['B_NUMBER'].value_counts().reset_index()
    freq.columns = ['B Number', 'Count']

    freq['SMS'] = 0

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        freq.to_excel(writer, sheet_name='calls', index=False)
        df.to_excel(writer, sheet_name='cheet', index=False)

    return format_excel(output)

# ================== ORANGE ==================
def generate_orange_report(df):
    df = df.copy()

    df.columns = df.columns.str.upper()

    if 'OTHER_MSISDN' not in df.columns:
        st.error("ملف أورانج غير صحيح")
        return None

    freq = df['OTHER_MSISDN'].value_counts().reset_index()
    freq.columns = ['B Number', 'Count']

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        freq.to_excel(writer, sheet_name='calls', index=False)
        df.to_excel(writer, sheet_name='cheet', index=False)

    return format_excel(output)

# ================== ETISALAT COMPANY ==================
def generate_etisalat_company_report(df):
    df = df.copy()

    freq = df['b_number_full'].value_counts().reset_index()
    freq.columns = ['B Number', 'Count']

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        freq.to_excel(writer, sheet_name='calls', index=False)
        df.to_excel(writer, sheet_name='cheet', index=False)

    return format_excel(output)

# ================== BUTTONS ==================
if current_df is not None:

    st.subheader("التقارير")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        if st.button("اتصالات"):
            output = generate_etisalat_report(current_df, current_df.copy())
            st.download_button("تحميل", output, "etisalat.xlsx")

    with col2:
        if st.button("فودافون"):
            output = generate_vodafone_report(current_df)
            st.download_button("تحميل", output, "vodafone.xlsx")

    with col3:
        if st.button("اورانج"):
            output = generate_orange_report(current_df)
            if output:
                st.download_button("تحميل", output, "orange.xlsx")

    with col4:
        if st.button("اتصالات شركه"):
            output = generate_etisalat_company_report(current_df)
            st.download_button("تحميل", output, "etisalat_company.xlsx")
