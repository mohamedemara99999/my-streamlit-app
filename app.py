import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from io import BytesIO

# ================== تسجيل الدخول ==================
USERS = {
    "admin": "m7md3mara2025",
    "user1": "mostafatalaat",
    "user2": "mohamedelmasry",
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

# ================== بعد تسجيل الدخول ==================
if st.session_state.logged_in:

    st.sidebar.success(f"مرحباً {st.session_state.current_user}")
    if st.sidebar.button("تسجيل خروج"):
        st.session_state.logged_in = False
        st.session_state.current_user = None
        st.rerun()

    st.title("Excel Analyzer Tool")

    selected_company = st.selectbox(
        "اختر الشركة",
        ["etisalat", "etisalat_company", "vodafone", "orange"]
    )

    uploaded_file = st.file_uploader("ارفع ملف Excel", type=["xlsx"])

    current_df = None
    original_df = None

    if uploaded_file:
        current_df = pd.read_excel(uploaded_file, engine="openpyxl")
        original_df = current_df.copy()

        current_df.columns = current_df.columns.str.strip()
        st.success("تم تحميل الملف")
        st.dataframe(current_df)

# ================== تنسيق Excel ==================
def format_excel_sheets(output, header_color="006400", company="etisalat"):
    output.seek(0)
    wb = load_workbook(output)

    header_fill = PatternFill("solid", fgColor=header_color)
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

# ================== اتصالات ==================
def generate_etisalat_report(df, original_df):
    df = df.copy()

    df['Originating_Number'] = df['Originating_Number'].fillna('').astype(str)
    df['Terminating_Number'] = df['Terminating_Number'].fillna('').astype(str)

    numbers = pd.concat([df['Originating_Number'], df['Terminating_Number']])
    numbers = numbers[numbers != '']
    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number', 'Count']

    b_data = {}
    for _, row in df.iterrows():
        for col in ['Originating_Number', 'Terminating_Number']:
            num = str(row.get(col, ''))
            if num:
                b_data[num] = {
                    'B Full Name': row.get('B_Number_Full_Name', ''),
                    'B Address': row.get('B_Number_Address', ''),
                    'Latitude': row.get('B_Number_MU_Latitude', ''),
                    'Longitude': row.get('B_Number_MU_Longitude', '')
                }

    df_final = freq.copy()

    for col in ['B Full Name','B Address','Latitude','Longitude']:
        df_final[col] = df_final['B Number'].map(lambda x: b_data.get(x, {}).get(col, ''))

    df_final['SMS'] = 0

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='calls', index=False)
        original_df.to_excel(writer, sheet_name='cheet', index=False)

    output.seek(0)
    return format_excel_sheets(output, company="etisalat")

# ================== فودافون (مُصحح بالكامل) ==================
def generate_vodafone_report(df):
    df = df.copy()

    df['B_NUMBER'] = df['B_NUMBER'].fillna('').astype(str)
    df['IMEI'] = df['IMEI'].fillna('').astype(str)
    df['FULL_DATE'] = pd.to_datetime(df['FULL_DATE'], errors='coerce')

    freq = df['B_NUMBER'].value_counts().reset_index()
    freq.columns = ['B Number', 'Count']

    sms = df[df['SERVICE'].astype(str).str.contains("Short message", na=False)]
    sms_count = sms.groupby('B_NUMBER').size().reset_index(name='SMS')

    base = df[[
        'B_NUMBER',
        'B_NUMBER_FIRST_NAME',
        'B_NUMBER_LAST_NAME',
        'B_NUMBER_ADDRESS',
        'B_NUMBER_SITE_ADDRESS',
        'B_NUMBER_NATIONAL_ID'
    ]].drop_duplicates()

    base['B Full Name'] = base['B_NUMBER_FIRST_NAME'].fillna('') + ' ' + base['B_NUMBER_LAST_NAME'].fillna('')

    df_final = freq.merge(base, left_on='B Number', right_on='B_NUMBER', how='left')
    df_final = df_final.merge(sms, left_on='B Number', right_on='B_NUMBER', how='left')

    df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)
    df_final['Count'] = df_final['Count'].astype(int)

    call_dates = df.groupby('B_NUMBER')['FULL_DATE'].agg(
        First_Call='min',
        Last_Call='max'
    ).reset_index()

    df_final = df_final.merge(call_dates, left_on='B Number', right_on='B_NUMBER', how='left')

    df_final = df_final[[
        'B Number','Count','B Full Name',
        'B_NUMBER_NATIONAL_ID','B_NUMBER_ADDRESS',
        'B_NUMBER_SITE_ADDRESS','SMS',
        'First_Call','Last_Call'
    ]]

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name="calls", index=False)
        df.to_excel(writer, sheet_name="cheet", index=False)

    output.seek(0)
    return format_excel_sheets(output, header_color="FF0000", company="vodafone")

# ================== أورانج ==================
def generate_orange_report(df):
    df = df.copy()
    df.columns = df.columns.str.upper()

    freq = df['OTHER_MSISDN'].value_counts().reset_index()
    freq.columns = ['B Number','Count']

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        freq.to_excel(writer, sheet_name="calls", index=False)
        df.to_excel(writer, sheet_name="cheet", index=False)

    output.seek(0)
    return format_excel_sheets(output, header_color="FF6600", company="orange")

# ================== أزرار التحليل ==================
if current_df is not None:

    st.subheader("تحليل البيانات")

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("اتصالات"):
            out = generate_etisalat_report(current_df, original_df)
            st.download_button("تحميل", out, "etisalat.xlsx")

    with col2:
        if st.button("فودافون"):
            out = generate_vodafone_report(current_df)
            st.download_button("تحميل", out, "vodafone.xlsx")

    with col3:
        if st.button("أورانج"):
            out = generate_orange_report(current_df)
            st.download_button("تحميل", out, "orange.xlsx")
