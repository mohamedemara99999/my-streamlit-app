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
            st.experimental_rerun()
        else:
            st.error("❌ بيانات غير صحيحة")

# ================== بعد تسجيل الدخول ==================
if st.session_state.logged_in:

    st.sidebar.success(f"مرحباً {st.session_state.current_user}")
    if st.sidebar.button("تسجيل خروج"):
        st.session_state.logged_in = False
        st.session_state.current_user = None
        st.experimental_rerun()

    st.title("Excel Analyzer Tool - Streamlit")

    selected_company = st.selectbox(
        "اختر الشركة",
        ["etisalat", "vodafone", "orange"]
    )

    uploaded_file = st.file_uploader("اختر ملف Excel", type=["xlsx", "xls"])

    current_df = None
    original_df = None

    if uploaded_file is not None:
        try:
            if selected_company == "orange":
                current_df = pd.read_excel(uploaded_file, header=4, engine="openpyxl")
            else:
                current_df = pd.read_excel(uploaded_file, engine="openpyxl")

            original_df = current_df.copy()

            current_df.columns = current_df.columns.str.strip()
            current_df = current_df.loc[:, ~current_df.columns.str.contains('^Unnamed')]
            current_df = current_df.dropna(how='all', axis=1)

            st.success("تم فتح الملف بنجاح")
            st.dataframe(current_df)

        except Exception as e:
            st.error(f"خطأ في قراءة الملف: {e}")

# ================== دالة تنسيق Excel ==================
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

# ================== تقرير اتصالات ==================
def generate_etisalat_report(df, original_df):
    df = df.copy()

    df['Originating_Number'] = df['Originating_Number'].astype(str)
    df['Terminating_Number'] = df['Terminating_Number'].astype(str)

    numbers = pd.concat([df['Originating_Number'], df['Terminating_Number']])
    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number', 'Count']

    b_data = {}

    for _, row in df.iterrows():
        for col in ['Originating_Number', 'Terminating_Number']:
            num = str(row[col])
            if num not in b_data:
                b_data[num] = {
                    'B Full Name': row.get('B_Number_Full_Name', ''),
                    'B Address': row.get('B_Number_Address', ''),
                    'B_NUMBER_SITE_ADDRESS': row.get('B_Number_MU_Site_Address', ''),
                    'Latitude': row.get('B_Number_MU_Latitude', ''),
                    'Longitude': row.get('B_Number_MU_Longitude', '')
                }

    df_final = freq.copy()

    for col in ['B Full Name','B Address','B_NUMBER_SITE_ADDRESS','Latitude','Longitude']:
        df_final[col] = df_final['B Number'].map(lambda x: b_data[x][col] if x in b_data else '')

    df_final['Map'] = df_final.apply(
        lambda r: f'https://www.google.com/maps/search/?api=1&query={r["Latitude"]},{r["Longitude"]}'
        if r['Latitude'] != '' else '',
        axis=1
    )

    temp_df = df.copy()
    temp_df['activity_clean'] = temp_df['Network_Activity_Type_Name'].astype(str)

    sms_stats = temp_df.groupby('Originating_Number').agg(
        SMS=('activity_clean', lambda x: (x == 'SMS').sum())
    ).reset_index()

    df_final = df_final.merge(sms_stats, left_on='B Number', right_on='Originating_Number', how='left').drop(columns='Originating_Number')

    df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)
    df_final['Count'] = df_final['Count'].astype(int)

    temp_df['Call_Start_Date'] = pd.to_datetime(temp_df['Call_Start_Date'], errors='coerce')

    calls = pd.concat([
        temp_df[['Originating_Number','Call_Start_Date']].rename(columns={'Originating_Number':'B Number'}),
        temp_df[['Terminating_Number','Call_Start_Date']].rename(columns={'Terminating_Number':'B Number'})
    ])

    first_last = calls.groupby('B Number').agg(
        First_Call=('Call_Start_Date','min'),
        Last_Call=('Call_Start_Date','max')
    ).reset_index()

    df_final = df_final.merge(first_last, on='B Number', how='left')

    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='calls', index=False)
        original_df.to_excel(writer, sheet_name='cheet', index=False)

    output.seek(0)
    return format_excel_sheets(output)

# ================== تقرير فودافون ==================
def generate_vodafone_report(df):
    df = df.copy()

    df['B_NUMBER'] = df['B_NUMBER'].astype(str)

    freq = df['B_NUMBER'].value_counts().reset_index()
    freq.columns = ['B Number', 'Count']

    base = df[['B_NUMBER','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS']].drop_duplicates()

    df_final = freq.merge(base, left_on='B Number', right_on='B_NUMBER', how='left')

    df_final['Count'] = df_final['Count'].astype(int)

    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='calls', index=False)
        df.to_excel(writer, sheet_name='cheet', index=False)

    output.seek(0)
    return format_excel_sheets(output, header_color="FF0000", company="vodafone")

# ================== تقرير أورانج ==================
def generate_orange_report(df):
    df = df.copy()

    df.columns = df.columns.str.upper()

    freq = df['OTHER_MSISDN'].value_counts().reset_index()
    freq.columns = ['B Number', 'Count']

    base = df[['OTHER_MSISDN','OTHER_NAME','OTHER_ADDRESS']].drop_duplicates()

    df_final = freq.merge(base, left_on='B Number', right_on='OTHER_MSISDN', how='left')

    df_final['Count'] = df_final['Count'].astype(int)

    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='calls', index=False)
        df.to_excel(writer, sheet_name='cheet', index=False)

    output.seek(0)
    return format_excel_sheets(output, header_color="FF6600", company="orange")

# ================== أزرار التحليل ==================
if current_df is not None:

    st.subheader("توليد تقارير")

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("تقرير اتصالات"):
            output = generate_etisalat_report(current_df, original_df)
            st.download_button("تحميل", output, "etisalat.xlsx")

    with col2:
        if st.button("تقرير فودافون"):
            output = generate_vodafone_report(current_df)
            st.download_button("تحميل", output, "vodafone.xlsx")

    with col3:
        if st.button("تقرير أورانج"):
            output = generate_orange_report(current_df)
            st.download_button("تحميل", output, "orange.xlsx")
