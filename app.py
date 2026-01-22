import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from io import BytesIO

# ================== تسجيل الدخول ==================
USERS = {
    "admin": "m7md3mara2025",
    "user1": "1234",
    "user2": "5678",
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
    original_df = None   # 👈 نسخة الشيت الأصلي

    if uploaded_file is not None:
        try:
            if selected_company == "orange":
                current_df = pd.read_excel(uploaded_file, header=4, engine="openpyxl")
            else:
                current_df = pd.read_excel(uploaded_file, engine="openpyxl")

            # ===== حفظ نسخة أصلية قبل أي تعديل =====
            original_df = current_df.copy()

            # ===== تنظيف الأعمدة =====
            current_df.columns = current_df.columns.str.strip()
            current_df = current_df.loc[:, ~current_df.columns.str.contains('^Unnamed')]
            current_df = current_df.dropna(how='all', axis=1)

            st.success("تم فتح الملف بنجاح")
            st.dataframe(current_df)

        except Exception as e:
            st.error(f"خطأ في قراءة الملف: {e}")

# ================== دالة تنسيق Excel ==================
def format_excel_sheets(output, header_color="228B22"):
    output.seek(0)
    wb = load_workbook(output)

    header_fill = PatternFill("solid", fgColor=header_color)
    header_font = Font(bold=True, color="FFFFFF")

    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        # تحويل أي روابط هايبرلينك
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("http"):
                    cell.hyperlink = cell.value
                    if "google.com/maps" in cell.value:
                        cell.value = "Map"
                    elif "imei.info" in cell.value:
                        cell.value = "IMEI Info"
                    cell.font = Font(color="006400", underline="single")

    final = BytesIO()
    wb.save(final)
    final.seek(0)
    return final

# ================== تقرير اتصالات ==================
def generate_etisalat_report(df, original_df):

    df = df.copy()
    df['Originating_Number'] = df['Originating_Number'].astype(str)
    df['Terminating_Number'] = df['Terminating_Number'].astype(str)

    # ===== حساب التكرار =====
    numbers = pd.concat([df['Originating_Number'], df['Terminating_Number']])
    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number', 'Count']

    # ===== إنشاء dictionary للـ B Data =====
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

    # ===== Map =====
    df_final['Map'] = df_final.apply(
        lambda r: f'https://www.google.com/maps/search/?api=1&query={r["Latitude"]},{r["Longitude"]}'
        if pd.notna(r['Latitude']) and r['Latitude'] != '' else '',
        axis=1
    )

    # ===== حساب SMS من Originating فقط =====
    temp_df = df.copy()
    temp_df['activity_clean'] = temp_df['Network_Activity_Type_Name'].astype(str).str.strip()
    sms_stats = temp_df.groupby('Originating_Number').agg(
        SMS=('activity_clean', lambda x: (x == 'SMS').sum())
    ).reset_index()

    df_final = df_final.merge(
        sms_stats, left_on='B Number', right_on='Originating_Number', how='left'
    ).drop(columns='Originating_Number')

    df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)
    df_final['Count'] = df_final['Count'].astype(int)

    # ===== First / Last Call =====
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

    # ===== IMEI =====
    def safe_imei(x):
        try:
            return str(int(float(x)))
        except:
            return ''

    imei_df = df.copy()
    imei_df['IMEI_Number'] = imei_df['IMEI_Number'].apply(safe_imei)
    imei_summary = imei_df.groupby('IMEI_Number').agg(
        Count=('IMEI_Number','count'),
        First_Use_Date=('Call_Start_Date','min'),
        Last_Use_Date=('Call_Start_Date','max'),
        First_Use_Address=('Site_Address','first'),
        Last_Use_Address=('Site_Address','last')
    ).reset_index()
    imei_summary.rename(columns={'IMEI_Number':'IMEI'}, inplace=True)
    imei_summary['Device Info'] = imei_summary['IMEI'].apply(lambda x: f'https://www.imei.info/calc/?imei={x}')
    imei_summary = imei_summary.sort_values(by='Count', ascending=False)

    # ===== Sites =====
    site_df = df[['Site_Address','Latitude','Longitude','Call_Start_Date']].copy()
    site_group = site_df.groupby('Site_Address').agg(
        Count=('Site_Address','count'),
        First_Use_Date=('Call_Start_Date','min'),
        Last_Use_Date=('Call_Start_Date','max'),
        Latitude=('Latitude','first'),
        Longitude=('Longitude','first')
    ).reset_index()
    site_group['Map'] = site_group.apply(
        lambda r: f'https://www.google.com/maps/search/?api=1&query={r["Latitude"]},{r["Longitude"]}', axis=1
    )
    site_group = site_group[['Site_Address','Count','Map','First_Use_Date','Last_Use_Date']].sort_values(by='Count', ascending=False)

    # ===== إخراج Excel =====
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='calls', index=False)
        imei_summary.to_excel(writer, sheet_name='imei', index=False)
        site_group.to_excel(writer, sheet_name='site', index=False)
        original_df.to_excel(writer, sheet_name='cheet', index=False)    # الشيت الأصلي

    output.seek(0)
    return format_excel_sheets(output)

# ================== تقرير فودافون ==================
def generate_vodafone_report(df):
    required_cols = [
        'B_NUMBER','B_NUMBER_FIRST_NAME','B_NUMBER_LAST_NAME','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS',
        'B_NUMBER_NATIONAL_ID','IMEI','HANDSET_MANUFACTURER','HANDSET_MARKETING_NAME',
        'FULL_DATE','SITE_ADDRESS','LATITUDE','LONGITUDE','SERVICE'
    ]
    for col in required_cols:
        if col not in df.columns:
            st.error(f"العمود {col} غير موجود في الملف")
            return None

    df2 = df.copy()
    df2['B Full Name'] = df2['B_NUMBER_FIRST_NAME'].fillna('') + ' ' + df2['B_NUMBER_LAST_NAME'].fillna('')
    df2['IMEI'] = df2['IMEI'].astype(str)
    numbers = df2['B_NUMBER'].astype(str)
    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number','Count']

    # ===== حساب SMS =====
    sms_count = df2[df2['SERVICE'].astype(str).str.strip().isin(["Short message MO/PP","Short message MT/PP"])].groupby('B_NUMBER').size().reset_index(name='SMS')

    # ===== دمج البيانات =====
    df_final = freq.merge(
        df2[['B_NUMBER','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS','B_NUMBER_NATIONAL_ID']].drop_duplicates(subset='B_NUMBER'),
        left_on='B Number', right_on='B_NUMBER', how='left'
    )
    df_final = df_final.merge(sms_count, left_on='B Number', right_on='B_NUMBER', how='left')
    df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)

    # ===== إضافة B Number id بعد B Full Name =====
    df_final['B Number id'] = df_final['B_NUMBER_NATIONAL_ID'].astype(str)

    # ===== ترتيب الأعمدة النهائي =====
    df_final = df_final[['B Number','Count','B Full Name','B Number id','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS','SMS']]
    df_final['Count'] = df_final['Count'].astype(int)
    df_final = df_final.sort_values(by='Count', ascending=False)

    # ===== تجميع بيانات IMEI =====
    df2['FULL_DATE'] = pd.to_datetime(df2['FULL_DATE'])
    imei_group = df2.groupby('IMEI').agg(
        Count=('IMEI','count'),
        Device_Info=('IMEI', lambda x: f'https://www.imei.info/calc/?imei={x.iloc[0]}'),
        HANDSET_MANUFACTURER=('HANDSET_MANUFACTURER','first'),
        HANDSET_MARKETING_NAME=('HANDSET_MARKETING_NAME','first'),
        First_Use_Date=('FULL_DATE','min'),
        Last_Use_Date=('FULL_DATE','max')
    ).reset_index()

    # ===== أول وآخر عنوان لكل IMEI =====
    first_last_addr = []
    for imei in imei_group['IMEI']:
        sub = df2[df2['IMEI']==imei].sort_values('FULL_DATE')
        first_addr = sub.iloc[0]['SITE_ADDRESS']
        last_addr = sub.iloc[-1]['SITE_ADDRESS']
        first_last_addr.append((first_addr,last_addr))
    imei_group['First_Use_Address'] = [x[0] for x in first_last_addr]
    imei_group['Last_Use_Address'] = [x[1] for x in first_last_addr]

    imei_group = imei_group[['IMEI','Count','Device_Info','HANDSET_MANUFACTURER','HANDSET_MARKETING_NAME',
                             'First_Use_Date','Last_Use_Date','First_Use_Address','Last_Use_Address']]
    imei_group['Count'] = imei_group['Count'].astype(int)
    imei_group = imei_group.sort_values(by='Count', ascending=False)

    # ===== تجميع بيانات المواقع =====
    site_df = df2[['SITE_ADDRESS','LATITUDE','LONGITUDE','FULL_DATE']].copy()
    site_group = site_df.groupby('SITE_ADDRESS').agg(
        Count=('SITE_ADDRESS','count'),
        Map=('LATITUDE', lambda x: f'https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"LONGITUDE"]}'),
        First_Use_Date=('FULL_DATE','min'),
        Last_Use_Date=('FULL_DATE','max')
    ).reset_index()
    site_group['Count'] = site_group['Count'].astype(int)
    site_group = site_group.sort_values(by='Count', ascending=False)
    site_group = site_group[['SITE_ADDRESS','Count','Map','First_Use_Date','Last_Use_Date']]

    # ===== حفظ Excel في BytesIO =====
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name="calls", index=False)
        imei_group.to_excel(writer, sheet_name="imei", index=False)
        site_group.to_excel(writer, sheet_name="site", index=False)
    output.seek(0)

    # ===== تطبيق التنسيقات =====
    try:
        final_output = format_excel_sheets(output, header_color="FF0000")
    except Exception as e:
        st.error(f"خطأ في تطبيق التنسيقات: {e}")
        return output

    return final_output

# ================== أزرار التحليل ==================
if st.session_state.logged_in and uploaded_file is not None:
    if selected_company == "etisalat":
        if st.button("تحليل الملف - اتصالات"):
            result = generate_etisalat_report(current_df, original_df)
            if result:
                st.download_button(
                    "تحميل تقرير اتصالات",
                    data=result,
                    file_name="etisalat_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    elif selected_company == "vodafone":
        if st.button("تحليل الملف - فودافون"):
            result = generate_vodafone_report(current_df)
            if result:
                st.download_button(
                    "تحميل تقرير فودافون",
                    data=result,
                    file_name="vodafone_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
