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
        ["etisalat_معمل", "etisalat_شركه", "vodafone", "orange"]

    )

    uploaded_file = st.file_uploader("اختر ملف Excel", type=["xlsx", "xls"])

    current_df = None
    original_df = None

    # نسخة الشيت الأصلي
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

            # ===== تحقق من الأعمدة =====
            if selected_company == "etisalat" and 'Originating_Number' not in current_df.columns:
                st.error("ملف غير صالح لاتصالات")
            else:
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

    first_row_fill_calls = PatternFill("solid", fgColor="FFFF00")
    first_row_font_calls = Font(bold=True, color="000000")

    for ws in wb.worksheets:

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        if company.lower() == "etisalat" and ws.title.lower() == "calls" and ws.max_row > 1:
            for cell in ws[2]:
                cell.fill = first_row_fill_calls
                cell.font = first_row_font_calls

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
        if pd.notna(r['Latitude']) and r['Latitude'] != '' else '',
        axis=1
    )

    temp_df = df.copy()
    temp_df['activity_clean'] = temp_df['Network_Activity_Type_Name'].astype(str).str.strip()

    sms_stats = temp_df.groupby('Originating_Number').agg(
        SMS=('activity_clean', lambda x: (x == 'SMS').sum())
    ).reset_index()

    df_final = df_final.merge(
        sms_stats,
        left_on='B Number',
        right_on='Originating_Number',
        how='left'
    ).drop(columns='Originating_Number')

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

    if not df_final.empty:
        top_number = df_final.iloc[0]['B Number']
        mask = df_final['B Number'] == top_number

        df_final.loc[mask, [
            'B Full Name','B Address','B_NUMBER_SITE_ADDRESS',
            'Latitude','Longitude','Map','SMS'
        ]] = [
            f"{df.iloc[0].get('A_Number_Details_First_Name','')} {df.iloc[0].get('A_Number_Details_Last_Name','')}",
            '28607102800033',
            df.iloc[0].get('MU_Site_Address',''),
            '',
            '',
            '',
            0
        ]

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

    imei_summary['Device Info'] = imei_summary['IMEI'].apply(
        lambda x: f'https://www.imei.info/calc/?imei={x}'
    )

    imei_summary = imei_summary.sort_values(by='Count', ascending=False)

    site_df = df[['Site_Address','Latitude','Longitude','Call_Start_Date']].copy()

    site_group = site_df.groupby('Site_Address').agg(
        Count=('Site_Address','count'),
        First_Use_Date=('Call_Start_Date','min'),
        Last_Use_Date=('Call_Start_Date','max'),
        Latitude=('Latitude','first'),
        Longitude=('Longitude','first')
    ).reset_index()

    site_group['Map'] = site_group.apply(
        lambda r: f'https://www.google.com/maps/search/?api=1&query={r["Latitude"]},{r["Longitude"]}',
        axis=1
    )

    site_group = site_group[['Site_Address','Count','Map','First_Use_Date','Last_Use_Date']].sort_values(by='Count', ascending=False)

    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='calls', index=False)
        imei_summary.to_excel(writer, sheet_name='imei', index=False)
        site_group.to_excel(writer, sheet_name='site', index=False)
        original_df.to_excel(writer, sheet_name='cheet', index=False)

    output.seek(0)
    return format_excel_sheets(output, header_color="006400", company="etisalat")

# ================== تقرير فودافون ==================
def generate_vodafone_report(df):
    required_cols = [
        'B_NUMBER','B_NUMBER_NATIONAL_ID','B_NUMBER_SITE_ADDRESS',
        'FULL_DATE','SERVICE',
        'IMEI','HANDSET_MANUFACTURER','HANDSET_MARKETING_NAME',
        'SITE_ADDRESS','LATITUDE','LONGITUDE'
    ]

    for col in required_cols:
        if col not in df.columns:
            st.error(f"العمود {col} غير موجود في الملف")
            return None

    df2 = df.copy()
    df2['B_NUMBER'] = df2['B_NUMBER'].astype(str)
    df2['IMEI'] = df2['IMEI'].astype(str)
    df2['FULL_DATE'] = pd.to_datetime(df2['FULL_DATE'], errors='coerce')

    freq = df2['B_NUMBER'].value_counts().reset_index()
    freq.columns = ['B Number','Count']

    sms_count = (
        df2[df2['SERVICE'].astype(str).str.strip()
            .isin(["Short message MO/PP","Short message MT/PP"])]
        .groupby('B_NUMBER')
        .size()
        .reset_index(name='SMS')
    )

    call_dates = (
        df2.groupby('B_NUMBER')['FULL_DATE']
        .agg(First_Call='min', Last_Call='max')
        .reset_index()
    )

    base_info = (
        df2[['B_NUMBER','B_NUMBER_NATIONAL_ID','B_NUMBER_SITE_ADDRESS']]
        .drop_duplicates(subset='B_NUMBER')
    )

    df_final = freq.merge(
        base_info,
        left_on='B Number',
        right_on='B_NUMBER',
        how='left'
    )

    df_final = df_final.merge(
        sms_count,
        left_on='B Number',
        right_on='B_NUMBER',
        how='left'
    )

    df_final = df_final.merge(
        call_dates,
        left_on='B Number',
        right_on='B_NUMBER',
        how='left'
    )

    df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)
    df_final['Count'] = df_final['Count'].astype(int)
    df_final['B Number Id'] = df_final['B_NUMBER_NATIONAL_ID'].astype(str)

    df_final = df_final[
        ['B Number','Count','B Number Id',
         'B_NUMBER_SITE_ADDRESS','SMS','First_Call','Last_Call']
    ].sort_values(by='Count', ascending=False)

    imei_group = df2.groupby('IMEI').agg(
        Count=('IMEI','count'),
        Device_Info=('IMEI', lambda x: f'https://www.imei.info/calc/?imei={x.iloc[0]}'),
        HANDSET_MANUFACTURER=('HANDSET_MANUFACTURER','first'),
        HANDSET_MARKETING_NAME=('HANDSET_MARKETING_NAME','first'),
        First_Use_Date=('FULL_DATE','min'),
        Last_Use_Date=('FULL_DATE','max')
    ).reset_index()

    first_last_addr = []

    for imei in imei_group['IMEI']:
        sub = df2[df2['IMEI'] == imei].sort_values('FULL_DATE')
        first_last_addr.append(
            (sub.iloc[0]['SITE_ADDRESS'], sub.iloc[-1]['SITE_ADDRESS'])
        )

    imei_group['First_Use_Address'] = [x[0] for x in first_last_addr]
    imei_group['Last_Use_Address'] = [x[1] for x in first_last_addr]

    imei_group['Count'] = imei_group['Count'].astype(int)
    imei_group = imei_group.sort_values(by='Count', ascending=False)

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

    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name="calls", index=False)
        imei_group.to_excel(writer, sheet_name="imei", index=False)
        site_group.to_excel(writer, sheet_name="site", index=False)
        df.to_excel(writer, sheet_name="cheet", index=False)

    output.seek(0)
    return format_excel_sheets(output, header_color="FF0000", company="vodafone")


# ================== تقرير أورانج ==================
def generate_orange_report(df):
    df.columns = df.columns.str.strip().str.upper()

    required_cols = [
        'TARGET_MSISDN','TARGET_IMEI','TARGET_IMSI','TARGET_IMEI_TYPE',
        'EVENT_START_TIME','CALL_DURATION','EVENT_DIRECTION',
        'OTHER_MSISDN','OTHER_NAME','OTHER_ID','OTHER_ADDRESS',
        'CELL_ADDRESS','CELL_LAT','CELL_LONG','OTHER_CELL_ADDRESS'
    ]

    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        st.error(f"الأعمدة التالية غير موجودة في الملف: {missing_cols}")
        return None

    numbers = df['OTHER_MSISDN'].astype(str)
    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number','Count']

    base_info = (
        df[['OTHER_MSISDN','OTHER_NAME','OTHER_ADDRESS','OTHER_ID','OTHER_CELL_ADDRESS']]
        .drop_duplicates(subset='OTHER_MSISDN')
    )

    calls_df = freq.merge(
        base_info,
        left_on='B Number',
        right_on='OTHER_MSISDN',
        how='left'
    )

    sms_count = (
        df[df['EVENT_DIRECTION'].astype(str).str.strip() == "SMSMT"]
        .groupby('OTHER_MSISDN')
        .size()
        .reset_index(name='SMS')
    )

    calls_df = calls_df.merge(
        sms_count,
        left_on='B Number',
        right_on='OTHER_MSISDN',
        how='left'
    )

    calls_df['SMS'] = calls_df['SMS'].fillna(0).astype(int)
    calls_df['Count'] = calls_df['Count'].astype(int)

    calls_df['B Number id'] = calls_df['OTHER_ID'].apply(
        lambda x: str(int(x)) if pd.notna(x) else ''
    )

    calls_df = calls_df[
        ['B Number','Count','OTHER_NAME','OTHER_ADDRESS',
         'B Number id','OTHER_CELL_ADDRESS','SMS']
    ]

    calls_df.columns = [
        'B Number','Count','B Full Name','B Address',
        'B Number id','other site','SMS'
    ]

    df['EVENT_START_TIME'] = pd.to_datetime(df['EVENT_START_TIME'], errors='coerce')

    call_dates = (
        df.groupby('OTHER_MSISDN')['EVENT_START_TIME']
        .agg(First_Call='min', Last_Call='max')
        .reset_index()
    )

    calls_df = calls_df.merge(
        call_dates,
        left_on='B Number',
        right_on='OTHER_MSISDN',
        how='left'
    ).drop(columns='OTHER_MSISDN')

    calls_df = calls_df.sort_values(by='Count', ascending=False)

    df['TARGET_IMEI'] = df['TARGET_IMEI'].apply(
        lambda x: str(int(x)) if pd.notna(x) else ''
    )

    imei_group = df.groupby('TARGET_IMEI').agg(
        Count=('TARGET_IMEI','count'),
        TARGET_IMEI_TYPE=('TARGET_IMEI_TYPE','first'),
        First_Use_Date=('EVENT_START_TIME','min'),
        Last_Use_Date=('EVENT_START_TIME','max'),
        First_Use_Address=('CELL_ADDRESS','first'),
        Last_Use_Address=('CELL_ADDRESS','last')
    ).reset_index()

    imei_group['Device Info'] = imei_group['TARGET_IMEI'].apply(
        lambda x: f'https://www.imei.info/calc/?imei={x}'
    )

    imei_group = imei_group[
        ['TARGET_IMEI','Count','TARGET_IMEI_TYPE','Device Info',
         'First_Use_Date','Last_Use_Date','First_Use_Address','Last_Use_Address']
    ]

    imei_group.columns = [
        'IMEI','Count','TARGET_IMEI_TYPE','Device Info',
        'First_Use_Date','Last_Use_Date','First_Use_Address','Last_Use_Address'
    ]

    imei_group['Count'] = imei_group['Count'].astype(int)
    imei_group = imei_group.sort_values(by='Count', ascending=False)

    site_df = df.groupby('CELL_ADDRESS').agg(
        Count=('CELL_ADDRESS','count'),
        First_Use_Date=('EVENT_START_TIME','min'),
        Last_Use_Date=('EVENT_START_TIME','max'),
        LAT=('CELL_LAT','first'),
        LON=('CELL_LONG','first')
    ).reset_index()

    site_df['Map'] = site_df.apply(
        lambda row: f'https://www.google.com/maps/search/?api=1&query={row["LAT"]},{row["LON"]}'
        if pd.notna(row["LAT"]) and pd.notna(row["LON"]) else '',
        axis=1
    )

    site_df = site_df[['CELL_ADDRESS','Count','Map','First_Use_Date','Last_Use_Date']]
    site_df = site_df.sort_values(by='Count', ascending=False)

    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        calls_df.to_excel(writer, sheet_name="calls", index=False)
        imei_group.to_excel(writer, sheet_name="imei", index=False)
        site_df.to_excel(writer, sheet_name="site", index=False)
        df.to_excel(writer, sheet_name="cheet", index=False)

    output.seek(0)
    return format_excel_sheets(output, header_color="FF6600", company="orange")

def generate_etisalat_new_report(df):
    df = df.copy()

    # ================= تنظيف =================
    df['Subscriber_Number'] = df['Subscriber_Number'].astype(str)
    df['b_number_full'] = df['b_number_full'].astype(str)

    df['Call_Start_Date'] = pd.to_datetime(df['Call_Start_Date'], errors='coerce')

    # ================= Frequency =================
    numbers = pd.concat([df['Subscriber_Number'], df['b_number_full']])
    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number', 'Count']

    # ================= Base Info =================
    base_info = df[[
        'b_number_full',
        'B_Number_Full_Name',
        'B_Number_Address',
        'B_Number_NID',
        'B_Number_Most_Location',
        'B_Num_Most_Location_Address'
    ]].drop_duplicates(subset='b_number_full')

    base_info.rename(columns={'b_number_full': 'B Number'}, inplace=True)

    # ================= Calls =================
    df_final = freq.merge(base_info, on='B Number', how='left')

    sms_stats = df.groupby('b_number_full').agg(
        SMS=('Service', lambda x: (x.astype(str).str.contains("SMS")).sum())
    ).reset_index()

    sms_stats.rename(columns={'b_number_full': 'B Number'}, inplace=True)

    df_final = df_final.merge(sms_stats, on='B Number', how='left')
    df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)

    # ================= Calls Time =================
    calls = pd.concat([
        df[['Subscriber_Number', 'Call_Start_Date']].rename(columns={'Subscriber_Number': 'B Number'}),
        df[['b_number_full', 'Call_Start_Date']].rename(columns={'b_number_full': 'B Number'})
    ])

    time_stats = calls.groupby('B Number').agg(
        First_Call=('Call_Start_Date', 'min'),
        Last_Call=('Call_Start_Date', 'max')
    ).reset_index()

    df_final = df_final.merge(time_stats, on='B Number', how='left')

    # ================= IMEI =================
    imei_df = df.copy()
    imei_df['Subscriber_IMEI'] = imei_df['Subscriber_IMEI'].astype(str)

    imei_group = imei_df.groupby('Subscriber_IMEI').agg(
        Count=('Subscriber_IMEI', 'count'),
        First_Use=('Call_Start_Date', 'min'),
        Last_Use=('Call_Start_Date', 'max'),
        First_Location=('Subscriber_Location_Address', 'first'),
        Last_Location=('Subscriber_Location_Address', 'last')
    ).reset_index()

    imei_group.rename(columns={'Subscriber_IMEI': 'IMEI'}, inplace=True)

    imei_group['Device Info'] = imei_group['IMEI'].apply(
        lambda x: f'https://www.imei.info/calc/?imei={x}'
    )

    # ================= SITE =================
    site_group = df.groupby('Subscriber_Location_Address').agg(
        Count=('Subscriber_Location_Address', 'count'),
        First_Use=('Call_Start_Date', 'min'),
        Last_Use=('Call_Start_Date', 'max'),
        LAT=('Subscriber_Location', 'first'),
        LON=('Subscriber_Location', 'first')
    ).reset_index()

    site_group['Map'] = site_group.apply(
        lambda r: f'https://www.google.com/maps/search/?api=1&query={r["LAT"]},{r["LON"]}'
        if pd.notna(r["LAT"]) else '',
        axis=1
    )

    # ================= OUTPUT =================
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='calls', index=False)
        imei_group.to_excel(writer, sheet_name='imei', index=False)
        site_group.to_excel(writer, sheet_name='site', index=False)
        df.to_excel(writer, sheet_name='cheet', index=False)

    output.seek(0)
    return format_excel_sheets(output, header_color="006400", company="etisalat")
# ================== أزرار التقارير ==================
if current_df is not None:
    st.subheader("توليد تقارير")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        if st.button("تقرير اتصالات معمل"):
            output = generate_etisalat_report(current_df, original_df, "etisalat")
            if output:
                st.download_button("تحميل اتصالات", output, "etisalat.xlsx")

    with col2:
        if st.button("تقرير اتصالات شركة"):
            output = generate_etisalat_report(current_df, original_df, "etisalat_company")
            if output:
                st.download_button("تحميل اتصالات شركة", output, "etisalat_company.xlsx")

    with col3:
        if st.button("تقرير فودافون"):
            st.info("نفس الدالة الحالية")

    with col4:
        if st.button("تقرير أورانج"):
            st.info("نفس الدالة الحالية")
