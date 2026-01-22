import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from io import BytesIO

# ================== قائمة المستخدمين ==================
USERS = {
    "admin": "m7md3mara2025",
    "user1": "1234",
    "user2": "5678",
    "user3": "2468"
}
if "active_sessions" not in st.session_state:
    st.session_state.active_sessions = {}

# حالة تسجيل الدخول
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "current_user" not in st.session_state:
    st.session_state.current_user = None

# ================== صفحة تسجيل الدخول ==================
if not st.session_state.logged_in:
    st.title("🔐 تسجيل الدخول")

    username = st.text_input("اسم المستخدم")
    password = st.text_input("كلمة المرور", type="password")

    if st.button("دخول"):
        if username in USERS and USERS[username] == password:
            # ===== منع الدخول المزدوج =====
            if username in st.session_state.active_sessions and st.session_state.active_sessions[username]:
                st.error("❌ هذا الحساب يستخدم حالياً على جهاز آخر")
            else:
                st.session_state.logged_in = True
                st.session_state.current_user = username
                st.session_state.active_sessions[username] = True
                st.experimental_rerun()

        else:
            st.error("❌ اسم المستخدم أو كلمة المرور غير صحيحة")

# ================== لو المستخدم سجل دخول ==================
if st.session_state.logged_in:
    st.sidebar.success(f"مرحباً يا {st.session_state.current_user} 👋")

    # ===== زر تسجيل الخروج =====
    st.sidebar.button("تسجيل خروج", on_click=lambda: (
        st.session_state.active_sessions.update({st.session_state.current_user: False}),
        st.session_state.update({"logged_in": False, "current_user": None}),
        st.experimental_rerun()
    ))

    st.title("Excel Analyzer Tool - Streamlit")
    

    selected_company = st.selectbox(
        "اختر الشركة",
        ["etisalat", "vodafone", "orange"]
    )
    uploaded_file = st.file_uploader("اختر ملف Excel", type=["xlsx", "xls"])
    current_df = None
    if uploaded_file is not None:
        try:
            if selected_company == "orange":
                # أورانج يبدأ الهيدر من الصف الخامس (B5)
                current_df = pd.read_excel(uploaded_file, header=4, engine="openpyxl")
            else:
                current_df = pd.read_excel(uploaded_file, engine="openpyxl")
                # ===== تنظيف الأعمدة =====
            current_df.columns = current_df.columns.str.strip()  # إزالة الفراغات
            # حذف الأعمدة كلها فارغة أو مسماها Unnamed
            current_df = current_df.loc[:, ~current_df.columns.str.contains('^Unnamed')]
            current_df = current_df.dropna(how='all', axis=1)  # حذف الأعمدة الفارغة تمامًا

        # ===== التحقق من تطابق الأعمدة مع الشركة =====
            if selected_company == "etisalat" and 'Originating_Number' not in current_df.columns:
                st.error("ملف غير صالح لشركة اتصالات.")
            elif selected_company == "vodafone" and 'B_NUMBER' not in current_df.columns:
                st.error("ملف غير صالح لشركة فودافون.")
            elif selected_company == "orange" and 'OTHER_MSISDN' not in current_df.columns:
                st.error("ملف غير صالح لشركة اورانج.")
            else:
                st.success(f"تم فتح الملف: {uploaded_file.name}")
                st.dataframe(current_df)
        except Exception as e:
            st.error(f"خطأ في فتح الملف: {e}")

# ================== دوال تنسيق Excel ==================
def format_excel_sheets(output, header_color="228B22", highlight_row=None, highlight_color="FFFF00"):
    output.seek(0)
    wb = load_workbook(output)

    header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=14)
    green_link_font = Font(color="006400", underline="single")

    for ws in wb.worksheets:
        # ===== تلوين الهيدر في كل الصفحات =====
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        # ===== تلوين الصف المستثنى فقط في ورقة calls =====
        if highlight_row and ws.title.lower() == "calls":
            for cell in ws[highlight_row]:
                cell.fill = PatternFill(start_color=highlight_color, end_color=highlight_color, fill_type="solid")

        # ===== تحويل الروابط لكلمة مختصرة مع الحفاظ على هايبرلينك =====
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("http"):
                    url = cell.value
                    if "google.com/maps" in url:
                        cell.value = "map"
                    elif "imei.info" in url:
                        cell.value = "check info"
                    cell.hyperlink = url
                    cell.font = green_link_font

    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

def generate_etisalat_report(df):
    import pandas as pd
    from io import BytesIO

    # ===== تنظيف الأعمدة =====
    df = df.copy()
    df.columns = df.columns.str.strip()
    df['Originating_Number'] = df['Originating_Number'].astype(str)
    df['Terminating_Number'] = df['Terminating_Number'].astype(str)

    # ===== حساب تكرار الأرقام (من الطرفين) =====
    numbers = pd.concat([df['Originating_Number'], df['Terminating_Number']])
    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number', 'Count']
    df_final = freq.copy()

    # ===== dictionary لربط أي رقم ببياناته =====
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

    for col in ['B Full Name','B Address','B_NUMBER_SITE_ADDRESS','Latitude','Longitude']:
        df_final[col] = df_final['B Number'].map(lambda x: b_data[x][col] if x in b_data else '')

    # ===== Map =====
    df_final['Map'] = df_final.apply(
        lambda r: f'https://www.google.com/maps/search/?api=1&query={r["Latitude"]},{r["Longitude"]}'
        if pd.notna(r['Latitude']) and pd.notna(r['Longitude']) and r['Latitude'] != '' else '',
        axis=1
    )

    # ===== SMS (من Originating فقط – زي كود 2 بالظبط) =====
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

    # ===== First / Last Call (من الطرفين) =====
    temp_df['Call_Start_Date'] = pd.to_datetime(temp_df['Call_Start_Date'])
    calls = pd.concat([
        temp_df[['Originating_Number','Call_Start_Date']]
            .rename(columns={'Originating_Number':'B Number'}),
        temp_df[['Terminating_Number','Call_Start_Date']]
            .rename(columns={'Terminating_Number':'B Number'})
    ])

    first_last = calls.groupby('B Number').agg(
        First_Call=('Call_Start_Date','min'),
        Last_Call=('Call_Start_Date','max')
    ).reset_index()

    df_final = df_final.merge(first_last, on='B Number', how='left')
    df_final = df_final.sort_values(by='Count', ascending=False)

    # ===== استثناء أول رقم (نفس كود 2) =====
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
            '', '', '', 0
        ]

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
    imei_summary['Device Info'] = imei_summary['IMEI'].apply(
        lambda x: f'https://www.imei.info/calc/?imei={x}'
    )
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
        lambda r: f'https://www.google.com/maps/search/?api=1&query={r["Latitude"]},{r["Longitude"]}',
        axis=1
    )

    site_group = site_group[['Site_Address','Count','Map','First_Use_Date','Last_Use_Date']]
    site_group = site_group.sort_values(by='Count', ascending=False)

    # ===== إخراج Streamlit =====
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='calls', index=False)
        imei_summary.to_excel(writer, sheet_name='imei', index=False)
        site_group.to_excel(writer, sheet_name='site', index=False)

    output.seek(0)
    return output

def generate_vodafone_report():
    global current_df, current_file
    if current_df is None:
        messagebox.showinfo("مطلوب", "افتح ملف Excel أولاً")
        return

    required_cols = [
        'B_NUMBER','B_NUMBER_FIRST_NAME','B_NUMBER_LAST_NAME','B_NUMBER_NATIONAL_ID','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS',
        'IMEI','HANDSET_MANUFACTURER','HANDSET_MARKETING_NAME','FULL_DATE','SITE_ADDRESS','LATITUDE','LONGITUDE','SERVICE'
        ]
    for col in required_cols:
        if col not in current_df.columns:
            messagebox.showerror("خطأ", f"العمود {col} غير موجود في الملف")
            return

    df = current_df.copy()
    df['B Full Name'] = df['B_NUMBER_FIRST_NAME'].fillna('') + ' ' + df['B_NUMBER_LAST_NAME'].fillna('')

    numbers = df['B_NUMBER'].astype(str)
    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number','Count']

    sms_count = df[df['SERVICE'].astype(str).str.strip().isin(["Short message MO/PP","Short message MT/PP"])].groupby('B_NUMBER').size().reset_index(name='SMS')

    df_final = freq.merge(
        df[['B_NUMBER','B Full Name','B_NUMBER_NATIONAL_ID','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS']].drop_duplicates(subset='B_NUMBER'),
        left_on='B Number', right_on='B_NUMBER', how='left'
    )
    df_final = df_final.merge(sms_count, left_on='B Number', right_on='B_NUMBER', how='left')
    df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)
    df_final['B Number Id'] = df_final['B_NUMBER_NATIONAL_ID'].apply(str)

    # ===== إضافة أعمدة First Call و Last Call =====
    df['FULL_DATE'] = pd.to_datetime(df['FULL_DATE'])
    call_dates = df.groupby('B_NUMBER')['FULL_DATE'].agg(First_Call='min', Last_Call='max').reset_index()
    df_final = df_final.merge(call_dates, left_on='B Number', right_on='B_NUMBER', how='left')
    df_final = df_final.drop(columns=['B_NUMBER'])

    # ترتيب الأعمدة
    df_final = df_final[['B Number','Count','B Full Name','B Number Id','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS','SMS','First_Call','Last_Call']]
    df_final['B Number'] = df_final['B Number'].apply(str)
    df_final['Count'] = df_final['Count'].astype(int)
    df_final = df_final.sort_values(by='Count', ascending=False)

    # ===== صفحة IMEI =====
    df['IMEI'] = df['IMEI'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
    imei_group = df.groupby('IMEI').agg(
        Count=('IMEI','count'),
        Device_Info=('IMEI', lambda x: f'=HYPERLINK("https://www.imei.info/calc/?imei={x.iloc[0]}","IMEI Info")'),
        HANDSET_MANUFACTURER=('HANDSET_MANUFACTURER','first'),
        HANDSET_MARKETING_NAME=('HANDSET_MARKETING_NAME','first'),
        First_Use_Date=('FULL_DATE','min'),
        Last_Use_Date=('FULL_DATE','max')
    ).reset_index()

    first_last_addr = []
    for imei in imei_group['IMEI']:
        sub = df[df['IMEI']==imei].sort_values('FULL_DATE')
        first_addr = sub.iloc[0]['SITE_ADDRESS']
        last_addr = sub.iloc[-1]['SITE_ADDRESS']
        first_last_addr.append((first_addr,last_addr))
    imei_group['First_Use_Address'] = [x[0] for x in first_last_addr]
    imei_group['Last_Use_Address'] = [x[1] for x in first_last_addr]

    imei_group = imei_group[['IMEI','Count','Device_Info','HANDSET_MANUFACTURER','HANDSET_MARKETING_NAME',
                             'First_Use_Date','Last_Use_Date','First_Use_Address','Last_Use_Address']]
    imei_group['Count'] = imei_group['Count'].astype(int)
    imei_group = imei_group.sort_values(by='Count', ascending=False)

    # ===== صفحة SITE =====
    site_df = df[['SITE_ADDRESS','LATITUDE','LONGITUDE','FULL_DATE']].copy()
    site_group = site_df.groupby('SITE_ADDRESS').agg(
        Count=('SITE_ADDRESS','count'),
        Map=('LATITUDE', lambda x: f'=HYPERLINK("https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"LONGITUDE"]}","Map")'),
        First_Use_Date=('FULL_DATE','min'),
        Last_Use_Date=('FULL_DATE','max')
    ).reset_index()
    site_group['Count'] = site_group['Count'].astype(int)
    site_group = site_group.sort_values(by='Count', ascending=False)
    site_group = site_group[['SITE_ADDRESS','Count','Map','First_Use_Date','Last_Use_Date']]

    # ===== حفظ التقرير =====
    output_file = os.path.join(os.path.dirname(current_file), "vodafone_report.xlsx")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name="calls", index=False)
        imei_group.to_excel(writer, sheet_name="imei", index=False)
        site_group.to_excel(writer, sheet_name="site", index=False)
        current_df.to_excel(writer, sheet_name="cheet", index=False)

    wb = load_workbook(output_file)
    format_sheet(wb["calls"], header_color="FF0000")
    format_sheet(wb["imei"], header_color="FF0000", hyperlink_col=3)
    format_sheet(wb["site"], header_color="FF0000", hyperlink_col=3)
    wb.save(output_file)
    messagebox.showinfo("نجاح", f"تم إنشاء تقرير فودافون\nالملف:\n{output_file}")


# ================== أورانج ==================
def generate_orange_report():
    global current_df, current_file
    if current_df is None or current_file is None:
        messagebox.showinfo("مطلوب", "افتح ملف Excel أولاً")
        return

    try:
        df = pd.read_excel(current_file, engine="openpyxl", header=4)
    except Exception as e:
        messagebox.showerror("خطأ في قراءة الملف", str(e))
        return

    required_cols = [
        'TARGET_MSISDN','TARGET_IMEI','TARGET_IMSI','TARGET_IMEI_TYPE','EVENT_START_TIME',
        'CALL_DURATION','EVENT_DIRECTION','OTHER_MSISDN','OTHER_NAME','OTHER_ID',
        'OTHER_ID_TYPE','OTHER_ADDRESS','CELL_ADDRESS','CELL_LAT','CELL_LONG'
    ]
    for col in required_cols:
        if col not in df.columns:
            messagebox.showerror("خطأ", f"العمود {col} غير موجود في الملف")
            return

    numbers = df['OTHER_MSISDN'].astype(str)
    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number','Count']

    calls_df = freq.merge(
        df[['OTHER_MSISDN','OTHER_NAME','OTHER_ADDRESS','OTHER_ID']].drop_duplicates(subset='OTHER_MSISDN'),
        left_on='B Number', right_on='OTHER_MSISDN', how='left'
    )

    sms_count = df[df['EVENT_DIRECTION'].astype(str).str.strip()=="SMSMT"].groupby('OTHER_MSISDN').size().reset_index(name='SMS')
    calls_df = calls_df.merge(sms_count, left_on='B Number', right_on='OTHER_MSISDN', how='left')
    calls_df['SMS'] = calls_df['SMS'].fillna(0).astype(int)

    # ===== إضافة أعمدة First Call و Last Call =====
    df['EVENT_START_TIME'] = pd.to_datetime(df['EVENT_START_TIME'])
    call_dates = df.groupby('OTHER_MSISDN')['EVENT_START_TIME'].agg(First_Call='min', Last_Call='max').reset_index()
    calls_df = calls_df.merge(call_dates, left_on='B Number', right_on='OTHER_MSISDN', how='left')
    calls_df = calls_df.drop(columns=['OTHER_MSISDN'])

    # ترتيب الأعمدة مع الأعمدة الجديدة في الآخر
    calls_df = calls_df[['B Number','Count','OTHER_NAME','OTHER_ADDRESS','OTHER_ID','SMS','First_Call','Last_Call']]
    calls_df.columns = ['B Number','Count','B Full Name','B Address','B Number id','SMS','First_Call','Last_Call']
    calls_df['B Number'] = calls_df['B Number'].apply(str)
    calls_df['B Number id'] = calls_df['B Number id'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
    calls_df['Count'] = calls_df['Count'].astype(int)
    calls_df = calls_df.sort_values(by='Count', ascending=False)

    # ===== صفحة IMEI =====
    df['TARGET_IMEI'] = df['TARGET_IMEI'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
    imei_group = df.groupby('TARGET_IMEI').agg(
        Count=('TARGET_IMEI','count'),
        TARGET_IMEI_TYPE=('TARGET_IMEI_TYPE','first'),
        First_Use_Date=('EVENT_START_TIME','min'),
        Last_Use_Date=('EVENT_START_TIME','max'),
        First_Use_Address=('CELL_ADDRESS','first'),
        Last_Use_Address=('CELL_ADDRESS','last')
    ).reset_index()
    imei_group['Device Info'] = imei_group['TARGET_IMEI'].apply(lambda x: f'=HYPERLINK("https://www.imei.info/calc/?imei={x}","IMEI Info")')
    imei_group = imei_group[['TARGET_IMEI','Count','TARGET_IMEI_TYPE','Device Info','First_Use_Date','Last_Use_Date','First_Use_Address','Last_Use_Address']]
    imei_group.columns = ['IMEI','Count','TARGET_IMEI_TYPE','Device Info','First_Use_Date','Last_Use_Date','First_Use_Address','Last_Use_Address']
    imei_group['Count'] = imei_group['Count'].astype(int)
    imei_group = imei_group.sort_values(by='Count', ascending=False)

    # ===== صفحة SITE =====
    site_df = df.groupby('CELL_ADDRESS').agg(
        Count=('CELL_ADDRESS','count'),
        First_Use_Date=('EVENT_START_TIME','min'),
        Last_Use_Date=('EVENT_START_TIME','max'),
        LAT=('CELL_LAT','first'),
        LON=('CELL_LONG','first')
    ).reset_index()
    site_df['Map'] = site_df.apply(lambda row: f'=HYPERLINK("https://www.google.com/maps/search/?api=1&query={row["LAT"]},{row["LON"]}","Map")' 
                                   if pd.notna(row["LAT"]) and pd.notna(row["LON"]) else '', axis=1)
    site_df = site_df[['CELL_ADDRESS','Count','Map','First_Use_Date','Last_Use_Date']]
    site_df = site_df.sort_values(by='Count', ascending=False)

    # ===== حفظ التقرير =====
    output_file = os.path.join(os.path.dirname(current_file), "orange_report.xlsx")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        calls_df.to_excel(writer, sheet_name="calls", index=False)
        imei_group.to_excel(writer, sheet_name="imei", index=False)
        site_df.to_excel(writer, sheet_name="site", index=False)
        current_df.to_excel(writer, sheet_name="cheet", index=False)

    wb = load_workbook(output_file)
    format_sheet(wb["calls"], header_color="FF6600")
    format_sheet(wb["imei"], header_color="FF6600", hyperlink_col=4)
    format_sheet(wb["site"], header_color="FF6600", hyperlink_col=3)
    wb.save(output_file)
    messagebox.showinfo("نجاح", f"تم إنشاء تقرير أورانج\nالملف:\n{output_file}")

























