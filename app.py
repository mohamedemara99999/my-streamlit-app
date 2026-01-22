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
            if st.session_state.active_sessions.get(username):
                st.error("❌ هذا الحساب مستخدم حالياً")
            else:
                st.session_state.logged_in = True
                st.session_state.current_user = username
                st.session_state.active_sessions[username] = True
                st.experimental_rerun()
        else:
            st.error("❌ بيانات غير صحيحة")

    st.stop()

# ================== واجهة البرنامج ==================
st.sidebar.success(f"مرحباً {st.session_state.current_user} 👋")
if st.sidebar.button("تسجيل خروج"):
    st.session_state.active_sessions[st.session_state.current_user] = False
    st.session_state.logged_in = False
    st.session_state.current_user = None
    st.experimental_rerun()

st.title("📊 Excel Analyzer Tool")

selected_company = st.selectbox("اختر الشركة", ["etisalat", "vodafone", "orange"])
uploaded_file = st.file_uploader("اختر ملف Excel", type=["xlsx", "xls"])

current_df = None

if uploaded_file:
    try:
        if selected_company == "orange":
            current_df = pd.read_excel(uploaded_file, header=4, engine="openpyxl")
        else:
            current_df = pd.read_excel(uploaded_file, engine="openpyxl")

        current_df.columns = current_df.columns.str.strip()
        current_df = current_df.loc[:, ~current_df.columns.str.contains("^Unnamed")]
        current_df = current_df.dropna(how="all", axis=1)

        if selected_company == "etisalat" and "Originating_Number" not in current_df.columns:
            st.error("❌ ملف غير صالح لاتصالات")
        else:
            st.success("✔️ تم فتح الملف")
            st.dataframe(current_df)

    except Exception as e:
        st.error(f"خطأ في فتح الملف: {e}")

# ================== تنسيق Excel ==================
def format_excel_sheets(output, header_color="228B22", highlight_row=2):
    output.seek(0)
    wb = load_workbook(output)

    header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=13)
    link_font = Font(color="006400", underline="single")

    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        if ws.title == "calls":
            for cell in ws[highlight_row]:
                cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("http"):
                    url = cell.value
                    cell.value = "map" if "maps" in url else "check info"
                    cell.hyperlink = url
                    cell.font = link_font

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ================== تقرير اتصالات ==================
def generate_etisalat_report(df):
    df = df.copy()
    df['Originating_Number'] = df['Originating_Number'].astype(str)
    df['Terminating_Number'] = df['Terminating_Number'].astype(str)

    numbers = pd.concat([df['Originating_Number'], df['Terminating_Number']])
    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number', 'Count']
    df_final = freq.copy()

    b_data = {}
    for _, r in df.iterrows():
        for c in ['Originating_Number', 'Terminating_Number']:
            if r[c] not in b_data:
                b_data[r[c]] = {
                    "Name": r.get("B_Number_Full_Name", ""),
                    "Addr": r.get("B_Number_Address", ""),
                    "Site": r.get("B_Number_MU_Site_Address", ""),
                    "Lat": r.get("B_Number_MU_Latitude", ""),
                    "Lon": r.get("B_Number_MU_Longitude", "")
                }

    df_final["B Full Name"] = df_final["B Number"].map(lambda x: b_data.get(x, {}).get("Name", ""))
    df_final["B Address"] = df_final["B Number"].map(lambda x: b_data.get(x, {}).get("Addr", ""))
    df_final["B_NUMBER_SITE_ADDRESS"] = df_final["B Number"].map(lambda x: b_data.get(x, {}).get("Site", ""))
    df_final["Latitude"] = df_final["B Number"].map(lambda x: b_data.get(x, {}).get("Lat", ""))
    df_final["Longitude"] = df_final["B Number"].map(lambda x: b_data.get(x, {}).get("Lon", ""))

    df_final["Map"] = df_final.apply(
        lambda r: f"https://www.google.com/maps/search/?api=1&query={r['Latitude']},{r['Longitude']}"
        if r["Latitude"] != "" else "", axis=1
    )

    temp = df.copy()
    temp["Call_Start_Date"] = pd.to_datetime(temp["Call_Start_Date"])

    calls = pd.concat([
        temp[['Originating_Number','Call_Start_Date']].rename(columns={'Originating_Number':'B Number'}),
        temp[['Terminating_Number','Call_Start_Date']].rename(columns={'Terminating_Number':'B Number'})
    ])

    first_last = calls.groupby("B Number").agg(
        First_Call=('Call_Start_Date','min'),
        Last_Call=('Call_Start_Date','max')
    ).reset_index()

    df_final = df_final.merge(first_last, on="B Number", how="left")
    df_final = df_final.sort_values("Count", ascending=False)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(writer, sheet_name="calls", index=False)

    output.seek(0)
    return output

# ================== زر التحليل ==================
if selected_company == "etisalat" and current_df is not None:
    if st.button("🔍 تحليل تقرير اتصالات"):
        with st.spinner("جارٍ التحليل..."):
            out = generate_etisalat_report(current_df)
            final_out = format_excel_sheets(out)

        st.success("✅ التقرير جاهز")

        st.download_button(
            "⬇️ تحميل التقرير",
            data=final_out,
            file_name="etisalat_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

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


























