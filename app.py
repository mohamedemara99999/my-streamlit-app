import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

st.set_page_config(page_title="Excel Analyzer Tool", layout="wide")

# ================== تسجيل دخول بسيط ==================
DEFAULT_USERS = {"admin": "1234"}  # يمكن تعديل المستخدمين
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""

if not st.session_state.logged_in:
    st.title("تسجيل الدخول")
    username = st.text_input("اسم المستخدم")
    password = st.text_input("كلمة المرور", type="password")
    if st.button("دخول"):
        if username in DEFAULT_USERS and password == DEFAULT_USERS[username]:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.experimental_rerun()
        else:
            st.error("اسم المستخدم أو كلمة المرور غير صحيحة")
    st.stop()

st.sidebar.write(f"متصل كـ: {st.session_state.username}")
st.title("تطبيق Excel Analyzer - داخل بعد تسجيل الدخول")
st.write("تحكم، اعرض بيانات، ارفع ملفات...")

# ================== رفع ملف ==================
uploaded_file = st.file_uploader("ارفع ملف Excel لاختبار التحليل", type=["xlsx", "xls"])
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.success("تم تحميل الملف بنجاح!")
        st.dataframe(df.head())
    except Exception as e:
        st.error(f"خطأ في قراءة الملف: {e}")
        st.stop()
else:
    st.info("يرجى رفع ملف Excel للبدء بالتحليل.")
    st.stop()

# ================== دوال تحليل ==================
def format_sheet(ws, hyperlink_col=None):
    header_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=14)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=ws.max_column, max_row=ws.max_row):
        for cell in row:
            cell.font = Font(size=12)
            cell.alignment = Alignment(horizontal="left")
    if hyperlink_col:
        for row in ws.iter_rows(min_row=2, min_col=hyperlink_col, max_col=hyperlink_col):
            for cell in row:
                if cell.value:
                    cell.font = Font(color="006400", size=12)

def save_excel(output_file, sheets_dict):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for name, df_sheet in sheets_dict.items():
            df_sheet.to_excel(writer, sheet_name=name, index=False)
    wb = load_workbook(output_file)
    for name, df_sheet in sheets_dict.items():
        if name=="calls":
            format_sheet(wb[name], hyperlink_col=8 if "Map" in df_sheet.columns else None)
        else:
            format_sheet(wb[name], hyperlink_col=3 if "Map" in df_sheet.columns or "Device Info" in df_sheet.columns else None)
    wb.save(output_file)

# ================== تقرير اتصالات ==================
def generate_etisalat_report(df, uploaded_file):
    required_cols = [
        'Originating_Number', 'Terminating_Number', 'Network_Activity_Type_Name',
        'Call_Start_Date','B_Number_Full_Name', 'B_Number_Address',
        'B_Number_MU_Site_Address','B_Number_MU_Latitude','B_Number_MU_Longitude',
        'IMEI_Number','Site_Address','Latitude','Longitude'
    ]
    for col in required_cols:
        if col not in df.columns:
            st.error(f"العمود {col} غير موجود في الملف")
            return

    numbers = pd.concat([df['Originating_Number'], df['Terminating_Number']]).astype(str)
    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number','Count']

    df_final = freq.merge(
        df[['Originating_Number','B_Number_Full_Name','B_Number_Address',
            'B_Number_MU_Site_Address','B_Number_MU_Latitude','B_Number_MU_Longitude']].drop_duplicates(subset='Originating_Number'),
        left_on='B Number', right_on='Originating_Number', how='left'
    )
    df_final = df_final[['B Number','Count','B_Number_Full_Name','B_Number_Address',
                         'B_Number_MU_Site_Address','B_Number_MU_Latitude','B_Number_MU_Longitude']]
    df_final.columns = ['B Number','Count','B Full Name','B Address',
                        'B_NUMBER_SITE_ADDRESS','Latitude','Longitude']
    df_final['Map'] = df_final.apply(
        lambda row: f'=HYPERLINK("https://www.google.com/maps/search/?api=1&query={row["Latitude"]},{row["Longitude"]}","Map")'
        if pd.notna(row['Latitude']) and pd.notna(row['Longitude']) else '', axis=1
    )

    temp_df = df.copy()
    temp_df['activity_clean'] = temp_df['Network_Activity_Type_Name'].astype(str).str.strip()
    activity_stats = temp_df.groupby('Originating_Number').agg(
        SMS=('activity_clean', lambda x: (x=="SMS").sum())
    ).reset_index()
    df_final = df_final.merge(activity_stats, left_on='B Number', right_on='Originating_Number', how='left')
    df_final.drop(columns=['Originating_Number'], inplace=True)
    df_final['Count'] = df_final['Count'].astype(int)
    df_final = df_final.sort_values(by='Count', ascending=False)

    # imei sheet
    imei_df = df.copy()
    imei_df['IMEI_Number'] = imei_df['IMEI_Number'].astype(str)
    imei_summary = imei_df.groupby('IMEI_Number').agg(
        Count=('IMEI_Number','count'),
        First_Use_Date=('Call_Start_Date','min'),
        Last_Use_Date=('Call_Start_Date','max'),
        First_Use_Address=('Site_Address','first'),
        Last_Use_Address=('Site_Address','last')
    ).reset_index()
    imei_summary.rename(columns={'IMEI_Number':'IMEI'}, inplace=True)
    imei_summary['Device Info'] = imei_summary['IMEI'].apply(
        lambda x: f'=HYPERLINK("https://www.imei.info/calc/?imei={x}","IMEI Info")'
    )
    imei_summary = imei_summary[['IMEI','Count','Device Info','First_Use_Date','Last_Use_Date',
                                 'First_Use_Address','Last_Use_Address']]
    imei_summary['Count'] = imei_summary['Count'].astype(int)
    imei_summary = imei_summary.sort_values(by='Count', ascending=False)

    # site sheet
    site_df = df[['Site_Address','Latitude','Longitude','Call_Start_Date']].copy()
    site_group = site_df.groupby('Site_Address').agg(
        Count=('Site_Address','count'),
        Map=('Latitude', lambda x: f'=HYPERLINK("https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"Longitude"]}","Map")'),
        First_Use_Date=('Call_Start_Date','min'),
        Last_Use_Date=('Call_Start_Date','max')
    ).reset_index()
    site_group['Count'] = site_group['Count'].astype(int)
    site_group = site_group.sort_values(by='Count', ascending=False)
    site_group = site_group[['Site_Address','Count','Map','First_Use_Date','Last_Use_Date']]

    output_file = os.path.join(os.getcwd(), "etisalat_report.xlsx")
    save_excel(output_file, {"calls": df_final, "imei": imei_summary, "site": site_group})
    st.success(f"تم إنشاء تقرير اتصالات بنجاح!\n[تحميل التقرير](file://{output_file})")

# ================== تقرير فودافون ==================
def generate_vodafone_report(df, uploaded_file):
    required_cols = [
        'B_NUMBER','B_NUMBER_FIRST_NAME','B_NUMBER_LAST_NAME','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS',
        'IMEI','HANDSET_MANUFACTURER','HANDSET_MARKETING_NAME','FULL_DATE','SITE_ADDRESS','LATITUDE','LONGITUDE','SERVICE'
    ]
    for col in required_cols:
        if col not in df.columns:
            st.error(f"العمود {col} غير موجود في الملف")
            return

    df_copy = df.copy()
    df_copy['B Full Name'] = df_copy['B_NUMBER_FIRST_NAME'].fillna('') + ' ' + df_copy['B_NUMBER_LAST_NAME'].fillna('')

    numbers = df_copy['B_NUMBER'].astype(str)
    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number','Count']

    sms_count = df_copy[df_copy['SERVICE'].astype(str).str.strip()=="Short message MO/PP"].groupby('B_NUMBER').size().reset_index(name='SMS')

    df_final = freq.merge(
        df_copy[['B_NUMBER','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS']].drop_duplicates(subset='B_NUMBER'),
        left_on='B Number', right_on='B_NUMBER', how='left'
    )
    df_final = df_final.merge(sms_count, left_on='B Number', right_on='B_NUMBER', how='left')
    df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)
    df_final = df_final[['B Number','Count','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS','SMS']]
    df_final['Count'] = df_final['Count'].astype(int)
    df_final = df_final.sort_values(by='Count', ascending=False)

    df_copy['FULL_DATE'] = pd.to_datetime(df_copy['FULL_DATE'])
    imei_group = df_copy.groupby('IMEI').agg(
        Count=('IMEI','count'),
        Device_Info=('IMEI', lambda x: f'=HYPERLINK("https://www.imei.info/calc/?imei={x.iloc[0]}","IMEI Info")'),
        HANDSET_MANUFACTURER=('HANDSET_MANUFACTURER','first'),
        HANDSET_MARKETING_NAME=('HANDSET_MARKETING_NAME','first'),
        First_Use_Date=('FULL_DATE','min'),
        Last_Use_Date=('FULL_DATE','max')
    ).reset_index()

    first_last_addr = []
    for imei in imei_group['IMEI']:
        sub = df_copy[df_copy['IMEI']==imei].sort_values('FULL_DATE')
        first_addr = sub.iloc[0]['SITE_ADDRESS']
        last_addr = sub.iloc[-1]['SITE_ADDRESS']
        first_last_addr.append((first_addr,last_addr))
    imei_group['First_Use_Address'] = [x[0] for x in first_last_addr]
    imei_group['Last_Use_Address'] = [x[1] for x in first_last_addr]

    imei_group = imei_group[['IMEI','Count','Device_Info','HANDSET_MANUFACTURER','HANDSET_MARKETING_NAME',
                             'First_Use_Date','Last_Use_Date','First_Use_Address','Last_Use_Address']]
    imei_group['Count'] = imei_group['Count'].astype(int)
    imei_group = imei_group.sort_values(by='Count', ascending=False)

    site_df = df_copy[['SITE_ADDRESS','LATITUDE','LONGITUDE','FULL_DATE']].copy()
    site_group = site_df.groupby('SITE_ADDRESS').agg(
        Count=('SITE_ADDRESS','count'),
        Map=('LATITUDE', lambda x: f'=HYPERLINK("https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"LONGITUDE"]}","Map")'),
        First_Use_Date=('FULL_DATE','min'),
        Last_Use_Date=('FULL_DATE','max')
    ).reset_index()
    site_group['Count'] = site_group['Count'].astype(int)
    site_group = site_group.sort_values(by='Count', ascending=False)
    site_group = site_group[['SITE_ADDRESS','Count','Map','First_Use_Date','Last_Use_Date']]

    output_file = os.path.join(os.getcwd(), "vodafone_report.xlsx")
    save_excel(output_file, {"calls": df_final, "imei": imei_group, "site": site_group})
    st.success(f"تم إنشاء تقرير فودافون بنجاح!\n[تحميل التقرير](file://{output_file})")

# ================== أزرار التحليل ==================
col1, col2 = st.columns(2)
if col1.button("اتصالات"):
    generate_etisalat_report(df, uploaded_file)
if col2.button("فودافون"):
    generate_vodafone_report(df, uploaded_file)
