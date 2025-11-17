import streamlit as st
import pandas as pd
from io import BytesIO

# ================== بيانات المستخدم ==================
DEFAULT_USERS = {"admin": "1234"}

def get_users():
    try:
        creds = st.secrets["credentials"]
        return dict(creds)
    except Exception:
        return DEFAULT_USERS

USERS = get_users()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""

if not st.session_state.logged_in:
    st.title("تسجيل الدخول")
    username = st.text_input("اسم المستخدم")
    password = st.text_input("كلمة المرور", type="password")
    if st.button("دخول"):
        if username in USERS and password == USERS[username]:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.success(f"مرحبا {username}!")
            st.experimental_rerun()
        else:
            st.error("اسم المستخدم أو كلمة المرور غير صحيحة")
    st.stop()

st.sidebar.write(f"متصل كـ: {st.session_state.username}")
st.title("تطبيق Streamlit - داخل بعد تسجيل الدخول")
st.write("تحكم، اعرض بيانات، ارفع ملفات...")

# ================== رفع الملف ==================
uploaded_file = st.file_uploader("ارفع ملف Excel (CSV أو XLSX)", type=["csv", "xlsx"])
if uploaded_file is not None:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    st.write("Preview:")
    st.dataframe(df.head())

    # ================== دوال التقارير ==================
    def generate_etisalat_report(df):
        # ====== calls sheet ======
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
            lambda row: f'https://www.google.com/maps/search/?api=1&query={row["Latitude"]},{row["Longitude"]}'
            if pd.notna(row['Latitude']) and pd.notna(row['Longitude']) else '', axis=1
        )

        # ====== imei sheet ======
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
            lambda x: f'https://www.imei.info/calc/?imei={x}'
        )
        imei_summary = imei_summary[['IMEI','Count','Device Info','First_Use_Date','Last_Use_Date',
                                     'First_Use_Address','Last_Use_Address']]

        # ====== site sheet ======
        site_df = df[['Site_Address','Latitude','Longitude','Call_Start_Date']].copy()
        site_group = site_df.groupby('Site_Address').agg(
            Count=('Site_Address','count'),
            Map=('Latitude', lambda x: f'https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"Longitude"]}'),
            First_Use_Date=('Call_Start_Date','min'),
            Last_Use_Date=('Call_Start_Date','max')
        ).reset_index()
        site_group = site_group[['Site_Address','Count','Map','First_Use_Date','Last_Use_Date']]

        # حفظ في BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name="calls", index=False)
            imei_summary.to_excel(writer, sheet_name="imei", index=False)
            site_group.to_excel(writer, sheet_name="site", index=False)
            writer.save()
        output.seek(0)
        return output

    def generate_vodafone_report(df):
        df['B Full Name'] = df['B_NUMBER_FIRST_NAME'].fillna('') + ' ' + df['B_NUMBER_LAST_NAME'].fillna('')
        numbers = df['B_NUMBER'].astype(str)
        freq = numbers.value_counts().reset_index()
        freq.columns = ['B Number','Count']

        sms_count = df[df['SERVICE'].astype(str).str.strip()=="Short message MO/PP"].groupby('B_NUMBER').size().reset_index(name='SMS')

        df_final = freq.merge(
            df[['B_NUMBER','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS']].drop_duplicates(subset='B_NUMBER'),
            left_on='B Number', right_on='B_NUMBER', how='left'
        )
        df_final = df_final.merge(sms_count, left_on='B Number', right_on='B_NUMBER', how='left')
        df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)
        df_final = df_final[['B Number','Count','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS','SMS']]
        df_final['Count'] = df_final['Count'].astype(int)

        # ====== imei sheet ======
        df['FULL_DATE'] = pd.to_datetime(df['FULL_DATE'])
        imei_group = df.groupby('IMEI').agg(
            Count=('IMEI','count'),
            Device_Info=('IMEI', lambda x: f'https://www.imei.info/calc/?imei={x.iloc[0]}'),
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

        # ====== site sheet ======
        site_df = df[['SITE_ADDRESS','LATITUDE','LONGITUDE','FULL_DATE']].copy()
        site_group = site_df.groupby('SITE_ADDRESS').agg(
            Count=('SITE_ADDRESS','count'),
            Map=('LATITUDE', lambda x: f'https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"LONGITUDE"]}'),
            First_Use_Date=('FULL_DATE','min'),
            Last_Use_Date=('FULL_DATE','max')
        ).reset_index()
        site_group = site_group[['SITE_ADDRESS','Count','Map','First_Use_Date','Last_Use_Date']]

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name="calls", index=False)
            imei_group.to_excel(writer, sheet_name="imei", index=False)
            site_group.to_excel(writer, sheet_name="site", index=False)
            writer.save()
        output.seek(0)
        return output

    # ================== أزرار التحميل ==================
    st.download_button(
        label="تحميل تقرير اتصالات",
        data=generate_etisalat_report(df),
        file_name="etisalat_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        label="تحميل تقرير فودافون",
        data=generate_vodafone_report(df),
        file_name="vodafone_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
