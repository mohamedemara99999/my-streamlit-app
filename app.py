import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
import io

# ---------- طريقة آمنة لقراءة اليوزر والباسورد من secrets ----------
DEFAULT_USERS = {"admin": "1234"}  # fallback محلي للاختبار

def get_users():
    try:
        creds = st.secrets["credentials"]
        return dict(creds)
    except Exception:
        return DEFAULT_USERS

USERS = get_users()

# ---------- نظام تسجيل دخول ----------
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

# ---------- هنا محتوى التطبيق بعد تسجيل الدخول ----------
st.sidebar.write(f"متصل كـ: {st.session_state.username}")
st.title("تطبيق Streamlit - داخل بعد تسجيل الدخول")
st.write("تحكم، اعرض بيانات، ارفع ملفات... أضف اللي تحتاجه هنا.")

uploaded = st.file_uploader("ارفع ملف Excel لاختبار التحليل", type=["xlsx","xls"])
if uploaded:
    df = pd.read_excel(uploaded, engine="openpyxl")

    st.write("Preview:")
    st.dataframe(df.head())

    # ======= تقرير اتصالات =========
    def generate_etisalat_report(df):
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
        df_final = df_final.sort_values(by='Count', ascending=False)
        return df_final

    # ======= تقرير فودافون =========
    def generate_vodafone_report(df):
        df['B Full Name'] = df['B_NUMBER_FIRST_NAME'].fillna('') + ' ' + df['B_NUMBER_LAST_NAME'].fillna('')
        numbers = df['B_NUMBER'].astype(str)
        freq = numbers.value_counts().reset_index()
        freq.columns = ['B Number','Count']
        df_final = freq.merge(
            df[['B_NUMBER','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS']].drop_duplicates(subset='B_NUMBER'),
            left_on='B Number', right_on='B_NUMBER', how='left'
        )
        df_final = df_final[['B Number','Count','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS']]
        df_final['Count'] = df_final['Count'].astype(int)
        df_final = df_final.sort_values(by='Count', ascending=False)
        return df_final

    st.subheader("تقارير الاتصالات")
    if st.button("توليد تقرير اتصالات"):
        etisalat_report = generate_etisalat_report(df)
        st.dataframe(etisalat_report.head())

        # زر التحميل
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            etisalat_report.to_excel(writer, index=False)
        buffer.seek(0)
        st.download_button(
            label="تحميل تقرير الاتصالات",
            data=buffer,
            file_name="etisalat_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.subheader("تقارير فودافون")
    if st.button("توليد تقرير فودافون"):
        vodafone_report = generate_vodafone_report(df)
        st.dataframe(vodafone_report.head())

        # زر التحميل
        buffer2 = io.BytesIO()
        with pd.ExcelWriter(buffer2, engine='openpyxl') as writer:
            vodafone_report.to_excel(writer, index=False)
        buffer2.seek(0)
        st.download_button(
            label="تحميل تقرير فودافون",
            data=buffer2,
            file_name="vodafone_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
