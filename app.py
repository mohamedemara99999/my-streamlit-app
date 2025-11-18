import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ================== كلمة المرور ==================
PASSWORD = "13579"

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# صفحة تسجيل الدخول
if not st.session_state.logged_in:
    st.title("تسجيل الدخول")
    password_input = st.text_input("ادخل كلمة المرور", type="password")
    if st.button("دخول"):
        if password_input == PASSWORD:
            st.session_state.logged_in = True
            st.experimental_rerun()
        else:
            st.error("كلمة المرور غير صحيحة")
else:
    st.title("📊 Excel Analyzer Tool - Streamlit")

    uploaded_file = st.file_uploader("اختر ملف Excel", type=["xlsx", "xls"])
    current_df = None
    current_file_name = ""

    if uploaded_file is not None:
        try:
            current_df = pd.read_excel(uploaded_file, engine="openpyxl")
            current_file_name = uploaded_file.name
            st.success(f"تم فتح الملف: {current_file_name} | عدد الصفوف: {len(current_df)}")
            st.dataframe(current_df)
        except Exception as e:
            st.error(f"خطأ في فتح الملف: {e}")

    # ================== دوال التنسيق ==================
    def format_excel_sheets(output, report_type):
        output.seek(0)
        wb = load_workbook(output)

        if report_type == "etisalat":
            header_color = "228B22"
            hyperlinks = {"calls": 8, "imei": 3, "site": 3}
        elif report_type == "vodafone":
            header_color = "FF0000"
            hyperlinks = {"calls": None, "imei": 3, "site": 3}
        elif report_type == "orange":
            header_color = "FF6600"
            hyperlinks = {"calls": None, "imei": 4, "site": 3}
        else:
            header_color = "000000"
            hyperlinks = {}

        for ws_name in wb.sheetnames:
            ws = wb[ws_name]
            # رؤوس الأعمدة
            for cell in ws[1]:
                cell.fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
                cell.font = Font(bold=True, color="FFFFFF", size=14)
                cell.alignment = Alignment(horizontal="center")
            # باقي الخلايا
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
                for cell in row:
                    cell.font = Font(size=12)
                    cell.alignment = Alignment(horizontal="left")
            # هايبرلينك
            if ws_name in hyperlinks and hyperlinks[ws_name]:
                col = hyperlinks[ws_name]
                for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
                    for cell in row:
                        if cell.value:
                            cell.font = Font(color="006400", size=12)

        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)
        return final_output

    # ================== تقرير اتصالات ==================
    def generate_etisalat_report(df):
        required_cols = [
            'Originating_Number','Terminating_Number','Network_Activity_Type_Name',
            'Call_Start_Date','B_Number_Full_Name','B_Number_Address',
            'B_Number_MU_Site_Address','B_Number_MU_Latitude','B_Number_MU_Longitude',
            'IMEI_Number','Site_Address','Latitude','Longitude'
        ]
        for col in required_cols:
            if col not in df.columns:
                st.error(f"العمود {col} غير موجود في الملف")
                return None

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
        df_final.columns = ['B Number','Count','B Full Name','B Address','B_NUMBER_SITE_ADDRESS','Latitude','Longitude']

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

        # ====== imei sheet ======
        imei_df = df.copy()
        imei_df['IMEI_Number'] = imei_df['IMEI_Number'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
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
        imei_summary = imei_summary[['IMEI','Count','Device Info','First_Use_Date','Last_Use_Date','First_Use_Address','Last_Use_Address']]
        imei_summary['Count'] = imei_summary['Count'].astype(int)
        imei_summary = imei_summary.sort_values(by='Count', ascending=False)

        # ====== site sheet ======
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

        # حفظ Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name="calls", index=False)
            imei_summary.to_excel(writer, sheet_name="imei", index=False)
            site_group.to_excel(writer, sheet_name="site", index=False)
            df.to_excel(writer, sheet_name="cheet", index=False)
        output.seek(0)

        final_output = format_excel_sheets(output, "etisalat")
        return final_output

    # ================== أزرار التحليل ==================
    if current_df is not None:
        st.subheader("توليد تقارير")
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("تقرير اتصالات"):
                output = generate_etisalat_report(current_df)
                if output:
                    st.download_button(
                        label="تحميل تقرير اتصالات",
                        data=output,
                        file_name="etisalat_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        with col2:
            if st.button("تقرير فودافون"):
                output = generate_vodafone_report(current_df)
                if output:
                    st.download_button(
                        label="تحميل تقرير فودافون",
                        data=output,
                        file_name="vodafone_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        with col3:
            if st.button("تقرير أورانج"):
                output = generate_orange_report(current_df)
                if output:
                    st.download_button(
                        label="تحميل تقرير أورانج",
                        data=output,
                        file_name="orange_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
