import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from io import BytesIO

# ================== كلمة المرور ==================
PASSWORD = "13579"

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# ===== صفحة تسجيل الدخول =====
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
    uploaded_file = st.file_uploader("اختر ملف Excel", type=["xlsx","xls"])
    current_df = None

    if uploaded_file is not None:
        try:
            # لأورانج يبدأ من الصف 5
            header_row = 4 if uploaded_file.name.lower().startswith("orange") else 0
            current_df = pd.read_excel(uploaded_file, engine="openpyxl", header=header_row)
            st.success(f"تم فتح الملف: {uploaded_file.name}")
            st.dataframe(current_df)
        except Exception as e:
            st.error(f"خطأ في فتح الملف: {e}")

    # ===== دالة تنسيق الـ Excel =====
    def format_excel_sheets(output, sheet_hyperlink_cols=None, header_colors=None):
        output.seek(0)
        wb = load_workbook(output)
        sheet_hyperlink_cols = sheet_hyperlink_cols or {}
        header_colors = header_colors or {}

        for ws in wb.worksheets:
            color = header_colors.get(ws.title, "228B22")  # اللون الافتراضي أخضر
            header_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=14)
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    cell.font = Font(size=12)
                    cell.alignment = Alignment(horizontal="left")
            if ws.title in sheet_hyperlink_cols:
                col_idx = sheet_hyperlink_cols[ws.title]
                for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        if cell.value:
                            cell.font = Font(color="006400", underline="single", size=12)
        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)
        return final_output

    # ================== دوال التقارير ==================
    def generate_etisalat_report(df):
        df.columns = [str(c) for c in df.columns]
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

        # ===== calls sheet =====
        numbers = pd.concat([df['Originating_Number'], df['Terminating_Number']]).astype(str)
        freq = numbers.value_counts().reset_index()
        freq.columns = ['B Number','Count']
        df_final = freq.merge(
            df[['Originating_Number','B_Number_Full_Name','B_Number_Address','B_Number_MU_Site_Address','B_Number_MU_Latitude','B_Number_MU_Longitude']].drop_duplicates('Originating_Number'),
            left_on='B Number', right_on='Originating_Number', how='left'
        )
        df_final = df_final[['B Number','Count','B_Number_Full_Name','B_Number_Address','B_Number_MU_Site_Address','B_Number_MU_Latitude','B_Number_MU_Longitude']]
        df_final.columns = ['B Number','Count','B Full Name','B Address','B_NUMBER_SITE_ADDRESS','Latitude','Longitude']
        df_final['Map'] = df_final.apply(lambda row: f'https://www.google.com/maps/search/?api=1&query={row["Latitude"]},{row["Longitude"]}' if pd.notna(row['Latitude']) else '', axis=1)

        # ===== imei sheet =====
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
        imei_summary['Device Info'] = imei_summary['IMEI'].apply(lambda x: f'https://www.imei.info/calc/?imei={x}')
        imei_summary['Count'] = imei_summary['Count'].astype(int)
        imei_summary = imei_summary.sort_values('Count', ascending=False)

        # ===== site sheet =====
        site_df = df[['Site_Address','Latitude','Longitude','Call_Start_Date']].copy()
        site_group = site_df.groupby('Site_Address').agg(
            Count=('Site_Address','count'),
            Map=('Latitude', lambda x: f'https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"Longitude"]}'),
            First_Use_Date=('Call_Start_Date','min'),
            Last_Use_Date=('Call_Start_Date','max')
        ).reset_index()
        site_group['Count'] = site_group['Count'].astype(int)
        site_group = site_group.sort_values('Count', ascending=False)
        site_group = site_group[['Site_Address','Count','Map','First_Use_Date','Last_Use_Date']]

        # ===== cheet sheet =====
        cheet_df = df.copy()

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name="calls", index=False)
            imei_summary.to_excel(writer, sheet_name="imei", index=False)
            site_group.to_excel(writer, sheet_name="site", index=False)
            cheet_df.to_excel(writer, sheet_name="cheet", index=False)
        output.seek(0)
        return format_excel_sheets(output, sheet_hyperlink_cols={'calls':8,'imei':3,'site':3}, header_colors={'calls':'228B22','imei':'228B22','site':'228B22'})

    def generate_vodafone_report(df):
        df.columns = [str(c) for c in df.columns]
        required_cols = [
            'B_NUMBER','B_NUMBER_FIRST_NAME','B_NUMBER_LAST_NAME','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS',
            'IMEI','HANDSET_MANUFACTURER','HANDSET_MARKETING_NAME','FULL_DATE','SITE_ADDRESS','LATITUDE','LONGITUDE','SERVICE'
        ]
        for col in required_cols:
            if col not in df.columns:
                st.error(f"العمود {col} غير موجود في الملف")
                return None

        df2 = df.copy()
        df2['B Full Name'] = df2['B_NUMBER_FIRST_NAME'].fillna('') + ' ' + df2['B_NUMBER_LAST_NAME'].fillna('')
        numbers = df2['B_NUMBER'].astype(str)
        freq = numbers.value_counts().reset_index()
        freq.columns = ['B Number','Count']
        sms_count = df2[df2['SERVICE'].astype(str).str.strip()=="Short message MO/PP"].groupby('B_NUMBER').size().reset_index(name='SMS')
        df_final = freq.merge(df2[['B_NUMBER','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS']].drop_duplicates('B_NUMBER'), left_on='B Number', right_on='B_NUMBER', how='left')
        df_final = df_final.merge(sms_count, left_on='B Number', right_on='B_NUMBER', how='left')
        df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)
        df_final = df_final[['B Number','Count','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS','SMS']]
        df_final['Count'] = df_final['Count'].astype(int)
        df_final = df_final.sort_values('Count', ascending=False)

        df2['FULL_DATE'] = pd.to_datetime(df2['FULL_DATE'])
        df2['IMEI'] = df2['IMEI'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
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
            sub = df2[df2['IMEI']==imei].sort_values('FULL_DATE')
            first_last_addr.append((sub.iloc[0]['SITE_ADDRESS'], sub.iloc[-1]['SITE_ADDRESS']))
        imei_group['First_Use_Address'] = [x[0] for x in first_last_addr]
        imei_group['Last_Use_Address'] = [x[1] for x in first_last_addr]
        imei_group['Count'] = imei_group['Count'].astype(int)
        imei_group = imei_group.sort_values('Count', ascending=False)

        site_df = df2[['SITE_ADDRESS','LATITUDE','LONGITUDE','FULL_DATE']].copy()
        site_group = site_df.groupby('SITE_ADDRESS').agg(
            Count=('SITE_ADDRESS','count'),
            Map=('LATITUDE', lambda x: f'https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"LONGITUDE"]}'),
            First_Use_Date=('FULL_DATE','min'),
            Last_Use_Date=('FULL_DATE','max')
        ).reset_index()
        site_group['Count'] = site_group['Count'].astype(int)
        site_group = site_group.sort_values('Count', ascending=False)
        site_group = site_group[['SITE_ADDRESS','Count','Map','First_Use_Date','Last_Use_Date']]

        cheet_df = df2.copy()

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name="calls", index=False)
            imei_group.to_excel(writer, sheet_name="imei", index=False)
            site_group.to_excel(writer, sheet_name="site", index=False)
            cheet_df.to_excel(writer, sheet_name="cheet", index=False)
        output.seek(0)
        return format_excel_sheets(output, sheet_hyperlink_cols={'calls':6,'imei':3,'site':3}, header_colors={'calls':'FF0000','imei':'FF0000','site':'FF0000'})

    def generate_orange_report(df):
        df.columns = [str(c) for c in df.columns]
        required_cols = [
            'TARGET_MSISDN','TARGET_IMEI','TARGET_IMSI','TARGET_IMEI_TYPE','EVENT_START_TIME',
            'CALL_DURATION','EVENT_DIRECTION','OTHER_MSISDN','OTHER_NAME','OTHER_ID',
            'OTHER_ID_TYPE','OTHER_ADDRESS','CELL_ADDRESS','CELL_LAT','CELL_LONG'
        ]
        for col in required_cols:
            if col not in df.columns:
                st.error(f"العمود {col} غير موجود في الملف")
                return None

        numbers = df['OTHER_MSISDN'].astype(str)
        freq = numbers.value_counts().reset_index()
        freq.columns = ['B Number','Count']
        calls_df = freq.merge(
            df[['OTHER_MSISDN','OTHER_NAME','OTHER_ADDRESS','OTHER_ID']].drop_duplicates('OTHER_MSISDN'),
            left_on='B Number', right_on='OTHER_MSISDN', how='left'
        )
        sms_count = df[df['EVENT_DIRECTION'].astype(str).str.strip()=="SMSMT"].groupby('OTHER_MSISDN').size().reset_index(name='SMS')
        calls_df = calls_df.merge(sms_count, left_on='B Number', right_on='OTHER_MSISDN', how='left')
        calls_df['SMS'] = calls_df['SMS'].fillna(0).astype(int)
        calls_df = calls_df[['B Number','Count','OTHER_NAME','OTHER_ADDRESS','OTHER_ID','SMS']]
        calls_df.columns = ['B Number','Count','B Full Name','B Address','B Number id','SMS']
        calls_df['Count'] = calls_df['Count'].astype(int)
        calls_df = calls_df.sort_values('Count', ascending=False)

        df['TARGET_IMEI'] = df['TARGET_IMEI'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
        imei_group = df.groupby('TARGET_IMEI').agg(
            Count=('TARGET_IMEI','count'),
            TARGET_IMEI_TYPE=('TARGET_IMEI_TYPE','first'),
            First_Use_Date=('EVENT_START_TIME','min'),
            Last_Use_Date=('EVENT_START_TIME','max'),
            First_Use_Address=('CELL_ADDRESS','first'),
            Last_Use_Address=('CELL_ADDRESS','last')
        ).reset_index()
        imei_group['Device Info'] = imei_group['TARGET_IMEI'].apply(lambda x: f'https://www.imei.info/calc/?imei={x}')
        imei_group['Count'] = imei_group['Count'].astype(int)
        imei_group = imei_group.sort_values('Count', ascending=False)

        site_df = df.groupby('CELL_ADDRESS').agg(
            Count=('CELL_ADDRESS','count'),
            First_Use_Date=('EVENT_START_TIME','min'),
            Last_Use_Date=('EVENT_START_TIME','max'),
            LAT=('CELL_LAT','first'),
            LON=('CELL_LONG','first')
        ).reset_index()
        site_df['Map'] = site_df.apply(lambda row: f'https://www.google.com/maps/search/?api=1&query={row["LAT"]},{row["LON"]}' if pd.notna(row["LAT"]) else '', axis=1)
        site_df = site_df[['CELL_ADDRESS','Count','Map','First_Use_Date','Last_Use_Date']]
        site_df = site_df.sort_values('Count', ascending=False)

        cheet_df = df.copy()

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            calls_df.to_excel(writer, sheet_name="calls", index=False)
            imei_group.to_excel(writer, sheet_name="imei", index=False)
            site_df.to_excel(writer, sheet_name="site", index=False)
            cheet_df.to_excel(writer, sheet_name="cheet", index=False)
        output.seek(0)
        return format_excel_sheets(output, sheet_hyperlink_cols={'calls':6,'imei':4,'site':3}, header_colors={'calls':'FF6600','imei':'FF6600','site':'FF6600'})

    # ================= أزرار التحليل ==================
    if current_df is not None:
        st.subheader("توليد تقارير")
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("تقرير اتصالات"):
                output = generate_etisalat_report(current_df)
                if output:
                    st.download_button("تحميل تقرير اتصالات", output, "etisalat_report.xlsx")
        with col2:
            if st.button("تقرير فودافون"):
                output = generate_vodafone_report(current_df)
                if output:
                    st.download_button("تحميل تقرير فودافون", output, "vodafone_report.xlsx")
        with col3:
            if st.button("تقرير اورانج"):
                output = generate_orange_report(current_df)
                if output:
                    st.download_button("تحميل تقرير اورانج", output, "orange_report.xlsx")
