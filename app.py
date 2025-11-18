import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from io import BytesIO

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
    st.title("Excel Analyzer Tool - Streamlit")

    uploaded_file = st.file_uploader("اختر ملف Excel", type=["xlsx","xls"])
    current_df = None

    if uploaded_file is not None:
        try:
            current_df = pd.read_excel(uploaded_file, engine="openpyxl")
            st.success(f"تم فتح الملف: {uploaded_file.name}")
            st.dataframe(current_df)
        except Exception as e:
            st.error(f"خطأ في فتح الملف: {e}")

    # ================== دوال التنسيق ==================
    def format_excel_sheets(output, header_colors):
        output.seek(0)
        wb = load_workbook(output)
        link_font = Font(color="006400", underline="single")
        
        for ws in wb.worksheets:
            # رؤوس الأعمدة
            color = header_colors.get(ws.title, "FF0000")
            for cell in ws[1]:
                cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                cell.font = Font(bold=True, color="FFFFFF", size=14)
                cell.alignment = Alignment(horizontal="center")
            
            # الروابط
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    if isinstance(cell.value, str) and (cell.value.startswith("http") or cell.value.startswith("=HYPERLINK")):
                        cell.hyperlink = cell.value if not cell.value.startswith("=HYPERLINK") else None
                        # إذا الرابط خرائط ضع "Map"، إذا imei ضع "Check Info"
                        if "google.com/maps" in cell.value:
                            cell.value = "Map"
                        else:
                            cell.value = "Check Info"
                        cell.font = link_font

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
        df_final.columns = ['B Number','Count','B Full Name','B Address',
                            'B_NUMBER_SITE_ADDRESS','Latitude','Longitude']

        df_final['Map'] = df_final.apply(
            lambda row: f'https://www.google.com/maps/search/?api=1&query={row["Latitude"]},{row["Longitude"]}' 
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
            lambda x: f'https://www.imei.info/calc/?imei={x}' if x != '' else ''
        )
        imei_summary = imei_summary[['IMEI','Count','Device Info','First_Use_Date','Last_Use_Date',
                                     'First_Use_Address','Last_Use_Address']]
        imei_summary['Count'] = imei_summary['Count'].astype(int)
        imei_summary = imei_summary.sort_values(by='Count', ascending=False)

        site_df = df[['Site_Address','Latitude','Longitude','Call_Start_Date']].copy()
        site_group = site_df.groupby('Site_Address').agg(
            Count=('Site_Address','count'),
            Map=('Latitude', lambda x: f'https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"Longitude"]}'),
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
        output.seek(0)

        # تطبيق التنسيقات والهايبرلينك
        final_output = format_excel_sheets(output, header_colors={'calls':'228B22','imei':'228B22','site':'228B22'})
        return final_output

    # ================== تقرير فودافون ==================
    def generate_vodafone_report(df):
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

        df_final = freq.merge(
            df2[['B_NUMBER','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS']].drop_duplicates(subset='B_NUMBER'),
            left_on='B Number', right_on='B_NUMBER', how='left'
        )
        df_final = df_final.merge(sms_count, left_on='B Number', right_on='B_NUMBER', how='left')
        df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)
        df_final = df_final[['B Number','Count','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS','SMS']]
        df_final['Count'] = df_final['Count'].astype(int)
        df_final = df_final.sort_values(by='Count', ascending=False)

        # معالجة IMEI ليظهر أرقام صحيحة
        df2['IMEI'] = df2['IMEI'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
        df2['FULL_DATE'] = pd.to_datetime(df2['FULL_DATE'])
        imei_group = df2.groupby('IMEI').agg(
            Count=('IMEI','count'),
            Device_Info=('IMEI', lambda x: f'https://www.imei.info/calc/?imei={x.iloc[0]}' if x.iloc[0] != '' else ''),
            HANDSET_MANUFACTURER=('HANDSET_MANUFACTURER','first'),
            HANDSET_MARKETING_NAME=('HANDSET_MARKETING_NAME','first'),
            First_Use_Date=('FULL_DATE','min'),
            Last_Use_Date=('FULL_DATE','max')
        ).reset_index()

        first_last_addr = []
        for imei in imei_group['IMEI']:
            sub = df2[df2['IMEI']==imei].sort_values('FULL_DATE')
            first_addr = sub.iloc[0]['SITE_ADDRESS']
            last_addr = sub.iloc[-1]['SITE_ADDRESS']
            first_last_addr.append((first_addr,last_addr))
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
        output.seek(0)

        final_output = format_excel_sheets(output, header_colors={'calls':'FF0000','imei':'FF0000','site':'FF0000'})
        return final_output

    # ================== تقرير أورانج ==================
    def generate_orange_report(df):
        if df is None:
            st.error("افتح ملف أولاً")
            return None
        try:
            df_orange = pd.read_excel(uploaded_file, engine="openpyxl", header=4)
        except Exception as e:
            st.error(f"خطأ في قراءة ملف أورانج: {e}")
            return None

        required_cols = [
            'TARGET_MSISDN','TARGET_IMEI','TARGET_IMSI','TARGET_IMEI_TYPE','EVENT_START_TIME',
            'CALL_DURATION','EVENT_DIRECTION','OTHER_MSISDN','OTHER_NAME','OTHER_ID',
            'OTHER_ID_TYPE','OTHER_ADDRESS','CELL_ADDRESS','CELL_LAT','CELL_LONG'
        ]
        for col in required_cols:
            if col not in df_orange.columns:
                st.error(f"العمود {col} غير موجود في الملف")
                return None

        numbers = df_orange['OTHER_MSISDN'].astype(str)
        freq = numbers.value_counts().reset_index()
        freq.columns = ['B Number','Count']

        calls_df = freq.merge(
            df_orange[['OTHER_MSISDN','OTHER_NAME','OTHER_ADDRESS','OTHER_ID']].drop_duplicates(subset='OTHER_MSISDN'),
            left_on='B Number', right_on='OTHER_MSISDN', how='left'
        )

        sms_count = df_orange[df_orange['EVENT_DIRECTION'].astype(str).str.strip()=="SMSMT"].groupby('OTHER_MSISDN').size().reset_index(name='SMS')
        calls_df = calls_df.merge(sms_count, left_on='B Number', right_on='OTHER_MSISDN', how='left')
        calls_df['SMS'] = calls_df['SMS'].fillna(0).astype(int)

        calls_df = calls_df[['B Number','Count','OTHER_NAME','OTHER_ADDRESS','OTHER_ID','SMS']]
        calls_df.columns = ['B Number','Count','B Full Name','B Address','B Number id','SMS']
        calls_df['B Number'] = calls_df['B Number'].apply(str)
        calls_df['B Number id'] = calls_df['B Number id'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
        calls_df['Count'] = calls_df['Count'].astype(int)
        calls_df = calls_df.sort_values(by='Count', ascending=False)

        df_orange['TARGET_IMEI'] = df_orange['TARGET_IMEI'].apply(lambda x: str(int(x)) if pd.notna(x) else '')
        imei_group = df_orange.groupby('TARGET_IMEI').agg(
            Count=('TARGET_IMEI','count'),
            TARGET_IMEI_TYPE=('TARGET_IMEI_TYPE','first'),
            First_Use_Date=('EVENT_START_TIME','min'),
            Last_Use_Date=('EVENT_START_TIME','max'),
            First_Use_Address=('CELL_ADDRESS','first'),
            Last_Use_Address=('CELL_ADDRESS','last')
        ).reset_index()
        imei_group['Device Info'] = imei_group['TARGET_IMEI'].apply(lambda x: f'https://www.imei.info/calc/?imei={x}' if x != '' else '')
        imei_group = imei_group[['TARGET_IMEI','Count','TARGET_IMEI_TYPE','Device Info','First_Use_Date','Last_Use_Date','First_Use_Address','Last_Use_Address']]
        imei_group.columns = ['IMEI','Count','TARGET_IMEI_TYPE','Device Info','First_Use_Date','Last_Use_Date','First_Use_Address','Last_Use_Address']
        imei_group['Count'] = imei_group['Count'].astype(int)
        imei_group = imei_group.sort_values(by='Count', ascending=False)

        site_df = df_orange.groupby('CELL_ADDRESS').agg(
            Count=('CELL_ADDRESS','count'),
            First_Use_Date=('EVENT_START_TIME','min'),
            Last_Use_Date=('EVENT_START_TIME','max'),
            LAT=('CELL_LAT','first'),
            LON=('CELL_LONG','first')
        ).reset_index()
        site_df['Map'] = site_df.apply(lambda row: f'https://www.google.com/maps/search/?api=1&query={row["LAT"]},{row["LON"]}' 
                                       if pd.notna(row["LAT"]) and pd.notna(row["LON"]) else '', axis=1)
        site_df = site_df[['CELL_ADDRESS','Count','Map','First_Use_Date','Last_Use_Date']]
        site_df = site_df.sort_values(by='Count', ascending=False)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            calls_df.to_excel(writer, sheet_name="calls", index=False)
            imei_group.to_excel(writer, sheet_name="imei", index=False)
            site_df.to_excel(writer, sheet_name="site", index=False)
        output.seek(0)

        final_output = format_excel_sheets(output, header_colors={'calls':'FF6600','imei':'FF6600','site':'FF6600'})
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
