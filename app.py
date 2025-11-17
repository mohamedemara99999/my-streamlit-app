import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ================== وظائف المساعد ==================
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

def save_to_excel_sheets(sheets_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.book = load_workbook(output)
    for sheet_name in sheets_dict.keys():
        ws = writer.book[sheet_name]
        # hyperlink_col guess: last column with hyperlinks
        format_sheet(ws, hyperlink_col=None)
    writer.book.save(output)
    output.seek(0)
    return output

# ================== واجهة Streamlit ==================
st.title("تطبيق تحليل Excel - Streamlit")

uploaded_file = st.file_uploader("ارفع ملف Excel لاختبار التحليل", type=["xlsx","xls"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    st.subheader("Preview:")
    st.dataframe(df.head())
    st.subheader("معلومات عامة:")
    st.write(df.describe(include='all'))

    col1, col2 = st.columns(2)

    with col1:
        if st.button("تقرير اتصالات"):
            # ========== نسخ الكود الأصلي لتقرير اتصالات ==========
            current_df = df.copy()
            numbers = pd.concat([current_df['Originating_Number'], current_df['Terminating_Number']]).astype(str)
            freq = numbers.value_counts().reset_index()
            freq.columns = ['B Number','Count']

            df_final = freq.merge(
                current_df[['Originating_Number','B_Number_Full_Name','B_Number_Address',
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

            temp_df = current_df.copy()
            temp_df['activity_clean'] = temp_df['Network_Activity_Type_Name'].astype(str).str.strip()
            activity_stats = temp_df.groupby('Originating_Number').agg(
                SMS=('activity_clean', lambda x: (x=="SMS").sum())
            ).reset_index()
            df_final = df_final.merge(activity_stats, left_on='B Number', right_on='Originating_Number', how='left')
            df_final.drop(columns=['Originating_Number'], inplace=True)
            df_final['Count'] = df_final['Count'].astype(int)
            df_final = df_final.sort_values(by='Count', ascending=False)

            # ====== imei sheet ======
            imei_df = current_df.copy()
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
            imei_summary['Count'] = imei_summary['Count'].astype(int)
            imei_summary = imei_summary.sort_values(by='Count', ascending=False)

            # ====== site sheet ======
            site_df = current_df[['Site_Address','Latitude','Longitude','Call_Start_Date']].copy()
            site_group = site_df.groupby('Site_Address').agg(
                Count=('Site_Address','count'),
                Map=('Latitude', lambda x: f'https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"Longitude"]}'),
                First_Use_Date=('Call_Start_Date','min'),
                Last_Use_Date=('Call_Start_Date','max')
            ).reset_index()
            site_group['Count'] = site_group['Count'].astype(int)
            site_group = site_group.sort_values(by='Count', ascending=False)
            site_group = site_group[['Site_Address','Count','Map','First_Use_Date','Last_Use_Date']]

            # ========== حفظ الملفات في ذاكرة للتحميل ==========
            sheets = {"calls": df_final, "imei": imei_summary, "site": site_group}
            output = save_to_excel_sheets(sheets)
            st.download_button("تحميل تقرير اتصالات", data=output, file_name="etisalat_report.xlsx")

    with col2:
        if st.button("تقرير فودافون"):
            current_df = df.copy()
            current_df['B Full Name'] = current_df['B_NUMBER_FIRST_NAME'].fillna('') + ' ' + current_df['B_NUMBER_LAST_NAME'].fillna('')
            numbers = current_df['B_NUMBER'].astype(str)
            freq = numbers.value_counts().reset_index()
            freq.columns = ['B Number','Count']
            sms_count = current_df[current_df['SERVICE'].astype(str).str.strip()=="Short message MO/PP"].groupby('B_NUMBER').size().reset_index(name='SMS')
            df_final = freq.merge(
                current_df[['B_NUMBER','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS']].drop_duplicates(subset='B_NUMBER'),
                left_on='B Number', right_on='B_NUMBER', how='left'
            )
            df_final = df_final.merge(sms_count, left_on='B Number', right_on='B_NUMBER', how='left')
            df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)
            df_final = df_final[['B Number','Count','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS','SMS']]
            df_final['Count'] = df_final['Count'].astype(int)
            df_final = df_final.sort_values(by='Count', ascending=False)

            # ====== imei sheet ======
            current_df['FULL_DATE'] = pd.to_datetime(current_df['FULL_DATE'])
            imei_group = current_df.groupby('IMEI').agg(
                Count=('IMEI','count'),
                Device_Info=('IMEI', lambda x: f'https://www.imei.info/calc/?imei={x.iloc[0]}'),
                HANDSET_MANUFACTURER=('HANDSET_MANUFACTURER','first'),
                HANDSET_MARKETING_NAME=('HANDSET_MARKETING_NAME','first'),
                First_Use_Date=('FULL_DATE','min'),
                Last_Use_Date=('FULL_DATE','max')
            ).reset_index()

            first_last_addr = []
            for imei in imei_group['IMEI']:
                sub = current_df[current_df['IMEI']==imei].sort_values('FULL_DATE')
                first_addr = sub.iloc[0]['SITE_ADDRESS']
                last_addr = sub.iloc[-1]['SITE_ADDRESS']
                first_last_addr.append((first_addr,last_addr))
            imei_group['First_Use_Address'] = [x[0] for x in first_last_addr]
            imei_group['Last_Use_Address'] = [x[1] for x in first_last_addr]

            imei_group = imei_group[['IMEI','Count','Device_Info','HANDSET_MANUFACTURER','HANDSET_MARKETING_NAME',
                                     'First_Use_Date','Last_Use_Date','First_Use_Address','Last_Use_Address']]
            imei_group['Count'] = imei_group['Count'].astype(int)
            imei_group = imei_group.sort_values(by='Count', ascending=False)

            # ====== site sheet ======
            site_df = current_df[['SITE_ADDRESS','LATITUDE','LONGITUDE','FULL_DATE']].copy()
            site_group = site_df.groupby('SITE_ADDRESS').agg(
                Count=('SITE_ADDRESS','count'),
                Map=('LATITUDE', lambda x: f'https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"LONGITUDE"]}'),
                First_Use_Date=('FULL_DATE','min'),
                Last_Use_Date=('FULL_DATE','max')
            ).reset_index()
            site_group['Count'] = site_group['Count'].astype(int)
            site_group = site_group.sort_values(by='Count', ascending=False)
            site_group = site_group[['SITE_ADDRESS','Count','Map','First_Use_Date','Last_Use_Date']]

            sheets = {"calls": df_final, "imei": imei_group, "site": site_group}
            output = save_to_excel_sheets(sheets)
            st.download_button("تحميل تقرير فودافون", data=output, file_name="vodafone_report.xlsx")
