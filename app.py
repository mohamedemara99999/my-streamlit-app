import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

st.title('Excel Analyzer Tool - Streamlit Version')

uploaded_file = st.file_uploader("اختر ملف Excel", type=['xlsx','xls'])

if uploaded_file:
    # قراءة الملف
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    st.subheader('معاينة البيانات')
    st.dataframe(df)

    # ================== تقرير اتصالات ==================
    def generate_etisalat_report(df):
        required_cols = ['Originating_Number', 'Terminating_Number', 'Network_Activity_Type_Name',
                         'Call_Start_Date','B_Number_Full_Name', 'B_Number_Address',
                         'B_Number_MU_Site_Address','B_Number_MU_Latitude','B_Number_MU_Longitude',
                         'IMEI_Number','Site_Address','Latitude','Longitude']
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
        imei_summary['Device Info'] = imei_summary['IMEI'].apply(lambda x: f'https://www.imei.info/calc/?imei={x}')
        imei_summary['Count'] = imei_summary['Count'].astype(int)

        # site sheet
        site_df = df[['Site_Address','Latitude','Longitude','Call_Start_Date']].copy()
        site_group = site_df.groupby('Site_Address').agg(
            Count=('Site_Address','count'),
            Map=('Latitude', lambda x: f'https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"Longitude"]}'),
            First_Use_Date=('Call_Start_Date','min'),
            Last_Use_Date=('Call_Start_Date','max')
        ).reset_index()
        site_group['Count'] = site_group['Count'].astype(int)

        # حفظ جميع Sheets في BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='calls', index=False)
            imei_summary.to_excel(writer, sheet_name='imei', index=False)
            site_group.to_excel(writer, sheet_name='site', index=False)
        output.seek(0)
        return output

    # ================== تقرير فودافون ==================
    def generate_vodafone_report(df):
        required_cols = ['B_NUMBER','B_NUMBER_FIRST_NAME','B_NUMBER_LAST_NAME','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS',
                         'IMEI','HANDSET_MANUFACTURER','HANDSET_MARKETING_NAME','FULL_DATE','SITE_ADDRESS','LATITUDE','LONGITUDE','SERVICE']
        for col in required_cols:
            if col not in df.columns:
                st.error(f"العمود {col} غير موجود في الملف")
                return None

        df_copy = df.copy()
        df_copy['B Full Name'] = df_copy['B_NUMBER_FIRST_NAME'].fillna('') + ' ' + df_copy['B_NUMBER_LAST_NAME'].fillna('')

        numbers = df_copy['B_NUMBER'].astype(str)
        freq = numbers.value_counts().reset_index()
        freq.columns = ['B Number','Count']

        sms_count = df_copy[df_copy['SERVICE'].astype(str).str.strip()=='Short message MO/PP'].groupby('B_NUMBER').size().reset_index(name='SMS')
        df_final = freq.merge(
            df_copy[['B_NUMBER','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS']].drop_duplicates(subset='B_NUMBER'),
            left_on='B Number', right_on='B_NUMBER', how='left'
        )
        df_final = df_final.merge(sms_count, left_on='B Number', right_on='B_NUMBER', how='left')
        df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)
        df_final = df_final[['B Number','Count','B Full Name','B_NUMBER_ADDRESS','B_NUMBER_SITE_ADDRESS','SMS']]
        df_final['Count'] = df_final['Count'].astype(int)

        df_copy['FULL_DATE'] = pd.to_datetime(df_copy['FULL_DATE'])
        imei_group = df_copy.groupby('IMEI').agg(
            Count=('IMEI','count'),
            Device_Info=('IMEI', lambda x: f'https://www.imei.info/calc/?imei={x.iloc[0]}'),
            HANDSET_MANUFACTURER=('HANDSET_MANUFACTURER','first'),
            HANDSET_MARKETING_NAME=('HANDSET_MARKETING_NAME','first'),
            First_Use_Date=('FULL_DATE','min'),
            Last_Use_Date=('FULL_DATE','max')
        ).reset_index()

        first_last_addr = []
        for imei in imei_group['IMEI']:
            sub = df_copy[df_copy['IMEI']==imei].sort_values('FULL_DATE')
            first_last_addr.append((sub.iloc[0]['SITE_ADDRESS'], sub.iloc[-1]['SITE_ADDRESS']))
        imei_group['First_Use_Address'] = [x[0] for x in first_last_addr]
        imei_group['Last_Use_Address'] = [x[1] for x in first_last_addr]

        imei_group['Count'] = imei_group['Count'].astype(int)

        site_df = df_copy[['SITE_ADDRESS','LATITUDE','LONGITUDE','FULL_DATE']].copy()
        site_group = site_df.groupby('SITE_ADDRESS').agg(
            Count=('SITE_ADDRESS','count'),
            Map=('LATITUDE', lambda x: f'https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"LONGITUDE"]}'),
            First_Use_Date=('FULL_DATE','min'),
            Last_Use_Date=('FULL_DATE','max')
        ).reset_index()
        site_group['Count'] = site_group['Count'].astype(int)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='calls', index=False)
            imei_group.to_excel(writer, sheet_name='imei', index=False)
            site_group.to_excel(writer, sheet_name='site', index=False)
        output.seek(0)
        return output

    st.subheader('توليد التقارير')
    if st.button('تقرير اتصالات'):
        result = generate_etisalat_report(df)
        if result:
            st.download_button('تحميل تقرير اتصالات', result, file_name='etisalat_report.xlsx')

    if st.button('تقرير فودافون'):
        result = generate_vodafone_report(df)
        if result:
            st.download_button('تحميل تقرير فودافون', result, file_name='vodafone_report.xlsx')
