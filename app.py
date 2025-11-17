import streamlit as st
import pandas as pd
import os
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

st.set_page_config(page_title="Excel Analyzer Tool", layout="wide")

# ================== تسجيل الدخول ==================
password_input = st.text_input("ادخل الباسورد:", type="password")
if password_input != "m7md3mara2592025":
    st.warning("الباسورد غير صحيح")
    st.stop()

st.title("Excel Analyzer Tool")

# ================== رفع الملف ==================
uploaded_file = st.file_uploader("اختر ملف Excel", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.success("تم رفع الملف بنجاح")
    except Exception as e:
        st.error(f"خطأ في فتح الملف: {e}")
        st.stop()

    # ================== عرض البيانات ==================
    st.subheader("Preview of Data")
    st.dataframe(df.head())

    # ================== تحليل اتصالات ==================
    if st.button("تقرير اتصالات"):
        required_cols = [
            'Originating_Number', 'Terminating_Number', 'Network_Activity_Type_Name',
            'Call_Start_Date','B_Number_Full_Name', 'B_Number_Address',
            'B_Number_MU_Site_Address','B_Number_MU_Latitude','B_Number_MU_Longitude',
            'IMEI_Number','Site_Address','Latitude','Longitude'
        ]
        missing_cols = [c for c in required_cols if c not in df.columns]
        if missing_cols:
            st.error(f"الأعمدة التالية مفقودة: {missing_cols}")
        else:
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
            site_df = df[['Site_Address','Latitude','Longitude','Call_Start_Date']].copy()
            site_group = site_df.groupby('Site_Address').agg(
                Count=('Site_Address','count'),
                Map=('Latitude', lambda x: f'https://www.google.com/maps/search/?api=1&query={x.iloc[0]},{site_df.loc[x.index[0],"Longitude"]}' ),
                First_Use_Date=('Call_Start_Date','min'),
                Last_Use_Date=('Call_Start_Date','max')
            ).reset_index()
            site_group['Count'] = site_group['Count'].astype(int)
            site_group = site_group.sort_values(by='Count', ascending=False)
            site_group = site_group[['Site_Address','Count','Map','First_Use_Date','Last_Use_Date']]

            # ====== حفظ الملفات في Excel ======
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name="calls", index=False)
                imei_summary.to_excel(writer, sheet_name="imei", index=False)
                site_group.to_excel(writer, sheet_name="site", index=False)

                wb = writer.book
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

                format_sheet(wb["calls"], hyperlink_col=8)
                format_sheet(wb["imei"], hyperlink_col=3)
                format_sheet(wb["site"], hyperlink_col=3)
                writer.save()
            st.success("تم إنشاء تقرير الاتصالات بنجاح!")

            # زر تحميل
            st.download_button(
                label="تحميل تقرير اتصالات Excel",
                data=output.getvalue(),
                file_name="etisalat_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
