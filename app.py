# ================== تقرير فودافون ==================
def generate_vodafone_report(df):
    required_cols = [
        'B_NUMBER','B_NUMBER_NATIONAL_ID','B_NUMBER_SITE_ADDRESS',
        'FULL_DATE','SERVICE',
        'IMEI','HANDSET_MANUFACTURER','HANDSET_MARKETING_NAME',
        'SITE_ADDRESS','LATITUDE','LONGITUDE'
    ]

    for col in required_cols:
        if col not in df.columns:
            st.error(f"العمود {col} غير موجود في الملف")
            return None

    df2 = df.copy()
    df2['B_NUMBER'] = df2['B_NUMBER'].astype(str)
    df2['IMEI'] = df2['IMEI'].astype(str)
    df2['FULL_DATE'] = pd.to_datetime(df2['FULL_DATE'], errors='coerce')

    freq = df2['B_NUMBER'].value_counts().reset_index()
    freq.columns = ['B Number','Count']

    sms_count = (
        df2[df2['SERVICE'].astype(str).str.strip()
            .isin(["Short message MO/PP","Short message MT/PP"])]
        .groupby('B_NUMBER')
        .size()
        .reset_index(name='SMS')
    )

    call_dates = (
        df2.groupby('B_NUMBER')['FULL_DATE']
        .agg(First_Call='min', Last_Call='max')
        .reset_index()
    )

    base_info = (
        df2[['B_NUMBER','B_NUMBER_NATIONAL_ID','B_NUMBER_SITE_ADDRESS']]
        .drop_duplicates(subset='B_NUMBER')
    )

    df_final = freq.merge(
        base_info,
        left_on='B Number',
        right_on='B_NUMBER',
        how='left'
    )

    df_final = df_final.merge(
        sms_count,
        left_on='B Number',
        right_on='B_NUMBER',
        how='left'
    )

    df_final = df_final.merge(
        call_dates,
        left_on='B Number',
        right_on='B_NUMBER',
        how='left'
    )

    df_final['SMS'] = df_final['SMS'].fillna(0).astype(int)
    df_final['Count'] = df_final['Count'].astype(int)
    df_final['B Number Id'] = df_final['B_NUMBER_NATIONAL_ID'].astype(str)

    df_final = df_final[
        ['B Number','Count','B Number Id',
         'B_NUMBER_SITE_ADDRESS','SMS','First_Call','Last_Call']
    ].sort_values(by='Count', ascending=False)

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
        sub = df2[df2['IMEI'] == imei].sort_values('FULL_DATE')
        first_last_addr.append(
            (sub.iloc[0]['SITE_ADDRESS'], sub.iloc[-1]['SITE_ADDRESS'])
        )

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
        df.to_excel(writer, sheet_name="cheet", index=False)

    output.seek(0)
    return format_excel_sheets(output, header_color="FF0000", company="vodafone")


# ================== تقرير أورانج ==================
def generate_orange_report(df):
    df.columns = df.columns.str.strip().str.upper()

    required_cols = [
        'TARGET_MSISDN','TARGET_IMEI','TARGET_IMSI','TARGET_IMEI_TYPE',
        'EVENT_START_TIME','CALL_DURATION','EVENT_DIRECTION',
        'OTHER_MSISDN','OTHER_NAME','OTHER_ID','OTHER_ADDRESS',
        'CELL_ADDRESS','CELL_LAT','CELL_LONG','OTHER_CELL_ADDRESS'
    ]

    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        st.error(f"الأعمدة التالية غير موجودة في الملف: {missing_cols}")
        return None

    numbers = df['OTHER_MSISDN'].astype(str)
    freq = numbers.value_counts().reset_index()
    freq.columns = ['B Number','Count']

    base_info = (
        df[['OTHER_MSISDN','OTHER_NAME','OTHER_ADDRESS','OTHER_ID','OTHER_CELL_ADDRESS']]
        .drop_duplicates(subset='OTHER_MSISDN')
    )

    calls_df = freq.merge(
        base_info,
        left_on='B Number',
        right_on='OTHER_MSISDN',
        how='left'
    )

    sms_count = (
        df[df['EVENT_DIRECTION'].astype(str).str.strip() == "SMSMT"]
        .groupby('OTHER_MSISDN')
        .size()
        .reset_index(name='SMS')
    )

    calls_df = calls_df.merge(
        sms_count,
        left_on='B Number',
        right_on='OTHER_MSISDN',
        how='left'
    )

    calls_df['SMS'] = calls_df['SMS'].fillna(0).astype(int)
    calls_df['Count'] = calls_df['Count'].astype(int)

    calls_df['B Number id'] = calls_df['OTHER_ID'].apply(
        lambda x: str(int(x)) if pd.notna(x) else ''
    )

    calls_df = calls_df[
        ['B Number','Count','OTHER_NAME','OTHER_ADDRESS',
         'B Number id','OTHER_CELL_ADDRESS','SMS']
    ]

    calls_df.columns = [
        'B Number','Count','B Full Name','B Address',
        'B Number id','other site','SMS'
    ]

    df['EVENT_START_TIME'] = pd.to_datetime(df['EVENT_START_TIME'], errors='coerce')

    call_dates = (
        df.groupby('OTHER_MSISDN')['EVENT_START_TIME']
        .agg(First_Call='min', Last_Call='max')
        .reset_index()
    )

    calls_df = calls_df.merge(
        call_dates,
        left_on='B Number',
        right_on='OTHER_MSISDN',
        how='left'
    ).drop(columns='OTHER_MSISDN')

    calls_df = calls_df.sort_values(by='Count', ascending=False)

    df['TARGET_IMEI'] = df['TARGET_IMEI'].apply(
        lambda x: str(int(x)) if pd.notna(x) else ''
    )

    imei_group = df.groupby('TARGET_IMEI').agg(
        Count=('TARGET_IMEI','count'),
        TARGET_IMEI_TYPE=('TARGET_IMEI_TYPE','first'),
        First_Use_Date=('EVENT_START_TIME','min'),
        Last_Use_Date=('EVENT_START_TIME','max'),
        First_Use_Address=('CELL_ADDRESS','first'),
        Last_Use_Address=('CELL_ADDRESS','last')
    ).reset_index()

    imei_group['Device Info'] = imei_group['TARGET_IMEI'].apply(
        lambda x: f'https://www.imei.info/calc/?imei={x}'
    )

    imei_group = imei_group[
        ['TARGET_IMEI','Count','TARGET_IMEI_TYPE','Device Info',
         'First_Use_Date','Last_Use_Date','First_Use_Address','Last_Use_Address']
    ]

    imei_group.columns = [
        'IMEI','Count','TARGET_IMEI_TYPE','Device Info',
        'First_Use_Date','Last_Use_Date','First_Use_Address','Last_Use_Address'
    ]

    imei_group['Count'] = imei_group['Count'].astype(int)
    imei_group = imei_group.sort_values(by='Count', ascending=False)

    site_df = df.groupby('CELL_ADDRESS').agg(
        Count=('CELL_ADDRESS','count'),
        First_Use_Date=('EVENT_START_TIME','min'),
        Last_Use_Date=('EVENT_START_TIME','max'),
        LAT=('CELL_LAT','first'),
        LON=('CELL_LONG','first')
    ).reset_index()

    site_df['Map'] = site_df.apply(
        lambda row: f'https://www.google.com/maps/search/?api=1&query={row["LAT"]},{row["LON"]}'
        if pd.notna(row["LAT"]) and pd.notna(row["LON"]) else '',
        axis=1
    )

    site_df = site_df[['CELL_ADDRESS','Count','Map','First_Use_Date','Last_Use_Date']]
    site_df = site_df.sort_values(by='Count', ascending=False)

    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        calls_df.to_excel(writer, sheet_name="calls", index=False)
        imei_group.to_excel(writer, sheet_name="imei", index=False)
        site_df.to_excel(writer, sheet_name="site", index=False)
        df.to_excel(writer, sheet_name="cheet", index=False)

    output.seek(0)
    return format_excel_sheets(output, header_color="FF6600", company="orange")


# ================== أزرار التحليل ==================
if current_df is not None:
    st.subheader("توليد تقارير")

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("تقرير اتصالات"):
            output = generate_etisalat_report(current_df, original_df)
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
