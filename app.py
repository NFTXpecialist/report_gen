import streamlit as st
import pandas as pd

st.set_page_config(page_title="Business Establishment Report", layout="wide")
st.title("Business Establishment Report Generator")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"]) 

if uploaded_file:
    # Directly use the sheet named 'DATABASE'
    try:
        df = pd.read_excel(uploaded_file, sheet_name='DATABASE')
    except ValueError:
        st.error("Sheet named 'DATABASE' not found in the uploaded file.")
        st.stop()

    # Let user select STATUS column if not found automatically
    status_col = None
    for col in df.columns:
        if col.strip().lower() == 'status':
            status_col = col
            break
    if not status_col:
        status_col = st.selectbox(
            "Select the column to use as STATUS (exclude 'assigned' and blank)",
            options=df.columns,
            index=0
        )
    if status_col:
        # Filter out rows where STATUS is 'assigned' or blank
        filtered_df = df[~df[status_col].fillna('').astype(str).str.lower().isin(['assigned', ''])]

        # Try to automatically find the 'date returned' column
        date_returned_col = None
        for col in filtered_df.columns:
            if col.strip().lower() in ['date returned', 'datereturned', 'date_returned']:
                date_returned_col = col
                break
        if not date_returned_col:
            date_returned_col = st.selectbox(
                "Select the column to use as 'date returned' (only rows with a value will be shown)",
                options=filtered_df.columns,
                index=0
            )
        if date_returned_col:
            filtered_df = filtered_df[filtered_df[date_returned_col].notna() & (filtered_df[date_returned_col].astype(str).str.strip() != '')]
            # Parse the date column as datetime
            filtered_df[date_returned_col] = pd.to_datetime(filtered_df[date_returned_col], errors='coerce')
            min_date = filtered_df[date_returned_col].min()
            max_date = filtered_df[date_returned_col].max()
            # Period type selector
            period_type = st.selectbox(
                "Select period type",
                ["Custom Range", "Monthly", "Quarterly", "Semestral", "Annual"],
                index=0
            )
            # Determine default date range based on period type
            default_start = min_date.date() if pd.notnull(min_date) else None
            default_end = max_date.date() if pd.notnull(max_date) else None
            if period_type != "Custom Range" and pd.notnull(max_date):
                if period_type == "Monthly":
                    default_start = max_date.replace(day=1).date()
                    default_end = (max_date + pd.offsets.MonthEnd(0)).date()
                elif period_type == "Quarterly":
                    quarter_options = ["First", "Second", "Third", "Fourth"]
                    quarter_selected = st.selectbox("Select quarter", quarter_options, index=(((max_date.month-1)//3)))
                    q_idx = quarter_options.index(quarter_selected) + 1
                    q_start_month = 3*(q_idx-1)+1
                    q_end_month = q_start_month + 2
                    default_start = max_date.replace(month=q_start_month, day=1).date()
                    # Get last day of quarter
                    if q_end_month == 3:
                        default_end = max_date.replace(month=3, day=31).date()
                    elif q_end_month == 6:
                        default_end = max_date.replace(month=6, day=30).date()
                    elif q_end_month == 9:
                        default_end = max_date.replace(month=9, day=30).date()
                    else:
                        default_end = max_date.replace(month=12, day=31).date()
                elif period_type == "Semestral":
                    sem_options = ["First", "Second"]
                    sem_selected = st.selectbox("Select semester", sem_options, index=0 if max_date.month <= 6 else 1)
                    if sem_selected == "First":
                        default_start = max_date.replace(month=1, day=1).date()
                        default_end = max_date.replace(month=6, day=30).date()
                    else:
                        default_start = max_date.replace(month=7, day=1).date()
                        default_end = max_date.replace(month=12, day=31).date()
                elif period_type == "Annual":
                    default_start = max_date.replace(month=1, day=1).date()
                    default_end = max_date.replace(month=12, day=31).date()
            # Clamp default_start and default_end to available data range
            min_val = min_date.date() if pd.notnull(min_date) else None
            max_val = max_date.date() if pd.notnull(max_date) else None
            if default_start is not None and min_val is not None:
                default_start = max(default_start, min_val)
            if default_end is not None and max_val is not None:
                default_end = min(default_end, max_val)
            # Date range picker
            start_date, end_date = st.date_input(
                "Select date range to display", 
                value=(default_start, default_end),
                min_value=min_val,
                max_value=max_val
            )
            # Filter by the selected date range
            mask = (filtered_df[date_returned_col] >= pd.Timestamp(start_date)) & (filtered_df[date_returned_col] <= pd.Timestamp(end_date))
            filtered_df = filtered_df[mask]
            st.write(f"Filtered Data Preview ({start_date} to {end_date}):")
            st.dataframe(filtered_df.head())

            # Always define category_col and occupancy_col before advanced features block
            category_col = None
            occupancy_col = None
            for col in filtered_df.columns:
                if col.strip().lower() == 'category':
                    category_col = col
                if col.strip().lower() in ['occupancy', 'occupancy type', 'occupancy_type']:
                    occupancy_col = col
            if not category_col:
                for col in filtered_df.columns:
                    if 'category' in col.lower():
                        category_col = col
                        break
            if not occupancy_col:
                for col in filtered_df.columns:
                    if 'occupancy' in col.lower():
                        occupancy_col = col
                        break
            # Let user select AOR column if not found automatically
            aor_col = None
            for col in filtered_df.columns:
                if col.strip().lower() in ['aor', 'area of responsibility']:
                    aor_col = col
                    break
            if not aor_col:
                aor_col = st.selectbox(
                    "Select the column to use as AOR (Area of Responsibility)",
                    options=filtered_df.columns,
                    index=0
                )
            # --- Advanced Features: Only if all key columns are set ---
            if aor_col and status_col and category_col and occupancy_col:
                # Multi-select filters for AOR, STATUS, CATEGORY, OCCUPANCY TYPE
                filter_widgets = {}
                unique_aors = filtered_df[aor_col].dropna().unique().tolist()
                filter_widgets['AOR'] = st.multiselect('Filter by AOR', unique_aors, default=unique_aors)
                unique_status = filtered_df[status_col].dropna().unique().tolist()
                filter_widgets['STATUS'] = st.multiselect('Filter by STATUS', unique_status, default=unique_status)
                unique_cat = filtered_df[category_col].dropna().unique().tolist()
                filter_widgets['CATEGORY'] = st.multiselect('Filter by CATEGORY', unique_cat, default=unique_cat)
                unique_occ = filtered_df[occupancy_col].dropna().unique().tolist()
                filter_widgets['OCCUPANCY TYPE'] = st.multiselect('Filter by OCCUPANCY TYPE', unique_occ, default=unique_occ)
                # Apply filters
                if 'AOR' in filter_widgets:
                    filtered_df = filtered_df[filtered_df[aor_col].isin(filter_widgets['AOR'])]
                if 'STATUS' in filter_widgets:
                    filtered_df = filtered_df[filtered_df[status_col].isin(filter_widgets['STATUS'])]
                if 'CATEGORY' in filter_widgets:
                    filtered_df = filtered_df[filtered_df[category_col].isin(filter_widgets['CATEGORY'])]
                if 'OCCUPANCY TYPE' in filter_widgets:
                    filtered_df = filtered_df[filtered_df[occupancy_col].isin(filter_widgets['OCCUPANCY TYPE'])]
                st.write(f"Filtered Data Preview (after advanced filters):")
                st.dataframe(filtered_df.head())

                # --- Summary Cards ---
                total_apps = len(filtered_df)
                total_new = filtered_df[category_col].astype(str).str.upper().eq('NEW').sum()
                total_renew = filtered_df[category_col].astype(str).str.upper().eq('RENEW').sum()
                top_aor = filtered_df[aor_col].mode()[0] if not filtered_df.empty else 'N/A'
                top_occ = filtered_df[occupancy_col].mode()[0] if not filtered_df.empty else 'N/A'
                st.markdown('---')
                col1, col2, col3, col4, col5 = st.columns(5)
                col1.metric('Total Applications', total_apps)
                col2.metric('Total NEW', total_new)
                col3.metric('Total RENEW', total_renew)
                col4.metric('Top AOR', top_aor)
                col5.metric('Top Occupancy Type', top_occ)
                st.markdown('---')

                # --- Export Buttons ---
                import io
                export_buffer = io.BytesIO()
                with pd.ExcelWriter(export_buffer, engine='xlsxwriter') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='Filtered Data')
                st.download_button('Download Filtered Data as Excel', data=export_buffer.getvalue(), file_name='filtered_data.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

            if aor_col:
                # Count each STATUS value per AOR
                # Define the desired AOR order
                desired_aor_order = [
                    'ILIGAN CITY CENTRAL FS',
                    'SAN MIGUEL FSS',
                    'SARAY FSS',
                    'STA. FILOMENA FSS',
                    'DALIPUGA FSS',
                    'BURUUN FSS',
                    'MA. CRISTINA FSS',
                    'TUBOD FSS'
                ]
                # Normalize AOR values to canonical names
                canonical_aor_map = {
                    'iligan city central fs': 'ILIGAN CITY CENTRAL FS',
                    'central fs': 'ILIGAN CITY CENTRAL FS',
                    'san miguel fss': 'SAN MIGUEL FSS',
                    'saray fss': 'SARAY FSS',
                    'sta. filomena fss': 'STA. FILOMENA FSS',
                    'st. filomena fss': 'STA. FILOMENA FSS',
                    'dalipuga fss': 'DALIPUGA FSS',
                    'buruun fss': 'BURUUN FSS',
                    'buru-un fss': 'BURUUN FSS',
                    'ma. cristina fss': 'MA. CRISTINA FSS',
                    'tubod fss': 'TUBOD FSS',
                }
                def normalize_aor(aor):
                    if not isinstance(aor, str):
                        return aor
                    key = aor.strip().lower().replace('  ', ' ')
                    return canonical_aor_map.get(key, aor.strip())
                filtered_df[aor_col] = filtered_df[aor_col].apply(normalize_aor)
                # Generate the crosstab report
                report_table = pd.crosstab(filtered_df[aor_col], filtered_df[status_col], margins=True, margins_name='Total')
                # If empty, initialize with all expected STATUS columns
                if report_table.shape[1] == 0:
                    status_values = list(filtered_df[status_col].unique())
                    if len(status_values) == 0:
                        status_values = ['NEW', 'RENEW']
                    report_table = pd.DataFrame(columns=status_values)
                # Reindex to match the desired order, keep any additional AORs, and keep 'Total' at the end
                row_order = [aor for aor in desired_aor_order if aor in report_table.index]
                # Add missing AORs as zeros
                for aor in desired_aor_order:
                    if aor not in report_table.index:
                        report_table.loc[aor] = 0
                        row_order.append(aor)
                # Add any extra AORs
                extra_aors = [aor for aor in report_table.index if aor not in desired_aor_order and aor != 'Total']
                row_order += extra_aors
                if 'Total' in report_table.index:
                    row_order.append('Total')
                report_table = report_table.loc[row_order]
                st.write("### Report: Count of STATUS per AOR (with Totals)")
                st.dataframe(report_table)
                # Export STATUS report
                import io
                report_export_buffer = io.BytesIO()
                with pd.ExcelWriter(report_export_buffer, engine='xlsxwriter') as writer:
                    report_table.to_excel(writer, sheet_name='Status per AOR')
                st.download_button('Download STATUS per AOR Report as Excel', data=report_export_buffer.getvalue(), file_name='status_per_aor.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                # Bar chart for STATUS per AOR
                import plotly.express as px
                if not report_table.empty and len(report_table.columns) > 1:
                    status_cols = [col for col in report_table.columns if col != 'Total']
                    chart_df = report_table.drop('Total', errors='ignore')
                    chart_df = chart_df.drop('Total', axis=1, errors='ignore')
                    chart_df = chart_df.reset_index().melt(id_vars=aor_col, value_vars=status_cols, var_name='Status', value_name='Count')
                    fig = px.bar(chart_df, x=aor_col, y='Count', color='Status', barmode='group', title='Applications per AOR by Status')
                    st.plotly_chart(fig, use_container_width=True)

                # --- Second Report: Count by Occupancy Type per AOR ---
                # Try to automatically find the occupancy column
                occupancy_col = None
                for col in filtered_df.columns:
                    if col.strip().lower() in ['occupancy', 'occupancy type', 'occupancy_type']:
                        occupancy_col = col
                        break
                if not occupancy_col:
                    occupancy_col = st.selectbox(
                        "Select the column to use as Occupancy Type",
                        options=filtered_df.columns,
                        index=0
                    )
                if occupancy_col:
                    # Normalize occupancy type values to canonical names
                    canonical_occupancy_map = {
                        'assembly': 'Assembly',
                        'educational': 'Educational',
                        'day care': 'Day Care',
                        'health care': 'Health Care',
                        'residential board and care': 'Residential Board and Care',
                        'detention & correctional': 'Detention & Correctional',
                        'hotel i': 'Hotel I',
                        'hotel 1': 'Hotel I',
                        'dormitories': 'Dormitories',
                        'apartment buildings': 'Apartment Buildings',
                        'lodging & rooming house': 'Lodging & Rooming House',
                        'single & two family dwelling unit': 'Single & Two Family Dwelling Unit',
                        'mercantile': 'Mercantile',
                        'business': 'Business',
                        'industrial': 'Industrial',
                        'storage': 'Storage',
                        'special structures': 'Special Structures',
                        'non-structural (ex. vehicles used as rolling store & etc.)': 'Non-Structural (ex. Vehicles used as Rolling Store & etc.)',
                        'total number of inspected from 1st inspection': 'Total Number of Inspected from 1st Inspection',
                    }
                    def normalize_occupancy(val):
                        if not isinstance(val, str):
                            return val
                        key = val.strip().lower().replace('  ', ' ')
                        return canonical_occupancy_map.get(key, val.strip())
                    filtered_df[occupancy_col] = filtered_df[occupancy_col].apply(normalize_occupancy)
                    occupancy_table = pd.crosstab(filtered_df[aor_col], filtered_df[occupancy_col], margins=True, margins_name='Total')
                    # Reindex to match the desired AOR order, keep any additional AORs, and keep 'Total' at the end
                    row_order2 = [aor for aor in desired_aor_order if aor in occupancy_table.index]
                    for aor in desired_aor_order:
                        if aor not in occupancy_table.index:
                            occupancy_table.loc[aor] = 0
                            row_order2.append(aor)
                    extra_aors2 = [aor for aor in occupancy_table.index if aor not in desired_aor_order and aor != 'Total']
                    row_order2 += extra_aors2
                    if 'Total' in occupancy_table.index:
                        row_order2.append('Total')
                    occupancy_table = occupancy_table.loc[row_order2]

                    # Reorder columns according to the image
                    occupancy_col_order = [
                        'Assembly',
                        'Educational',
                        'Day Care',
                        'Health Care',
                        'Residential Board and Care',
                        'Detention & Correctional',
                        'Hotel I',
                        'Dormitories',
                        'Apartment Buildings',
                        'Lodging & Rooming House',
                        'Single & Two Family Dwelling Unit',
                        'Mercantile',
                        'Business',
                        'Industrial',
                        'Storage',
                        'Special Structures',
                        'Non-Structural (ex. Vehicles used as Rolling Store & etc.)',
                        'Total Number of Inspected from 1st Inspection'
                    ]
                    # Add NEW and RENEW columns (count per AOR)
                    new_col = 'NEW'
                    renew_col = 'RENEW'
                    if new_col not in occupancy_table.columns:
                        occupancy_table[new_col] = 0
                    if renew_col not in occupancy_table.columns:
                        occupancy_table[renew_col] = 0
                    # Calculate NEW and RENEW counts per AOR from CATEGORY column
                    category_col = None
                    for col in filtered_df.columns:
                        if col.strip().lower() == 'category':
                            category_col = col
                            break
                    if not category_col:
                        category_col = st.selectbox(
                            "Select the column to use as CATEGORY (for NEW/RENEW count)",
                            options=filtered_df.columns,
                            index=0
                        )
                    for aor in occupancy_table.index:
                        if aor in ['Total']:
                            continue
                        occupancy_table.at[aor, new_col] = ((filtered_df[aor_col] == aor) & (filtered_df[category_col].astype(str).str.upper() == new_col)).sum()
                        occupancy_table.at[aor, renew_col] = ((filtered_df[aor_col] == aor) & (filtered_df[category_col].astype(str).str.upper() == renew_col)).sum()
                    # Compute totals for NEW and RENEW columns
                    if 'Total' not in occupancy_table.index:
                        occupancy_table.loc['Total'] = 0
                    occupancy_table.at['Total', new_col] = occupancy_table.loc[[aor for aor in occupancy_table.index if aor != 'Total'], new_col].sum()
                    occupancy_table.at['Total', renew_col] = occupancy_table.loc[[aor for aor in occupancy_table.index if aor != 'Total'], renew_col].sum()
                    # Ensure all columns in the desired order are present
                    for col in occupancy_col_order:
                        if col not in occupancy_table.columns:
                            occupancy_table[col] = 0
                    # Insert NEW and RENEW before Assembly
                    col_order2 = [new_col, renew_col] + [col for col in occupancy_col_order]
                    extra_cols2 = [col for col in occupancy_table.columns if col not in col_order2 and col != 'Total']
                    col_order2 += extra_cols2
                    if 'Total' in occupancy_table.columns:
                        col_order2.append('Total')
                    occupancy_table = occupancy_table[col_order2]

                    st.write("### Report: Count of Occupancy Type per AOR (with Totals)")
                    st.dataframe(occupancy_table)
                    # Export OCCUPANCY report
                    occ_export_buffer = io.BytesIO()
                    with pd.ExcelWriter(occ_export_buffer, engine='xlsxwriter') as writer:
                        occupancy_table.to_excel(writer, sheet_name='Occupancy per AOR')
                    st.download_button('Download OCCUPANCY TYPE per AOR Report as Excel', data=occ_export_buffer.getvalue(), file_name='occupancy_per_aor.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                    # Bar chart for Occupancy per AOR (top 10 types)
                    occ_cols = [col for col in occupancy_table.columns if col not in ['NEW', 'RENEW', 'Total']]
                    occ_chart_df = occupancy_table.drop('Total', errors='ignore')
                    occ_chart_df = occ_chart_df.reset_index().melt(id_vars=aor_col, value_vars=occ_cols, var_name='Occupancy Type', value_name='Count')
                    occ_chart_df = occ_chart_df[occ_chart_df['Count'] > 0]
                    occ_chart_df = occ_chart_df.sort_values('Count', ascending=False).groupby('Occupancy Type').head(10)
                    fig2 = px.bar(occ_chart_df, x=aor_col, y='Count', color='Occupancy Type', barmode='group', title='Top Occupancy Types per AOR')
                    st.plotly_chart(fig2, use_container_width=True)
                    # Data validation & export issues
                    missing_mask = (
                        filtered_df[status_col].isna() |
                        filtered_df[aor_col].isna() |
                        filtered_df[category_col].isna() |
                        filtered_df[occupancy_col].isna() |
                        filtered_df[date_returned_col].isna()
                    )
                    issues_df = filtered_df[missing_mask]
                    if not issues_df.empty:
                        st.warning(f"There are {len(issues_df)} rows with missing key fields. Please review.")
                        issues_buffer = io.BytesIO()
                        with pd.ExcelWriter(issues_buffer, engine='xlsxwriter') as writer:
                            issues_df.to_excel(writer, index=False, sheet_name='Issues')
                        st.download_button('Download Issues Report as Excel', data=issues_buffer.getvalue(), file_name='issues_report.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                    # --- Dynamic Remark for Totals Difference ---
                    occ_total = None
                    status_total = None
                    if 'Total' in occupancy_table.index and 'Total' in occupancy_table.columns:
                        occ_total = occupancy_table.loc['Total', 'Total'] if 'Total' in occupancy_table.columns else None
                    if 'Total' in report_table.index and 'Total' in report_table.columns:
                        status_total = report_table.loc['Total', 'Total'] if 'Total' in report_table.columns else None

                    remark = ""
                    if occ_total is not None and status_total is not None:
                        if occ_total > status_total:
                            # Check for applications with multiple occupancy types
                            if occupancy_col in filtered_df.columns:
                                multi_occ = filtered_df.groupby(filtered_df.index)[occupancy_col].nunique()
                                multi_occ_count = (multi_occ > 1).sum()
                                if multi_occ_count > 0:
                                    remark += f"There are {multi_occ_count} applications with multiple occupancy types. This causes the occupancy type total to be higher than the status total. "
                            remark += "A single application can be counted under multiple occupancy types."
                        elif occ_total < status_total:
                            # Check for applications missing occupancy type
                            if occupancy_col in filtered_df.columns:
                                missing_occ = filtered_df[occupancy_col].isna().sum()
                                if missing_occ > 0:
                                    remark += f"There are {missing_occ} applications without an occupancy type, which reduces the occupancy type total. "
                            remark += "Some applications may not have an occupancy type."
                        else:
                            remark = "The totals match."
                    else:
                        remark = "Totals could not be compared."
                    st.caption(f"Remarks: {remark}")
                else:
                    st.error("No suitable Occupancy Type column found. Please check your file.")
            else:
                st.error("No suitable AOR column found. Please check your file.")
        else:
            st.error("No suitable 'date returned' column found. Please check your file.")
    else:
        st.error("No suitable STATUS column found. Please check your file.")

# --- Inspectors Output Section ---
st.markdown('---')
st.header('Inspectors Output')
if 'filtered_df' in locals() and not filtered_df.empty:
    # 1. Auto-detect or select Inspector and Team Leader columns
    inspector_col = None
    team_leader_col = None
    for col in filtered_df.columns:
        if col.strip().lower() in ['inspector', 'inspected by']:
            inspector_col = col
        if col.strip().lower() in ['team leader', 'team_leader', 'leader']:
            team_leader_col = col
    if not inspector_col:
        inspector_col = st.selectbox('Select Inspector column', filtered_df.columns, index=0)
    if not team_leader_col:
        team_leader_col = st.selectbox('Select Team Leader column', filtered_df.columns, index=0)
    # 2. Combine as Team key
    filtered_df['Team'] = filtered_df[inspector_col].astype(str).str.strip() + ' / ' + filtered_df[team_leader_col].astype(str).str.strip()
    team_groups = filtered_df.groupby('Team')
    # 3. Summary Table per Team using groupby.apply
    def team_summary_fn(group):
        total = len(group)
        new = (group[category_col].astype(str).str.upper() == 'NEW').sum()
        renew = (group[category_col].astype(str).str.upper() == 'RENEW').sum()
        top_aor = group[aor_col].mode()[0] if not group.empty else 'N/A'
        top_occ = group[occupancy_col].mode()[0] if not group.empty else 'N/A'
        first_date = group[date_returned_col].min()
        last_date = group[date_returned_col].max()
        days = ((last_date - first_date).days + 1) if pd.notnull(first_date) and pd.notnull(last_date) else 1
        days = max(days, 1)
        avg_per_day = round(total / days, 2)
        return pd.Series({
            'Total_Applications': total,
            'NEW': new,
            'RENEW': renew,
            'Top_AOR': top_aor,
            'Top_Occupancy': top_occ,
            'First_Date': first_date,
            'Last_Date': last_date,
            'Days': days,
            'Avg per Day': avg_per_day
        })
    team_summary = team_groups.apply(team_summary_fn)
    st.subheader('Team Output Summary')
    st.dataframe(team_summary)
    # 4. Drilldown per Team
    st.subheader('Team Drilldown (expand to view details)')
    for team, group in team_groups:
        with st.expander(f"{team} ({len(group)} applications)"):
            st.dataframe(group)
    # 5. Export Inspectors Output
    import io
    team_export_buffer = io.BytesIO()
    import re
    def sanitize_sheet_name(name, existing):
        # Remove invalid characters and truncate
        safe = re.sub(r'[\[\]\:\*\?\/\\]', '', str(name))
        safe = safe.strip()
        safe = safe[:31]
        orig = safe
        i = 1
        # Ensure uniqueness
        while safe in existing:
            safe = (orig[:28] + f'_{i}')[:31]
            i += 1
        existing.add(safe)
        return safe
    with pd.ExcelWriter(team_export_buffer, engine='xlsxwriter') as writer:
        team_summary.to_excel(writer, sheet_name='Team Summary')
        used_sheets = set(['Team Summary'])
        for team, group in team_groups:
            sheet = sanitize_sheet_name(team, used_sheets)
            group.to_excel(writer, sheet_name=sheet, index=False)
    st.download_button('Download Inspectors Output as Excel', data=team_export_buffer.getvalue(), file_name='inspectors_output.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    # 6. Bar chart for team output
    import plotly.express as px
    fig_team = px.bar(team_summary.reset_index(), x='Team', y='Total_Applications', title='Applications per Team', text_auto=True)
    st.plotly_chart(fig_team, use_container_width=True)
else:
    st.info('No filtered data available for Inspectors Output.')
