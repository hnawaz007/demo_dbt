import pandas as pd

def run_attendance_quality_checks_to_excel(df: pd.DataFrame, month_start: str, month_end: str,
                                           excel_path: str = 'attendance_quality_report.xlsx',
                                           valid_statuses=None):
    """
    Runs data quality checks on access logs and saves results + summary to Excel.

    Args:
        df (pd.DataFrame): Data with columns: who, when, What, location, cardnum
        month_start (str): Start of the month (e.g., '2025-05-01')
        month_end (str): End of the month (e.g., '2025-05-31')
        excel_path (str): Output Excel file path
        valid_statuses (set): Allowed values for the 'What' column
    """
    df = df.copy()
    df['when'] = pd.to_datetime(df['when'])

    if valid_statuses is None:
        valid_statuses = {'Access Granted'}

    issues = {}

    # 1. Completeness
    days_expected = pd.date_range(month_start, month_end).nunique()
    days_per_location = df.groupby('location')['when'].nunique()
    issues['locations_with_missing_days'] = days_per_location[days_per_location < days_expected].reset_index()

    emp_days = df.groupby(['location', 'who'])['when'].nunique()
    issues['employees_with_incomplete_attendance'] = emp_days[emp_days < days_expected].reset_index()

    # 2. Duplicates
    duplicates = df[df.duplicated(subset=['location', 'who', 'when'], keep=False)]
    issues['duplicate_entries'] = duplicates

    # 3. Employee in multiple locations
    multi_location_emps = df.groupby('who')['location'].nunique()
    issues['employees_in_multiple_locations'] = multi_location_emps[multi_location_emps > 1].reset_index()

    # 4. Invalid date and status
    valid_range = df['when'].between(month_start, month_end)
    issues['dates_out_of_range'] = df[~valid_range]

    issues['invalid_access_statuses'] = df[~df['What'].isin(valid_statuses)]

    # 5. Weekend Access
    df['weekday'] = df['when'].dt.weekday
    issues['access_granted_on_weekends'] = df[(df['weekday'] >= 5) & (df['What'] == 'Access Granted')]

    # 6. Last date in data
    last_date = df['when'].max()
    issues['last_date_in_data'] = pd.DataFrame({'last_date': [last_date]})

    # 7. Build summary sheet
    summary_rows = []
    for key, value in issues.items():
        if isinstance(value, pd.Series):
            count = value.shape[0]
        elif isinstance(value, pd.DataFrame):
            count = value.shape[0]
        else:
            count = 1
        summary_rows.append({'check': key, 'issue_count': count})
    summary_df = pd.DataFrame(summary_rows)

    # 8. Write to Excel
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='summary', index=False)

        for key, value in issues.items():
            # Convert Series to DataFrame if needed
            if isinstance(value, pd.Series):
                value = value.reset_index(name='value')
            elif not isinstance(value, pd.DataFrame):
                value = pd.DataFrame([value])
            value.to_excel(writer, sheet_name=key[:31], index=False)

    print(f"✔️ Data quality report saved to: {excel_path}")
--------------------------------------------------------------------------------

run_attendance_quality_checks_to_excel(
    df,
    month_start='2025-05-01',
    month_end='2025-05-31',
    excel_path='May_2025_Attendance_Quality_Report.xlsx'
)

--------------
employee_counts = df.groupby('location')['who'].nunique().reset_index()
employee_counts.columns = ['location', 'unique_employees']
print(employee_counts)

------------------------------------------------------------------------------

def calculate_attendance_percentage(df, month_start: str, month_end: str):
    df = df.copy()
    df['when'] = pd.to_datetime(df['when'])
    
    # Filter to date range and Access Granted only
    df = df[(df['when'].between(month_start, month_end)) & (df['What'] == 'Access Granted')]

    # Expected attendance: each employee per location should show up each weekday
    working_days = pd.date_range(month_start, month_end, freq='B')  # Business days only
    num_workdays = len(working_days)

    # Actual attendance days logged per employee per location
    actual = df.groupby(['location', 'who'])['when'].nunique().reset_index(name='days_present')

    # Expected days per employee
    actual['expected_days'] = num_workdays

    # % per employee, then aggregate by location
    actual['attendance_pct'] = actual['days_present'] / actual['expected_days'] * 100
    office_attendance = actual.groupby('location')['attendance_pct'].mean().reset_index()

    return office_attendance
----------------------------------------
attendance_pct_df = calculate_attendance_percentage(df, '2025-05-01', '2025-05-31')
print(attendance_pct_df)


-------------------------
def compare_monthly_attendance_trend(df, current_start, current_end, previous_start, previous_end):
    def get_attendance_pct(df_subset, start, end):
        df_subset = df_subset.copy()
        df_subset['when'] = pd.to_datetime(df_subset['when'])
        df_subset = df_subset[(df_subset['when'].between(start, end)) & (df_subset['What'] == 'Access Granted')]

        working_days = pd.date_range(start, end, freq='B')
        num_workdays = len(working_days)

        actual = df_subset.groupby(['location', 'who'])['when'].nunique().reset_index(name='days_present')
        actual['expected_days'] = num_workdays
        actual['attendance_pct'] = actual['days_present'] / actual['expected_days'] * 100

        return actual.groupby('location')['attendance_pct'].mean().reset_index()

    # Calculate for current and previous
    current = get_attendance_pct(df, current_start, current_end)
    previous = get_attendance_pct(df, previous_start, previous_end)

    # Merge to compare
    comparison = pd.merge(previous, current, on='location', how='outer', suffixes=('_previous', '_current'))
    comparison['trend'] = comparison['attendance_pct_current'] - comparison['attendance_pct_previous']

    return comparison
	
---------------------------------------
trend_df = compare_monthly_attendance_trend(
    df,
    current_start='2025-05-01',
    current_end='2025-05-31',
    previous_start='2025-04-01',
    previous_end='2025-04-30'
)

print(trend_df)

-----------------------------------------------

📈 1. Trend Analysis
Monthly trends by location: Track attendance % over multiple months per office.

Day-of-week trend: Are some days consistently lower (e.g., Mondays or Fridays)?

Time-of-day analysis (if timestamp available): Late check-ins or peak access hours.

👥 2. Employee Behavior Patterns
Top absentees / most regular employees

Employees with irregular attendance patterns

Employees accessing multiple offices within the same month (could indicate travel or misuse)

📊 3. Office-Level KPIs
Average attendance %

Variance in attendance across employees

Offices with improving or declining trends

Office capacity utilization (if you know expected headcount per location)

🕵️ 4. Anomaly Detection
Employees showing up on weekends or holidays

Duplicate card usage (same card used at two places on the same day)

No attendance despite having an access card

Unexpected spikes/dips on specific dates (linked to weather, events, etc.)

🔄 5. Comparative Benchmarks
Compare attendance between departments (if available)

Compare remote vs on-site employees (if available)

Benchmark against company-wide average

📅 6. Calendar Mapping
Overlay attendance with:

Public holidays

Company events

Leaves (if data is available)

Work-from-home policies

📦 7. Data Enrichment Opportunities
If you can join with HR master data (e.g., role, team, manager), you can analyze:

Attendance by team or function

Correlation between attendance and performance (if available)

Flag high absenteeism before reviews

