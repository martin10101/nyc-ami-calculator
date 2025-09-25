import os
import pandas as pd

_SCENARIO_EXPORT_ORDER = [
    ("S1_Absolute_Best", "scenario_absolute_best", "AMI_S1_Absolute_Best"),
    ("S2_Client_Oriented", "scenario_client_oriented", "AMI_S2_Client_Oriented"),
    ("S3_Best_3_Band", "scenario_best_3_band", "AMI_S3_Best_3_Band"),
    ("S4_Best_2_Band", "scenario_best_2_band", "AMI_S4_Best_2_Band"),
    ("S5_Alternative", "scenario_alternative", "AMI_S5_Alternative"),
]

def _scenario_to_dataframe(scenario):
    df = pd.DataFrame(scenario['assignments'])
    report_df = df[['unit_id', 'bedrooms', 'net_sf', 'floor', 'assigned_ami']].copy()
    report_df.rename(columns={
        'unit_id': 'Unit ID',
        'bedrooms': 'Bedrooms',
        'net_sf': 'Net SF',
        'floor': 'Floor',
        'assigned_ami': 'Assigned AMI',
    }, inplace=True)
    report_df['Assigned AMI'] = (report_df['Assigned AMI'] * 100).map('{:.0f}%'.format)
    return report_df

def _scenario_summary_frame(display_name, scenario):
    metrics = scenario.get('metrics', {})
    band_mix = metrics.get('band_mix', [])
    band_summary = "; ".join(
        f"{entry['band']}%: {entry['units']} units ({entry['share_of_sf']*100:.1f}% SF)"
        for entry in band_mix
    )
    return pd.DataFrame([
        {
            'Scenario': display_name.replace('_', ' '),
            'WAAMI (%)': round(metrics.get('waami_percent', scenario.get('waami', 0) * 100), 4),
            'Bands Used': ", ".join(f"{band}%" for band in scenario.get('bands', [])),
            'Total Units': metrics.get('total_units'),
            'Total SF': metrics.get('total_sf'),
            'Revenue Score': round(metrics.get('revenue_score', 0.0), 2),
            'Band Mix': band_summary,
        }
    ])

def create_excel_reports(analysis_json, original_file_path, original_headers, output_dir='reports'):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    base_name = os.path.splitext(os.path.basename(original_file_path))[0]
    created_files = []

    scenario_lookup = {
        key: analysis_json.get(key)
        for _, key, _ in _SCENARIO_EXPORT_ORDER
    }

    # --- Individual scenario workbooks ---
    for display_name, analysis_key, _ in _SCENARIO_EXPORT_ORDER:
        scenario = scenario_lookup.get(analysis_key)
        if not scenario or 'assignments' not in scenario:
            continue
        filepath = os.path.join(output_dir, f"{base_name}_{display_name}_Report.xlsx")
        with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
            _scenario_to_dataframe(scenario).to_excel(writer, sheet_name='Assignments', index=False)
            _scenario_summary_frame(display_name, scenario).to_excel(writer, sheet_name='Summary', index=False)
        created_files.append(filepath)

    # --- Updated source workbook ---
    try:
        if original_file_path.endswith('.csv'):
            master_df = pd.read_csv(original_file_path, dtype=str)
        else:
            master_df = pd.read_excel(original_file_path, dtype=str)
    except Exception:
        return created_files

    unit_col = original_headers.get('unit_id')
    ami_col = original_headers.get('client_ami')
    if not unit_col or not ami_col or unit_col not in master_df.columns or ami_col not in master_df.columns:
        return created_files

    # Preserve original AMI column
    insert_idx = master_df.columns.get_loc(ami_col) + 1
    original_label = f"{ami_col}_Original"
    if original_label not in master_df.columns:
        master_df.insert(insert_idx, original_label, master_df[ami_col])
        insert_idx += 1

    # Add scenario columns
    for _, analysis_key, column_name in _SCENARIO_EXPORT_ORDER:
        scenario = scenario_lookup.get(analysis_key)
        if not scenario or 'assignments' not in scenario:
            continue
        mapping = {str(u['unit_id']): f"{int(round(u['assigned_ami'] * 100))}%" for u in scenario['assignments']}
        master_df.insert(insert_idx, column_name, master_df[unit_col].map(mapping).fillna(''))
        insert_idx += 1

    summary_rows = []
    for display_name, analysis_key, _ in _SCENARIO_EXPORT_ORDER:
        scenario = scenario_lookup.get(analysis_key)
        if not scenario:
            continue
        summary_df = _scenario_summary_frame(display_name, scenario)
        summary_rows.append(summary_df)
    summary_sheet = pd.concat(summary_rows, ignore_index=True) if summary_rows else pd.DataFrame()

    updated_source_filepath = os.path.join(output_dir, f"{base_name}_Updated_Source.xlsx")
    with pd.ExcelWriter(updated_source_filepath, engine='xlsxwriter') as writer:
        master_df.to_excel(writer, sheet_name='Units', index=False)
        if not summary_sheet.empty:
            summary_sheet.to_excel(writer, sheet_name='Scenario Summary', index=False)
    created_files.append(updated_source_filepath)

    return created_files
