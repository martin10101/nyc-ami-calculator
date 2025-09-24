import pandas as pd
import os

def create_excel_reports(analysis_json, original_file_path, original_headers, output_dir='reports'):
    """
    Generates multiple Excel reports from the solver's analysis JSON, including an
    updated version of the user's original file with new assignment columns.

    Args:
        analysis_json (dict): The JSON output from the AMI-Optix solver.
        original_file_path (str): The path to the user's original uploaded file.
        original_headers (dict): The mapping of internal names to original column headers.
        output_dir (str): The directory to save the reports in.

    Returns:
        list: A list of file paths for the generated reports.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    base_name = os.path.splitext(os.path.basename(original_file_path))[0]
    created_files = []
    scenarios = analysis_json.get("scenarios", {})

    # --- 1. Create Individual Scenario Reports ---
    scenarios_to_report = {
        "S1_Absolute_Best": scenarios.get("absolute_best", [None])[0],
        "S2_Alternative": scenarios.get("alternative", [None])[0],
        "S3_Best_2_Band": scenarios.get("best_2_band"),
    }

    for name, scenario in scenarios_to_report.items():
        if not scenario or 'assignments' not in scenario:
            continue

        df = pd.DataFrame(scenario['assignments'])
        report_df = df[['unit_id', 'bedrooms', 'net_sf', 'floor', 'assigned_ami']].copy()
        report_df.rename(columns={
            'unit_id': 'Unit ID', 'bedrooms': 'Bedrooms', 'net_sf': 'Net SF',
            'floor': 'Floor', 'assigned_ami': 'Assigned AMI'
        }, inplace=True)
        report_df['Assigned AMI'] = (report_df['Assigned AMI'] * 100).map('{:.0f}%'.format)

        filepath = os.path.join(output_dir, f"{base_name}_{name}_Report.xlsx")
        report_df.to_excel(filepath, index=False)
        created_files.append(filepath)

    # --- 2. Create Updated Source File with Exact Column Placement ---
    try:
        if original_file_path.endswith('.csv'):
            master_df = pd.read_csv(original_file_path, dtype=str)
        else:
            master_df = pd.read_excel(original_file_path, dtype=str)
    except Exception:
        return created_files # Cannot proceed if original file is unreadable

    # Get the original column names for unit ID and AMI
    unit_col = original_headers.get('unit_id')
    ami_col = original_headers.get('client_ami')

    if not unit_col or not ami_col or unit_col not in master_df.columns or ami_col not in master_df.columns:
        # If essential columns are missing, cannot create the updated source file
        return created_files

    # Find the integer position of the client's AMI column
    ami_col_idx = master_df.columns.get_loc(ami_col)

    # Prepare scenario data for insertion
    s1_data = scenarios_to_report.get("S1_Absolute_Best")
    s2_data = scenarios_to_report.get("S2_Alternative")

    # Use a mapping approach to preserve original file's row order and non-affordable units
    if s1_data and 'assignments' in s1_data:
        s1_map = {str(u['unit_id']): f"{int(u['assigned_ami']*100)}%" for u in s1_data['assignments']}
        master_df.insert(ami_col_idx + 1, 'AMI_S1', master_df[unit_col].map(s1_map).fillna(''))

    if s2_data and 'assignments' in s2_data:
        s2_map = {str(u['unit_id']): f"{int(u['assigned_ami']*100)}%" for u in s2_data['assignments']}
        # Insert S2 next to S1 if S1 was added, otherwise next to the original AMI column
        insert_idx = ami_col_idx + 2 if 'AMI_S1' in master_df.columns else ami_col_idx + 1
        master_df.insert(insert_idx, 'AMI_S2', master_df[unit_col].map(s2_map).fillna(''))

    updated_source_filepath = os.path.join(output_dir, f"{base_name}_Updated_Source.xlsx")
    master_df.to_excel(updated_source_filepath, index=False)
    created_files.append(updated_source_filepath)

    return created_files
