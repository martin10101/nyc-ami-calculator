import pandas as pd
import os

def create_excel_reports(analysis_json, original_file_path, output_dir='reports'):
    """
    Generates multiple Excel reports from the solver's analysis JSON.

    Args:
        analysis_json (dict): The JSON output from the AMI-Optix solver.
        original_file_path (str): The path to the user's original uploaded file.
        output_dir (str): The directory to save the reports in.

    Returns:
        list: A list of file paths for the generated reports.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    base_name = os.path.splitext(os.path.basename(original_file_path))[0]
    created_files = []

    # --- 1. Create Individual Scenario Reports ---
    scenarios_to_report = {
        "Absolute_Best": analysis_json.get("scenario_absolute_best"),
        "Alternative": analysis_json.get("scenario_alternative"),
        "Client_Oriented": analysis_json.get("scenario_client_oriented"),
        "Best_2_Band": analysis_json.get("scenario_best_2_band"),
    }

    for name, scenario in scenarios_to_report.items():
        if not scenario:
            continue

        df = pd.DataFrame(scenario['assignments'])
        # Select and reorder columns for a clean report
        report_df = df[['unit_id', 'bedrooms', 'net_sf', 'floor', 'assigned_ami']].copy()
        report_df.rename(columns={
            'unit_id': 'Unit ID',
            'bedrooms': 'Bedrooms',
            'net_sf': 'Net SF',
            'floor': 'Floor',
            'assigned_ami': 'Assigned AMI'
        }, inplace=True)

        # Format the AMI column as a percentage
        report_df['Assigned AMI'] = (report_df['Assigned AMI'] * 100).map('{:.0f}%'.format)

        filepath = os.path.join(output_dir, f"{base_name}_{name}_Report.xlsx")
        report_df.to_excel(filepath, index=False)
        created_files.append(filepath)

    # --- 2. Create Master Report ---
    # Read the original file to preserve its structure
    try:
        if original_file_path.endswith('.csv'):
            master_df = pd.read_csv(original_file_path)
        else:
            master_df = pd.read_excel(original_file_path)
    except Exception:
        # If original file can't be read, skip master report
        return created_files

    # Standardize a unit ID column for merging
    # This assumes the fuzzy header mapping has already identified the unit column
    # For simplicity, we find it again here.
    unit_id_headers = ["APT", "UNIT", "UNIT ID", "APARTMENT", "APT #"]
    original_unit_col = next((col for col in master_df.columns if col.upper() in unit_id_headers), None)

    if original_unit_col:
        # Ensure the merge key is the same data type
        master_df[original_unit_col] = master_df[original_unit_col].astype(str)

        for name, scenario in scenarios_to_report.items():
            if not scenario:
                continue

            scenario_df = pd.DataFrame(scenario['assignments'])[['unit_id', 'assigned_ami']]
            scenario_df['unit_id'] = scenario_df['unit_id'].astype(str)

            # Format as percentage string
            scenario_df['assigned_ami'] = (scenario_df['assigned_ami'] * 100).map('{:.0f}%'.format)

            # Merge the assignments into the master dataframe
            master_df = pd.merge(
                master_df,
                scenario_df,
                how='left',
                left_on=original_unit_col,
                right_on='unit_id'
            ).rename(columns={'assigned_ami': f'AMI_{name}'}).drop('unit_id', axis=1)

    master_filepath = os.path.join(output_dir, f"{base_name}_Master_Report.xlsx")
    master_df.to_excel(master_filepath, index=False)
    created_files.append(master_filepath)

    return created_files
