import os
from pathlib import Path
from typing import Dict
import pandas as pd

from ami_optix.rent_calculator import save_rent_workbook_with_utilities
_SCENARIO_EXPORT_ORDER = [
    ("S1_Absolute_Best", "scenario_absolute_best", "AMI_S1_Absolute_Best"),
    ("S2_Client_Oriented", "scenario_client_oriented", "AMI_S2_Client_Oriented"),
    ("S3_Best_3_Band", "scenario_best_3_band", "AMI_S3_Best_3_Band"),
    ("S4_Best_2_Band", "scenario_best_2_band", "AMI_S4_Best_2_Band"),
    ("S5_Alternative", "scenario_alternative", "AMI_S5_Alternative"),
]


def _pandas_engine_for(path: str):
    ext = os.path.splitext(path.lower())[1]
    if ext == ".xlsb":
        return "pyxlsb"
    return None


def _scenario_to_dataframe(scenario):
    df = pd.DataFrame(scenario['assignments'])
    base_columns = {
        'unit_id': 'Unit ID',
        'bedrooms': 'Bedrooms',
        'net_sf': 'Net SF',
        'floor': 'Floor',
        'assigned_ami': 'Assigned AMI',
    }
    available_cols = [col for col in base_columns if col in df.columns]
    report_df = df[available_cols].copy()
    report_df.rename(columns={col: base_columns[col] for col in available_cols}, inplace=True)
    if 'Assigned AMI' in report_df.columns:
        report_df['Assigned AMI'] = (report_df['Assigned AMI'].astype(float) * 100).map('{:.0f}%'.format)

    if 'gross_rent' in df.columns:
        report_df['Gross Rent'] = df['gross_rent'].round(2)
    if 'monthly_rent' in df.columns:
        report_df['Monthly Rent'] = df['monthly_rent'].round(2)
    if 'allowance_total' in df.columns:
        report_df['Rent Deductions'] = df['allowance_total'].round(2)
    if 'annual_rent' in df.columns:
        report_df['Annual Rent'] = df['annual_rent'].round(2)
    if 'allowances' in df.columns:
        def _format_allowances(items):
            if not isinstance(items, list):
                return ''
            formatted = [
                f"{detail.get('label') or detail.get('category')}: -${detail.get('amount', 0):,.2f}"
                for detail in items
                if detail and detail.get('amount')
            ]
            return '; '.join(formatted)
        report_df['Allowance Detail'] = df['allowances'].apply(_format_allowances)
    return report_df


def _scenario_summary_frame(display_name, scenario):
    metrics = scenario.get('metrics', {})
    band_mix = metrics.get('band_mix', [])
    band_summary = "; ".join(
        f"{entry['band']}%: {entry['units']} units ({entry['share_of_sf']*100:.1f}% SF)"
        for entry in band_mix
    )
    allowance_breakdown = metrics.get('allowance_breakdown', {})
    allowance_summary = '; '.join(
        f"{detail.get('label') or category}: -${detail.get('monthly', 0):,.2f}"
        for category, detail in allowance_breakdown.items()
        if detail.get('monthly')
    )
    summary = {
        'Scenario': display_name.replace('_', ' '),
        'WAAMI (%)': round(metrics.get('waami_percent', scenario.get('waami', 0) * 100), 4),
        'Bands Used': ", ".join(f"{band}%" for band in scenario.get('bands', [])),
        'Total Units': metrics.get('total_units'),
        'Total SF': metrics.get('total_sf'),
        'Revenue Score': round(metrics.get('revenue_score', 0.0), 2),
        'Band Mix': band_summary,
        '40% Units': metrics.get('low_band_units'),
        '40% SF': round(metrics.get('low_band_sf', 0.0), 2),
        '40% Share (%)': round(metrics.get('low_band_share', 0.0) * 100, 2),
        'Gross Monthly Rent': round(metrics.get('gross_monthly_rent', 0.0), 2),
        'Gross Annual Rent': round(metrics.get('gross_annual_rent', 0.0), 2),
        'Rent Deductions (Monthly)': round(metrics.get('allowance_monthly_total', 0.0), 2),
        'Rent Deductions Detail': allowance_summary,
        'Total Monthly Rent': round(metrics.get('total_monthly_rent', 0.0), 2),
        'Total Annual Rent': round(metrics.get('total_annual_rent', 0.0), 2),
    }
    return pd.DataFrame([summary])


def create_excel_reports(
    analysis_json,
    original_file_path,
    original_headers,
    output_dir='reports',
    prefer_xlsb: bool = False,
    utilities: Dict[str, str] | None = None,
    rent_workbook_path: str | None = None,
):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    base_name = os.path.splitext(os.path.basename(original_file_path))[0]
    created_files = []

    scenario_lookup = {
        key: analysis_json.get(key)
        for _, key, _ in _SCENARIO_EXPORT_ORDER
    }

    # --- Individual scenario workbooks ---
    notes = analysis_json.setdefault('analysis_notes', [])
    xlsb_note_added = False

    def _register_xlsx(path_xlsx: str) -> str:
        nonlocal xlsb_note_added
        created_files.append(path_xlsx)
        if prefer_xlsb and not xlsb_note_added:
            notes.append(
                "XLSB export requested, but the current environment only delivers .xlsx files. "
                "Please reopen on a Windows host to receive .xlsb downloads."
            )
            xlsb_note_added = True
        return path_xlsx

    for display_name, analysis_key, _ in _SCENARIO_EXPORT_ORDER:
        scenario = scenario_lookup.get(analysis_key)
        if not scenario or 'assignments' not in scenario:
            continue
        filepath_xlsx = os.path.join(output_dir, f"{base_name}_{display_name}_Report.xlsx")
        with pd.ExcelWriter(filepath_xlsx, engine='xlsxwriter') as writer:
            _scenario_to_dataframe(scenario).to_excel(writer, sheet_name='Assignments', index=False)
            _scenario_summary_frame(display_name, scenario).to_excel(writer, sheet_name='Summary', index=False)
        _register_xlsx(filepath_xlsx)

    if utilities and rent_workbook_path:
        rent_ext = Path(rent_workbook_path).suffix.lower()
        rent_output_name = f"{base_name}_Rent_Calculator{rent_ext if rent_ext in {'.xlsx', '.xlsm'} else '.xlsx'}"
        rent_output_path = os.path.join(output_dir, rent_output_name)
        try:
            save_rent_workbook_with_utilities(rent_workbook_path, utilities, rent_output_path)
            created_files.append(rent_output_path)
        except ValueError as exc:
            notes.append(f"Rent workbook export warning: {exc}")
        except FileNotFoundError:
            notes.append("Rent workbook export warning: provided rent calculator file could not be found.")
        except Exception as exc:  # pragma: no cover - defensive
            notes.append(f"Rent workbook export warning: {exc}")

    # --- Updated source workbook ---
    try:
        if original_file_path.endswith('.csv'):
            master_df = pd.read_csv(original_file_path, dtype=str)
        else:
            engine = _pandas_engine_for(original_file_path)
            master_df = pd.read_excel(original_file_path, dtype=str, engine=engine)
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

    updated_source_filepath_xlsx = os.path.join(output_dir, f"{base_name}_Updated_Source.xlsx")
    with pd.ExcelWriter(updated_source_filepath_xlsx, engine='xlsxwriter') as writer:
        master_df.to_excel(writer, sheet_name='Units', index=False)
        if not summary_sheet.empty:
            summary_sheet.to_excel(writer, sheet_name='Scenario Summary', index=False)
    _register_xlsx(updated_source_filepath_xlsx)

    return created_files
