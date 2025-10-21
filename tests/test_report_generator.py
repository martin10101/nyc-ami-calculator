import pandas as pd
import pytest
from pathlib import Path

from ami_optix.report_generator import create_excel_reports


def test_create_excel_reports_adds_scenario_columns(tmp_path):
    original_path = tmp_path / "input.xlsx"
    original_df = pd.DataFrame({
        'APT': ['1A', '1B'],
        'AMI': ['50%', '60%'],
        'NET SF': [500, 700],
        'FLOOR': [1, 2],
    })
    original_df.to_excel(original_path, index=False)

    scenario_template = {
        'status': 'OPTIMAL',
        'bands': [40, 60, 80],
        'assignments': [
            {'unit_id': '1A', 'bedrooms': 1, 'net_sf': 500, 'floor': 1, 'client_ami': 0.5, 'premium_score': 0.1, 'assigned_ami': 0.4},
            {'unit_id': '1B', 'bedrooms': 2, 'net_sf': 700, 'floor': 2, 'client_ami': 0.6, 'premium_score': 0.2, 'assigned_ami': 0.6},
        ],
        'metrics': {
            'waami_percent': 50.0,
            'total_units': 2,
            'total_sf': 1200,
            'revenue_score': 550.0,
            'band_mix': [
                {'band': 40, 'units': 1, 'net_sf': 500, 'share_of_sf': 500/1200},
                {'band': 60, 'units': 1, 'net_sf': 700, 'share_of_sf': 700/1200},
            ],
        },
    }

    analysis_json = {
        'scenario_absolute_best': scenario_template,
        'scenario_client_oriented': scenario_template,
        'scenario_best_3_band': scenario_template,
        'scenario_best_2_band': {**scenario_template, 'bands': [40, 80]},
        'scenario_alternative': {**scenario_template, 'bands': [50, 90]},
        'scenarios': {
            'absolute_best': scenario_template,
            'client_oriented': scenario_template,
            'best_3_band': scenario_template,
            'best_2_band': {**scenario_template, 'bands': [40, 80]},
            'alternative': {**scenario_template, 'bands': [50, 90]},
        },
    }

    files = create_excel_reports(
        analysis_json,
        str(original_path),
        {'unit_id': 'APT', 'client_ami': 'AMI'},
        output_dir=str(tmp_path)
    )

    updated_path = [p for p in files if p.endswith('Updated_Source.xlsx')][0]
    units_df = pd.read_excel(updated_path, sheet_name='Units')
    assert 'AMI_S1_Absolute_Best' in units_df.columns
    assert 'AMI_S2_Client_Oriented' in units_df.columns
    summary_df = pd.read_excel(updated_path, sheet_name='Scenario Summary')
    assert not summary_df.empty


def test_create_excel_reports_prefers_xlsb_when_available(tmp_path, monkeypatch):
    original_path = tmp_path / "input.xlsx"
    pd.DataFrame({'APT': ['1A'], 'AMI': ['40%'], 'NET SF': [500]}).to_excel(original_path, index=False)

    analysis_json = {
        'analysis_notes': [],
        'scenario_absolute_best': {
            'assignments': [{'unit_id': '1A', 'net_sf': 500, 'assigned_ami': 0.4, 'premium_score': 0.1}],
            'metrics': {'band_mix': [], 'waami_percent': 40.0},
            'bands': [40],
        },
        'scenarios': {
            'absolute_best': {
                'assignments': [{'unit_id': '1A', 'net_sf': 500, 'assigned_ami': 0.4, 'premium_score': 0.1}],
                'metrics': {'band_mix': [], 'waami_percent': 40.0},
                'bands': [40],
            }
        },
    }

    def fake_convert(path, target_path=None, delete_source=True):
        src = Path(path)
        target = Path(target_path) if target_path else src.with_suffix('.xlsb')
        target.write_bytes(b'dummy')
        if delete_source and src.exists():
            src.unlink()
        return str(target)

    monkeypatch.setattr('ami_optix.report_generator.convert_xlsx_to_xlsb', fake_convert)

    files = create_excel_reports(
        analysis_json,
        str(original_path),
        {'unit_id': 'APT', 'client_ami': 'AMI'},
        output_dir=str(tmp_path),
        prefer_xlsb=True,
    )

    assert any(file.endswith('.xlsb') for file in files)
    assert all(not file.endswith('.xlsx') for file in files)


def test_create_excel_reports_falls_back_to_xlsx_when_conversion_fails(tmp_path, monkeypatch):
    original_path = tmp_path / "input.xlsx"
    pd.DataFrame({'APT': ['1A'], 'AMI': ['40%'], 'NET SF': [500]}).to_excel(original_path, index=False)

    analysis_json = {
        'analysis_notes': [],
        'scenario_absolute_best': {
            'assignments': [{'unit_id': '1A', 'net_sf': 500, 'assigned_ami': 0.4, 'premium_score': 0.1}],
            'metrics': {'band_mix': [], 'waami_percent': 40.0},
            'bands': [40],
        },
        'scenarios': {
            'absolute_best': {
                'assignments': [{'unit_id': '1A', 'net_sf': 500, 'assigned_ami': 0.4, 'premium_score': 0.1}],
                'metrics': {'band_mix': [], 'waami_percent': 40.0},
                'bands': [40],
            }
        },
    }

    def failing_convert(*_args, **_kwargs):
        raise RuntimeError("conversion failed")

    monkeypatch.setattr('ami_optix.report_generator.convert_xlsx_to_xlsb', failing_convert)

    files = create_excel_reports(
        analysis_json,
        str(original_path),
        {'unit_id': 'APT', 'client_ami': 'AMI'},
        output_dir=str(tmp_path),
        prefer_xlsb=True,
    )

    assert any(file.endswith('.xlsx') for file in files)
    assert any('Delivered as .xlsx' in note for note in analysis_json['analysis_notes'])
