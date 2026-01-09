import pytest
import pandas as pd
import os
from openpyxl import Workbook

from ami_optix.parser import Parser

@pytest.fixture
def temp_dir(tmp_path):
    """Create a temporary directory for test files."""
    return tmp_path

def create_csv(directory, filename, headers, data):
    """Helper function to create a CSV file in a given directory."""
    filepath = os.path.join(directory, filename)
    df = pd.DataFrame(data, columns=headers)
    df.to_csv(filepath, index=False)
    return filepath

def test_get_affordable_units_success(temp_dir):
    """Tests successful identification and validation of affordable units."""
    headers = ['APT #', 'BEDS', 'SQFT', 'AFF %']
    data = [
        ['1A', 1, 600, 0.8],   # Affordable
        ['1B', 2, 800, 1.2],   # Affordable
        ['2A', 0, 450, ''],    # Not affordable (empty AMI)
        ['2B', 1, 650, 0],     # Not affordable (zero AMI)
        ['3A', 2, 850, 'Market'],# Not affordable (non-numeric AMI)
        ['3B', 3, 1100, 1.2],  # Affordable
    ]
    filepath = create_csv(temp_dir, "mixed_units.csv", headers, data)

    parser = Parser(filepath)
    affordable_df = parser.get_affordable_units()

    assert len(affordable_df) == 3
    assert affordable_df['unit_id'].tolist() == ['1A', '1B', '3B']
    assert 'bedrooms' in affordable_df.columns
    assert 'net_sf' in affordable_df.columns
    assert 'client_ami' in affordable_df.columns

def test_missing_ami_column(temp_dir):
    """Tests that an error is raised if the AMI column is missing."""
    headers = ['APT #', 'BEDS', 'SQFT']
    data = [['1A', 1, 600]]
    filepath = create_csv(temp_dir, "no_ami.csv", headers, data)

    parser = Parser(filepath)
    with pytest.raises(ValueError) as excinfo:
        parser.get_affordable_units()
    assert "The 'Client AMI' column" in str(excinfo.value)

def test_no_affordable_units_found(temp_dir):
    """Tests that an error is raised if no units have a valid AMI value."""
    headers = ['APT #', 'BEDS', 'SQFT', 'AFF %']
    data = [['1A', 1, 600, 0], ['1B', 2, 800, 'N/A']]
    filepath = create_csv(temp_dir, "none_affordable.csv", headers, data)

    parser = Parser(filepath)
    with pytest.raises(ValueError) as excinfo:
        parser.get_affordable_units()
    assert "No affordable units found" in str(excinfo.value)

def test_invalid_data_in_required_column(temp_dir):
    """Tests for non-numeric data in critical columns of affordable units."""
    headers = ['APT #', 'BEDS', 'SQFT', 'AFF %']
    data = [
        ['1A', 1, 600, 0.8],
        ['1B', 'two', 800, 1.2], # Invalid 'BEDS'
    ]
    filepath = create_csv(temp_dir, "invalid_beds.csv", headers, data)

    parser = Parser(filepath)
    with pytest.raises(ValueError) as excinfo:
        parser.get_affordable_units()
    assert "Invalid non-numeric data" in str(excinfo.value)
    assert "'bedrooms'" in str(excinfo.value)
    assert "1B" in str(excinfo.value)

def test_non_positive_data_in_required_column(temp_dir):
    """Tests for zero or negative data in critical columns."""
    headers = ['APT #', 'BEDS', 'SQFT', 'AFF %']
    data = [
        ['1A', 1, 600, 0.8],
        ['1B', 2, -100, 1.2], # Invalid 'SQFT'
    ]
    filepath = create_csv(temp_dir, "invalid_sqft.csv", headers, data)

    parser = Parser(filepath)
    with pytest.raises(ValueError) as excinfo:
        parser.get_affordable_units()
    assert "must contain non-negative values" in str(excinfo.value)
    assert "'net_sf'" in str(excinfo.value)
    assert "1B" in str(excinfo.value)

def test_missing_required_header(temp_dir):
    """Tests if the parser raises ValueError for a missing required header like 'SQFT'."""
    headers = ['APT #', 'BEDS', 'AFF %'] # SQFT is missing
    data = [['1A', 1, 0.8]]
    filepath = create_csv(temp_dir, "missing_header.csv", headers, data)

    parser = Parser(filepath)
    with pytest.raises(ValueError) as excinfo:
        parser.get_affordable_units()
    assert "Missing required column" in str(excinfo.value)
    assert "'net_sf'" in str(excinfo.value)

def test_missing_unit_id_value(temp_dir):
    """Tests that an error is raised for affordable units with empty unit IDs."""
    headers = ['APT #', 'BEDS', 'SQFT', 'AFF %']
    data = [
        ['1A', 1, 600, 0.8],
        [' ', 2, 800, 1.2], # Invalid Unit ID
    ]
    filepath = create_csv(temp_dir, "missing_id.csv", headers, data)

    parser = Parser(filepath)
    with pytest.raises(ValueError) as excinfo:
        parser.get_affordable_units()
    assert "missing Unit ID" in str(excinfo.value)

def test_ami_column_with_percentages(temp_dir):
    """Tests that the parser correctly handles AMI values as percentages and decimals."""
    headers = ['unit id', 'bedrooms', 'net sf', 'AMI'] # Use a header from the mapping
    data = [
        ['1A', 1, 600, '60%'],      # Percentage
        ['1B', 2, 800, ' 80 % '],   # Percentage with whitespace
        ['2A', 0, 450, 0.5],        # Decimal
        ['2B', 1, 650, 'N/A'],      # Not affordable
    ]
    filepath = create_csv(temp_dir, "percentages.csv", headers, data)

    parser = Parser(filepath)
    df = parser.get_affordable_units()

    assert len(df) == 3
    assert df.loc[df['unit_id'] == '1A', 'client_ami'].iloc[0] == 0.6
    assert df.loc[df['unit_id'] == '1B', 'client_ami'].iloc[0] == 0.8
    assert df.loc[df['unit_id'] == '2A', 'client_ami'].iloc[0] == 0.5

def test_headers_with_whitespace_variations(temp_dir):
    """Parser should tolerate leading/trailing spaces and tabs in headers."""
    headers = ['FLOOR', 'APT', 'BED', ' NET SF', 'AMI', 'BALCONY']
    data = [
        [2, '2A', 1, 450.0, '60%', ''],
        [2, '2B', 2, 595.0, '75%', ''],
        [3, '3C', 1, 407.0, '60%', 'Yes'],
    ]
    filepath = create_csv(temp_dir, "whitespace_headers.csv", headers, data)

    parser = Parser(filepath)
    df = parser.get_affordable_units()

    assert len(df) == 3
    assert df['net_sf'].tolist() == [450.0, 595.0, 407.0]
    assert df['client_ami'].tolist()[0] == 0.60


def test_ami_for_35_years_header(temp_dir):
    """Parser should recognize HPD-style 'AMI FOR 35 Years' headers."""
    headers = ['UNIT', 'BEDS', 'SQ. FT.', 'AMI FOR 35 Years']
    data = [
        ['101', 0, 412, 0.6],
        ['201', 1, 511, 0.8],
        ['301', 2, 750, ''],
    ]
    filepath = create_csv(temp_dir, "ami_for_35_years.csv", headers, data)

    parser = Parser(filepath)
    df = parser.get_affordable_units()

    assert df['client_ami'].tolist() == [0.6, 0.8]


def test_ami_header_with_extra_text(temp_dir):
    """Parser should fall back to partial matches when headers contain extra text."""
    headers = ['UNIT', 'BEDS', 'SQ. FT.', 'AMI requirement (HPD 35 Years)']
    data = [
        ['101', 0, 412, '60%'],
        ['201', 1, 511, '80%'],
        ['301', 2, 750, ''],
    ]
    filepath = create_csv(temp_dir, "ami_with_extra_text.csv", headers, data)

    parser = Parser(filepath)
    df = parser.get_affordable_units()

    assert df['client_ami'].tolist() == [0.6, 0.8]


def test_client_ami_forward_fills_for_unit_rows(temp_dir):
    """Parser should inherit AMI values for units when the workbook uses merged cells."""
    headers = ['UNIT', 'BEDS', 'SQ. FT.', 'AMI FOR 35 Years', 'AMI AFTER 35 YEARS']
    data = [
        ['Commercial Retail', '', '', '', ''],
        ['101', 0, 412, '40%', 'After'],
        ['102', 0, 405, '', ''],
        ['103', 1, 550, '', ''],
        ['201', 1, 600, '60%', 'After'],
        ['202', 1, 620, '', ''],
    ]
    filepath = create_csv(temp_dir, "ami_forward_fill.csv", headers, data)

    parser = Parser(filepath)
    df = parser.get_affordable_units()

    assert df['unit_id'].tolist() == ['101', '102', '103', '201', '202']
    assert df['client_ami'].tolist() == [0.4, 0.4, 0.4, 0.6, 0.6]


def test_project_worksheet_unit_table_is_preferred(tmp_path):
    """Parser should prefer the PROJECT WORKSHEET unit table when available."""

    path = tmp_path / "project_worksheet.xlsx"

    wb = Workbook()

    ws = wb.active
    ws.title = "PROJECT WORKSHEET"

    ws.append([])
    ws.append([])
    ws.append(["UNIT INFORMATION"])
    ws.append(
        [
            "Unit No.",
            "Building Segment #",
            "Construction Story",
            "Marketing Story",
            "Apt #",
            "Number of Bedrooms",
            "Net Square Feet",
            "Affordable Housing Unit",
            "Affordable Housing Unit AMI Band",
        ]
    )

    unit_rows = [
        (1, None, 2, 2, "201", 1, 459.38, "YES", "60%"),
        (2, None, 2, 2, "202", 1, 480.25, "NO", ""),
        (3, None, 2, 2, "203", 2, 752.00, "YES", "60%"),
        (4, None, 3, 3, "301", 1, 459.38, "YES", "80%"),
        (5, None, 3, 3, "302", 2, 752.00, "NO", ""),
        (6, None, 3, 3, "303", 1, 490.58, "YES", "40%"),
    ]

    for row in unit_rows:
        ws.append(row)

    ws.append(["Totals", None, None, None, None, None, None, None, None])

    rentroll = wb.create_sheet("RentRoll")
    rentroll.append(["UNIT", "BEDS", "NET SF", "AMI FOR 35 Years"])
    rentroll.append(["201", 0, 400, "60%"])

    wb.save(path)

    parser = Parser(str(path))
    df = parser.get_affordable_units()

    assert sorted(df["unit_id"].tolist()) == ["201", "203", "301", "303"]
    assert df["bedrooms"].tolist() == [1.0, 2.0, 1.0, 1.0]
    assert df["client_ami"].tolist() == [0.6, 0.6, 0.8, 0.4]

