import pytest
import pandas as pd
import os
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
    assert "net_sf" in str(excinfo.value)

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

def test_header_with_leading_trailing_whitespace(temp_dir):
    """Tests that headers with surrounding whitespace are correctly mapped."""
    # Note the leading/trailing spaces around the header names
    headers = ['  APT #  ', 'BEDS', '  net sf  ', 'AFF %']
    data = [
        ['1A', 1, 600, 0.8],
    ]
    filepath = create_csv(temp_dir, "whitespace_headers.csv", headers, data)

    parser = Parser(filepath)
    # This should not raise an error
    affordable_df = parser.get_affordable_units()

    assert len(affordable_df) == 1
    assert 'net_sf' in affordable_df.columns
    assert affordable_df.iloc[0]['net_sf'] == 600
