import pandas as pd

# As defined in the Project Charter, Section 3.2
HEADER_MAPPING = {
    "unit_id": ["APT", "UNIT", "UNIT ID", "APARTMENT", "APT #"],
    "bedrooms": ["BED", "BEDS", "BEDROOMS"],
    "net_sf": ["NET SF", "NETSF", "SF", "S.F.", "SQFT", "SQ FT", "AREA"],
    "floor": ["FLOOR", "STORY", "LEVEL"],
    "balcony": ["BALCONY", "TERRACE", "OUTDOOR"],
    "client_ami": ["AMI", "AFFORDABILITY", "AFF %", "AFF", "AMI_INPUT"],
}

class Parser:
    """
    The Parser is the "Prep Cook" of the AMI-Optix system.
    Its sole responsibility is to read, validate, and prepare the input data
    from a client's spreadsheet (.xlsx or .csv). It ensures the data is 100%
    correct before any analysis begins.
    """
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = None
        self.mapped_headers = {}

    def read_data(self):
        """
        Reads the input file (Excel or CSV) into a pandas DataFrame.
        """
        if not self.file_path.endswith(('.csv', '.xlsx')):
            raise ValueError("Unsupported file type. Please provide a .csv or .xlsx file.")

        try:
            if self.file_path.endswith('.csv'):
                self.data = pd.read_csv(self.file_path)
            else:  # .xlsx
                self.data = pd.read_excel(self.file_path)

            # Normalize column headers to lowercase for easier matching
            self.data.columns = [str(col).lower() for col in self.data.columns]
        except FileNotFoundError:
            raise FileNotFoundError(f"Error: The file '{self.file_path}' was not found.")
        except Exception as e:
            raise IOError(f"Error reading or parsing the file '{self.file_path}': {e}")

    def _find_column(self, possible_names):
        """
        Finds the first matching column in the DataFrame from a list of possible names.
        """
        for name in possible_names:
            if name.lower() in self.data.columns:
                return name.lower()
        return None

    def map_headers(self):
        """
        Maps the fuzzy headers from the input file to the standardized internal names.
        """
        original_columns = list(self.data.columns)
        self.data.columns = [str(col).lower() for col in self.data.columns]

        for key, possible_names in HEADER_MAPPING.items():
            for name in possible_names:
                if name.lower() in self.data.columns:
                    # Find the original cased version of the header
                    original_header = next((h for h in original_columns if h.lower() == name.lower()), None)
                    self.mapped_headers[key] = original_header
                    break

        # Restore original column names for data manipulation
        self.data.columns = original_columns

        # Validate that essential columns were found
        required_columns = ["unit_id", "bedrooms", "net_sf"]
        for col in required_columns:
            if col not in self.mapped_headers:
                raise ValueError(f"Error: Missing required column. Could not find a match for '{col}' (e.g., {HEADER_MAPPING[col][0]}).")

        return self.mapped_headers

    def _get_standardized_dataframe(self):
        """
        Internal method to get a DataFrame with standardized column names.
        """
        # Create a new DataFrame with only the mapped columns and standardized names
        standardized_df = pd.DataFrame()
        for standard_name, original_name in self.mapped_headers.items():
            standardized_df[standard_name] = self.data[original_name]
        return standardized_df

    def get_affordable_units(self):
        """
        The main public method for the Parser. It reads the file, maps headers,
        identifies the affordable unit set, validates their data, and returns
        a clean DataFrame of units to be optimized.
        """
        if self.data is None:
            self.read_data()
        if not self.mapped_headers:
            self.map_headers()

        df = self._get_standardized_dataframe()

        # 1. Identify the Affordable Set
        if 'client_ami' not in df.columns:
            raise ValueError("Error: The 'Client AMI' column (e.g., 'AMI', 'AFF %') is required to identify affordable units, but it was not found.")

        # Clean and convert the client_ami column, handling percentages and decimals
        ami_series = df['client_ami'].astype(str).str.strip()
        is_percent = ami_series.str.contains('%', na=False)

        numeric_vals = pd.to_numeric(ami_series.str.replace('%', '', regex=False), errors='coerce')

        # Where it was a percentage, divide by 100 to normalize
        numeric_vals[is_percent] = numeric_vals[is_percent] / 100.0

        df['client_ami'] = numeric_vals

        # Filter for rows where client_ami is a positive number
        affordable_df = df[df['client_ami'] > 0].copy()

        if affordable_df.empty:
            raise ValueError("Error: No affordable units found. Please ensure the 'Client AMI' column has positive numerical values for the units to be included.")

        # 2. Validate Critical Data for the affordable set
        # Check for null or empty Unit IDs first
        # Convert to string and strip whitespace to handle various forms of "empty"
        affordable_df['unit_id'] = affordable_df['unit_id'].astype(str).str.strip()
        if affordable_df['unit_id'].isnull().any() or (affordable_df['unit_id'] == '').any():
            invalid_rows = affordable_df[affordable_df['unit_id'].isnull() | (affordable_df['unit_id'] == '')]
            raise ValueError(f"Error: Found affordable units with a missing Unit ID in rows: {invalid_rows.index.tolist()}")

        for col in ['bedrooms', 'net_sf']:
            affordable_df[col] = pd.to_numeric(affordable_df[col], errors='coerce')
            if affordable_df[col].isnull().any():
                invalid_rows = affordable_df[affordable_df[col].isnull()]
                unit_ids = invalid_rows['unit_id'].tolist()
                raise ValueError(f"Error: Invalid non-numeric data found in required column '{col}' for units: {unit_ids}")

            if (affordable_df[col] < 0).any():
                invalid_rows = affordable_df[affordable_df[col] < 0]
                unit_ids = invalid_rows['unit_id'].tolist()
                raise ValueError(f"Error: Column '{col}' must contain non-negative values. Found negative values for units: {unit_ids}")

        return affordable_df
