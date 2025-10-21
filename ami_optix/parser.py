import os
import pandas as pd
import re

# As defined in the Project Charter, Section 3.2
HEADER_MAPPING = {
    "unit_id": ["APT", "UNIT", "UNIT ID", "APARTMENT", "APT #"],
    "bedrooms": ["BED", "BEDS", "BEDROOMS"],
    "net_sf": ["NET SF", "NETSF", "SF", "S.F.", "SQFT", "SQ FT", "AREA"],
    "floor": ["FLOOR", "STORY", "LEVEL"],
    "balcony": ["BALCONY", "TERRACE", "OUTDOOR"],
    "client_ami": ["AMI", "AFFORDABILITY", "AFF %", "AFF", "AMI_INPUT"],
}


def _normalize_header(value):
    """Return a normalized representation of a header for fuzzy matching."""
    if value is None:
        return ""

    normalized = str(value).replace("\xa0", " ")
    normalized = normalized.replace(".", " ").replace("-", " ")
    normalized = re.sub(r"[_]+", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized.strip().lower())
    return normalized


class Parser:
    """
    The Parser is the "Prep Cook" of the AMI-Optix system.
    Its sole responsibility is to read, validate, and prepare the input data
    from a client's spreadsheet. It ensures the data is 100% correct before
    any analysis begins.
    """

    def __init__(self, file_path):
        self.file_path = file_path
        self.data = None
        self.mapped_headers = {}

    def read_data(self):
        """Reads the input file (Excel or CSV) into a pandas DataFrame."""
        _, ext = os.path.splitext(self.file_path.lower())
        if ext not in (".csv", ".xlsx", ".xlsm", ".xlsb"):
            raise ValueError("Unsupported file type. Please provide a .csv, .xlsx, .xlsm, or .xlsb file.")

        try:
            if ext == ".csv":
                self.data = pd.read_csv(self.file_path)
            else:
                self.data = self._read_excel_with_fallback(ext)
        except FileNotFoundError:
            raise FileNotFoundError(f"Error: The file '{self.file_path}' was not found.")
        except Exception as e:
            raise IOError(f"Error reading or parsing the file '{self.file_path}': {e}")

    def _read_excel_with_fallback(self, ext: str) -> pd.DataFrame:
        engine = "pyxlsb" if ext == ".xlsb" else None
        excel = pd.ExcelFile(self.file_path, engine=engine)

        preferred = ["RentRoll", "Units", "Sheet1"]
        ordered = preferred + [name for name in excel.sheet_names if name not in preferred]

        for sheet in ordered:
            # Try direct parse with default header
            try:
                direct_df = excel.parse(sheet)
                if self._sheet_has_viable_headers(direct_df.columns):
                    return direct_df
            except Exception:
                pass

            # Fallback: search for header row manually
            try:
                raw = excel.parse(sheet, header=None)
            except Exception:
                continue

            header_idx = self._locate_header_row(raw)
            if header_idx is None:
                continue

            header_row = raw.iloc[header_idx].fillna("")
            df = raw.iloc[header_idx + 1 :].copy()
            df.columns = header_row
            df.dropna(how="all", inplace=True)
            if not df.empty and self._sheet_has_viable_headers(df.columns):
                return df

        raise ValueError("Unable to locate a worksheet with recognizable columns for unit data.")

    def _sheet_has_viable_headers(self, columns) -> bool:
        normalized = {_normalize_header(col) for col in columns if col is not None}
        required = {_normalize_header(name) for name in HEADER_MAPPING["unit_id"]}
        bedrooms = {_normalize_header(name) for name in HEADER_MAPPING["bedrooms"]}
        net_sf = {_normalize_header(name) for name in HEADER_MAPPING["net_sf"]}
        return bool(normalized & required) and bool(normalized & bedrooms) and bool(normalized & net_sf)

    def _locate_header_row(self, dataframe: pd.DataFrame):
        for idx in range(len(dataframe)):
            row = dataframe.iloc[idx]
            normalized_values = {_normalize_header(val) for val in row if pd.notna(val)}
            hits = 0
            for names in HEADER_MAPPING.values():
                if any(_normalize_header(name) in normalized_values for name in names):
                    hits += 1
            if hits >= 3:
                return idx
        return None

    def map_headers(self):
        """Maps fuzzy headers from the input file to standardized internal names."""
        if self.data is None:
            self.read_data()

        original_columns = list(self.data.columns)
        normalized_to_original = {}
        for original in original_columns:
            normalized = _normalize_header(original)
            if normalized and normalized not in normalized_to_original:
                normalized_to_original[normalized] = original

        for key, possible_names in HEADER_MAPPING.items():
            candidates = possible_names + [key]
            for candidate in candidates:
                normalized_candidate = _normalize_header(candidate)
                if normalized_candidate in normalized_to_original:
                    self.mapped_headers[key] = normalized_to_original[normalized_candidate]
                    break

        required_columns = ["unit_id", "bedrooms", "net_sf"]
        for col in required_columns:
            if col not in self.mapped_headers:
                example = HEADER_MAPPING[col][0]
                raise ValueError(
                    f"Error: Missing required column. Could not find a match for '{col}' (e.g., {example})."
                )

        return self.mapped_headers

    def _get_standardized_dataframe(self):
        """Internal method to get a DataFrame with standardized column names."""
        if not self.mapped_headers:
            self.map_headers()

        standardized_df = pd.DataFrame()
        for standard_name, original_name in self.mapped_headers.items():
            standardized_df[standard_name] = self.data[original_name]
        return standardized_df

    def get_affordable_units(self):
        """
        Reads the file, maps headers, identifies the affordable unit set,
        validates their data, and returns a clean DataFrame of units.
        """
        if self.data is None:
            self.read_data()
        if not self.mapped_headers:
            self.map_headers()

        df = self._get_standardized_dataframe()

        if "client_ami" not in df.columns:
            raise ValueError(
                "Error: The 'Client AMI' column (e.g., 'AMI', 'AFF %') is required to identify affordable units, but it was not found."
            )

        ami_series = df["client_ami"].astype(str).str.strip()
        is_percent = ami_series.str.contains("%", na=False)

        numeric_vals = pd.to_numeric(ami_series.str.replace("%", "", regex=False), errors="coerce").astype(float)
        numeric_vals = numeric_vals.where(~is_percent, numeric_vals / 100.0)

        df["client_ami"] = numeric_vals

        affordable_df = df[df["client_ami"] > 0].copy()

        if affordable_df.empty:
            raise ValueError(
                "Error: No affordable units found. Please ensure the 'Client AMI' column has positive numerical values for the units to be included."
            )

        affordable_df["unit_id"] = affordable_df["unit_id"].astype(str).str.strip()
        if affordable_df["unit_id"].isnull().any() or (affordable_df["unit_id"] == "").any():
            invalid_rows = affordable_df[affordable_df["unit_id"].isnull() | (affordable_df["unit_id"] == "")]
            raise ValueError(
                f"Error: Found affordable units with a missing Unit ID in rows: {invalid_rows.index.tolist()}"
            )

        for col in ["bedrooms", "net_sf"]:
            affordable_df[col] = pd.to_numeric(affordable_df[col], errors="coerce")
            if affordable_df[col].isnull().any():
                invalid_rows = affordable_df[affordable_df[col].isnull()]
                unit_ids = invalid_rows["unit_id"].tolist()
                raise ValueError(
                    f"Error: Invalid non-numeric data found in required column '{col}' for units: {unit_ids}"
                )

            if (affordable_df[col] < 0).any():
                invalid_rows = affordable_df[affordable_df[col] < 0]
                unit_ids = invalid_rows["unit_id"].tolist()
                raise ValueError(
                    f"Error: Column '{col}' must contain non-negative values. Found negative values for units: {unit_ids}"
                )

        return affordable_df
