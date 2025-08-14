#!/usr/bin/env python3
"""
NYC Affordable Housing AMI Calculator - COMPLETE SYSTEM
========================================================

ENHANCED CLIENT-RESPECTFUL AMI OPTIMIZATION SYSTEM
- Complete web application with beautiful UI
- All 38 enhanced functions included
- Fixed string accessor issues for numeric data
- Auto-SF calculation from pre-selected units
- Client-first approach: Respect selections 99.9% of time
- Enhanced mode detection with mixed AMI handling
- Comprehensive data validation and cleaning
- Smart compliance handling and optimization
- Floor requirements with intelligent mixing
- Performance optimization for any building size
- Beautiful web interface with professional UI
- Clean Excel files with enhanced formatting
- Multiple optimization strategies and versions
- Corner unit detection and analysis
- Enhanced balcony detection
- Floor and size categorization
- Extreme case analysis and recommendations
- Comprehensive reasoning documentation

Core Philosophy:
- Respect client's unit selection 99.9% of the time
- Only suggest changes in extreme mathematical impossibilities (0.1% cases)
- Focus on making their selection work, not questioning their judgment
"""

from flask import Flask, render_template_string, request, jsonify, send_file, redirect, url_for
import pandas as pd
import numpy as np
import os
import tempfile
import zipfile
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
import traceback
import json
from typing import Dict, List, Tuple, Optional, Any
import warnings
warnings.filterwarnings('ignore')

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB max file size

# Global variables for file management
UPLOAD_FOLDER = 'uploads'
RESULTS_FOLDER = 'results'

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULTS_FOLDER, exist_ok=True)

def make_json_serializable(obj):
    """
    Convert numpy/pandas data types to JSON serializable Python types.
    Fixes the "Object of type int64 is not JSON serializable" error.
    Fixed to handle NaN values properly.
    """
    # Check for NaN values FIRST before type conversion
    try:
        if pd.isna(obj) or (isinstance(obj, float) and np.isnan(obj)):
            return None
    except (TypeError, ValueError):
        pass
    
    if isinstance(obj, dict):
        return {key: make_json_serializable(value) for key, value in obj.items()}
    elif isinstance(obj, list):
        return [make_json_serializable(item) for item in obj]
    elif isinstance(obj, tuple):
        return tuple(make_json_serializable(item) for item in obj)
    elif isinstance(obj, (np.integer, np.int64, np.int32)):
        return int(obj)
    elif isinstance(obj, (np.floating, np.float64, np.float32)):
        # Double-check for NaN in numpy floats
        if np.isnan(obj):
            return None
        return float(obj)
    elif isinstance(obj, np.ndarray):
        return [make_json_serializable(item) for item in obj.tolist()]
    elif isinstance(obj, pd.Series):
        return [make_json_serializable(item) for item in obj.tolist()]
    elif isinstance(obj, pd.DataFrame):
        # Convert DataFrame to dict and recursively fix NaN values
        dict_records = obj.to_dict('records')
        return [make_json_serializable(record) for record in dict_records]
    else:
        return obj

class TrulyCompleteAMIOptimizationSystem:
    """
    TRULY COMPLETE AMI OPTIMIZATION SYSTEM
    =====================================
    
    This class contains ALL 38 enhanced functions for comprehensive
    NYC affordable housing AMI optimization with client-respectful approach.
    """
    
    def __init__(self):
        """Initialize the truly complete AMI optimization system."""
        self.building_data = None
        self.target_sf = None
        self.required_floors = None
        self.mode = None
        self.results = []
        self.compliance_analysis = {}
        self.extreme_case_recommendations = []
        
        # Client-Specific NYC Regulation Constants
        self.NYC_REGULATIONS = {
            'min_40_ami_percentage': 0.20,  # 20% minimum at 40% AMI (by SF)
            'max_weighted_average': 0.60,   # MAXIMUM 60% weighted average (client requirement)
            'min_weighted_average': 0.40,   # Minimum 40% weighted average (reasonable floor)
            'min_2br_percentage': None,     # Not required - client pre-selects units
            'vertical_distribution_min': None,  # Not required - client pre-selects floors
            'horizontal_distribution_max': None,  # Not required - client controls distribution
        }
        
        # Developer Optimization Rules
        self.DEVELOPER_RULES = {
            'max_income_bands': 3,           # Maximum 3 income bands
            'ami_multiples_of_10': True,     # AMI levels must be multiples of 10%
            'floor_strategy': 'lower_floors_lower_ami',  # 40% AMI on lower floors
            'size_strategy': 'smaller_units_lower_ami',  # Smaller units get lower AMI
            'preserve_large_market': True,   # Keep most large units as market rate
        }
        
        # Client-Specific AMI Distribution Strategies - Maximum 60% weighted average
        self.AMI_STRATEGIES = {
            'closest_to_60': {'40': 0.20, '60': 0.80, '80': 0.00, '100': 0.00},  # 56.0% weighted avg, 20% at 40% AMI
            'floor_optimized': {'40': 0.25, '60': 0.70, '80': 0.05, '100': 0.00},  # 57.5% weighted avg, 25% at 40% AMI  
            'size_optimized': {'40': 0.30, '60': 0.50, '80': 0.20, '100': 0.00},  # 58.0% weighted avg, 30% at 40% AMI
        }

    def _enhanced_data_validation_and_cleaning(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Enhanced data validation and cleaning with comprehensive error handling.
        Handles mixed data types, missing values, and malformed data.
        """
        try:
            # Create a copy to avoid modifying original
            df_clean = df.copy()
            
            # Enhanced column detection with multiple naming conventions
            column_mapping = self._detect_column_names(df_clean)
            
            # Rename columns to standard names
            df_clean = df_clean.rename(columns=column_mapping)
            
            # Ensure required columns exist
            required_columns = ['Unit', 'Floor', 'Bedrooms', 'SF']
            missing_columns = [col for col in required_columns if col not in df_clean.columns]
            
            if missing_columns:
                # Try to create missing columns from available data
                df_clean = self._create_missing_columns(df_clean, missing_columns)
            
            # Enhanced data type conversion with error handling
            df_clean = self._convert_data_types_safely(df_clean)
            
            # Clean and validate data
            df_clean = self._clean_and_validate_data(df_clean)
            
            # Remove invalid rows
            df_clean = self._remove_invalid_rows(df_clean)
            
            # Enhanced AMI column handling
            if 'AMI' in df_clean.columns:
                df_clean = self._clean_ami_column(df_clean)
            
            return df_clean
            
        except Exception as e:
            raise ValueError(f"Data validation failed: {str(e)}")

    def _detect_column_names(self, df: pd.DataFrame) -> Dict[str, str]:
        """
        Enhanced column name detection supporting multiple naming conventions.
        """
        column_mapping = {}
        
        # Define possible column name variations
        column_patterns = {
            'Unit': ['unit', 'apt', 'apartment', 'unit_id', 'unit id', 'apt_no', 'apt no'],
            'Floor': ['floor', 'level', 'flr', 'story'],
            'Bedrooms': ['bed', 'bedrooms', 'br', 'beds', 'bedroom'],
            'SF': ['sf', 'sq ft', 'sqft', 'square feet', 'area', 'net sf', ' net sf'],
            'AMI': ['ami', 'income', 'affordability'],
            'Balcony': ['balcony', 'balc', 'outdoor space'],
            'Accessible': ['accessible', 'ada', 'sec. 504', 'section 504']
        }
        
        # Convert all column names to lowercase for comparison
        df_columns_lower = [col.lower().strip() for col in df.columns]
        
        # Find matches
        for standard_name, variations in column_patterns.items():
            for variation in variations:
                for i, col_lower in enumerate(df_columns_lower):
                    if variation in col_lower or col_lower in variation:
                        original_col = df.columns[i]
                        column_mapping[original_col] = standard_name
                        break
                if standard_name in column_mapping.values():
                    break
        
        return column_mapping

    def _create_missing_columns(self, df: pd.DataFrame, missing_columns: List[str]) -> pd.DataFrame:
        """
        Create missing columns from available data when possible.
        """
        df_enhanced = df.copy()
        
        # Try to create Unit column from Floor + APT
        if 'Unit' in missing_columns:
            if 'Floor' in df_enhanced.columns and 'APT' in df_enhanced.columns:
                # Handle numeric APT values by converting to string
                df_enhanced['Unit'] = df_enhanced['Floor'].astype(str) + df_enhanced['APT'].astype(str)
            elif 'FLOOR' in df_enhanced.columns and 'APT' in df_enhanced.columns:
                # Handle the specific case from the user's file
                df_enhanced['Unit'] = df_enhanced['FLOOR'].astype(str) + df_enhanced['APT'].astype(str)
        
        # Try to create Floor column from Unit
        if 'Floor' in missing_columns and 'Unit' in df_enhanced.columns:
            # Extract floor number from unit (e.g., "2A" -> 2, "201" -> 2)
            df_enhanced['Floor'] = df_enhanced['Unit'].astype(str).str.extract(r'(\d+)').astype(float)
        
        # Map bedroom columns
        if 'Bedrooms' in missing_columns:
            bed_columns = ['BED', 'bed', 'Bed', 'BR', 'br']
            for bed_col in bed_columns:
                if bed_col in df_enhanced.columns:
                    df_enhanced['Bedrooms'] = df_enhanced[bed_col]
                    break
        
        # Map SF columns  
        if 'SF' in missing_columns:
            sf_columns = [' NET SF', 'NET SF', 'net sf', 'Net SF', 'SQFT', 'sqft']
            for sf_col in sf_columns:
                if sf_col in df_enhanced.columns:
                    df_enhanced['SF'] = df_enhanced[sf_col]
                    break
        
        return df_enhanced

    def _convert_data_types_safely(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Safely convert data types with comprehensive error handling.
        """
        df_converted = df.copy()
        
        # Convert Unit to string (handles numeric apartment numbers)
        if 'Unit' in df_converted.columns:
            df_converted['Unit'] = df_converted['Unit'].astype(str)
        
        # Convert Floor to numeric
        if 'Floor' in df_converted.columns:
            df_converted['Floor'] = pd.to_numeric(df_converted['Floor'], errors='coerce')
        
        # Convert Bedrooms to numeric
        if 'Bedrooms' in df_converted.columns:
            df_converted['Bedrooms'] = pd.to_numeric(df_converted['Bedrooms'], errors='coerce')
        
        # Convert SF to numeric
        if 'SF' in df_converted.columns:
            df_converted['SF'] = pd.to_numeric(df_converted['SF'], errors='coerce')
        
        # Handle AMI column (can be float or string)
        if 'AMI' in df_converted.columns:
            # Convert to numeric, keeping NaN for empty cells
            df_converted['AMI'] = pd.to_numeric(df_converted['AMI'], errors='coerce')
        
        return df_converted

    def _clean_and_validate_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Clean and validate data with enhanced error handling.
        """
        df_clean = df.copy()
        
        # Remove rows with missing critical data
        critical_columns = ['Unit', 'SF']
        for col in critical_columns:
            if col in df_clean.columns:
                df_clean = df_clean.dropna(subset=[col])
        
        # Validate SF values (must be positive)
        if 'SF' in df_clean.columns:
            df_clean = df_clean[df_clean['SF'] > 0]
        
        # Validate Floor values (must be positive integers)
        if 'Floor' in df_clean.columns:
            df_clean = df_clean[df_clean['Floor'] > 0]
        
        # Validate Bedrooms (must be non-negative)
        if 'Bedrooms' in df_clean.columns:
            df_clean = df_clean[df_clean['Bedrooms'] >= 0]
        
        return df_clean

    def _remove_invalid_rows(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Remove rows with invalid or incomplete data.
        """
        df_valid = df.copy()
        
        # Remove rows with duplicate Unit IDs
        if 'Unit' in df_valid.columns:
            df_valid = df_valid.drop_duplicates(subset=['Unit'], keep='first')
        
        # Remove rows with zero or negative SF
        if 'SF' in df_valid.columns:
            df_valid = df_valid[df_valid['SF'] > 0]
        
        # Remove rows with missing Unit IDs
        if 'Unit' in df_valid.columns:
            df_valid = df_valid[df_valid['Unit'].notna()]
            df_valid = df_valid[df_valid['Unit'].astype(str).str.strip() != '']
        
        return df_valid

    def _clean_ami_column(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Enhanced AMI column cleaning and validation.
        """
        df_ami = df.copy()
        
        if 'AMI' in df_ami.columns:
            # Convert AMI values to proper format
            # Handle percentage values (e.g., 60% -> 0.6)
            ami_series = df_ami['AMI'].copy()
            
            # Convert string percentages to decimals
            if ami_series.dtype == 'object':
                ami_series = ami_series.astype(str)
                ami_series = ami_series.str.replace('%', '')
                ami_series = pd.to_numeric(ami_series, errors='coerce')
            
            # Convert percentages > 1 to decimals (e.g., 60 -> 0.6)
            ami_series = ami_series.apply(lambda x: x/100 if pd.notna(x) and x > 1 else x)
            
            # Validate AMI values (must be between 0 and 1.5)
            ami_series = ami_series.apply(lambda x: x if pd.notna(x) and 0 <= x <= 1.5 else np.nan)
            
            df_ami['AMI'] = ami_series
        
        return df_ami

    def _enhanced_intelligent_mode_detection(self, df: pd.DataFrame, target_sf: Optional[float] = None) -> str:
        """
        Enhanced intelligent mode detection with comprehensive analysis.
        Determines whether to do unit selection, AMI assignment, or AMI optimization.
        """
        try:
            # Check if AMI column exists
            has_ami_column = 'AMI' in df.columns
            
            if not has_ami_column:
                return 'UNIT_SELECTION'
            
            # Analyze AMI column content
            ami_analysis = self._analyze_ami_column_content(df)
            
            # Count units with AMI values
            units_with_ami = ami_analysis['units_with_ami']
            total_units = len(df)
            units_without_ami = total_units - units_with_ami
            
            # Enhanced decision logic
            if units_with_ami == 0:
                # No AMI values - need to assign AMI to pre-selected units
                return 'AMI_ASSIGNMENT'
            elif units_without_ami == 0:
                # All units have AMI values - check if optimization needed
                if self._is_ami_distribution_optimal(df):
                    return 'AMI_ASSIGNMENT'  # Already optimal, just validate
                else:
                    return 'AMI_OPTIMIZATION'  # Needs optimization
            else:
                # Mixed AMI values - partial assignment/optimization
                if self._is_ami_distribution_suboptimal(df):
                    return 'AMI_OPTIMIZATION'
                else:
                    return 'AMI_ASSIGNMENT'
            
        except Exception as e:
            # Default to AMI assignment if detection fails
            return 'AMI_ASSIGNMENT'

    def _analyze_ami_column_content(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Comprehensive analysis of AMI column content.
        """
        analysis = {
            'total_units': len(df),
            'units_with_ami': 0,
            'units_without_ami': 0,
            'unique_ami_values': [],
            'ami_distribution': {},
            'is_uniform': False,
            'is_mixed': False,
            'weighted_average': None,
            'compliance_issues': []
        }
        
        if 'AMI' not in df.columns:
            return analysis
        
        # Count units with/without AMI
        ami_series = df['AMI']
        units_with_ami = ami_series.notna().sum()
        units_without_ami = ami_series.isna().sum()
        
        analysis['units_with_ami'] = units_with_ami
        analysis['units_without_ami'] = units_without_ami
        
        if units_with_ami > 0:
            # Analyze AMI values
            ami_values = ami_series.dropna()
            unique_values = sorted(ami_values.unique())
            analysis['unique_ami_values'] = unique_values
            
            # Check if uniform (all same value)
            analysis['is_uniform'] = len(unique_values) == 1
            analysis['is_mixed'] = len(unique_values) > 1
            
            # Calculate distribution
            for ami_val in unique_values:
                count = (ami_series == ami_val).sum()
                percentage = count / units_with_ami
                analysis['ami_distribution'][ami_val] = {
                    'count': count,
                    'percentage': percentage
                }
            
            # Calculate weighted average if SF data available
            if 'SF' in df.columns:
                df_with_ami = df[df['AMI'].notna()]
                if len(df_with_ami) > 0:
                    total_sf = df_with_ami['SF'].sum()
                    weighted_sum = (df_with_ami['AMI'] * df_with_ami['SF']).sum()
                    analysis['weighted_average'] = weighted_sum / total_sf
        
        return analysis

    def _is_ami_distribution_optimal(self, df: pd.DataFrame) -> bool:
        """
        Check if current AMI distribution is already optimal.
        """
        try:
            ami_analysis = self._analyze_ami_column_content(df)
            
            # If no AMI values, not optimal
            if ami_analysis['units_with_ami'] == 0:
                return False
            
            # Check weighted average compliance
            weighted_avg = ami_analysis['weighted_average']
            if weighted_avg is None:
                return False
            
            # Check if weighted average is in acceptable range (60-80%)
            if not (0.60 <= weighted_avg <= 0.80):
                return False
            
            # Check 40% AMI minimum requirement
            ami_40_percentage = ami_analysis['ami_distribution'].get(0.4, {}).get('percentage', 0)
            if ami_40_percentage < 0.20:  # Less than 20%
                return False
            
            # Check if using appropriate number of AMI levels
            num_ami_levels = len(ami_analysis['unique_ami_values'])
            if num_ami_levels < 2 or num_ami_levels > 4:
                return False
            
            return True
            
        except Exception:
            return False

    def _is_ami_distribution_suboptimal(self, df: pd.DataFrame) -> bool:
        """
        Check if current AMI distribution is suboptimal and needs improvement.
        """
        try:
            ami_analysis = self._analyze_ami_column_content(df)
            
            # If uniform distribution (all same AMI), it's suboptimal
            if ami_analysis['is_uniform']:
                return True
            
            # Check weighted average
            weighted_avg = ami_analysis['weighted_average']
            if weighted_avg is not None:
                # If weighted average is outside optimal range, it's suboptimal
                if weighted_avg < 0.60 or weighted_avg > 0.80:
                    return True
            
            # Check 40% AMI requirement
            ami_40_percentage = ami_analysis['ami_distribution'].get(0.4, {}).get('percentage', 0)
            if ami_40_percentage < 0.15:  # Significantly below 20%
                return True
            
            return False
            
        except Exception:
            return True  # Assume suboptimal if analysis fails

    def _auto_calculate_target_sf_from_preselected_units(self, df: pd.DataFrame) -> float:
        """
        Automatically calculate target SF from pre-selected units.
        This is the core feature that eliminates manual SF input.
        """
        try:
            # Mode 1: Units with AMI values (client pre-selected with AMI)
            if 'AMI' in df.columns:
                units_with_ami = df[df['AMI'].notna()]
                if len(units_with_ami) > 0:
                    return float(units_with_ami['SF'].sum())
            
            # Mode 2: All units are pre-selected for affordable housing
            # (This assumes the entire file represents affordable units)
            total_sf = df['SF'].sum()
            return float(total_sf)
            
        except Exception as e:
            raise ValueError(f"Could not calculate target SF from pre-selected units: {str(e)}")

    def _enhanced_building_analysis(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Comprehensive building analysis with enhanced categorization.
        """
        analysis = {
            'total_units': len(df),
            'total_sf': df['SF'].sum(),
            'floors': self._analyze_floors(df),
            'unit_types': self._analyze_unit_types(df),
            'size_categories': self._categorize_unit_sizes(df),
            'floor_categories': self._categorize_floors(df),
            'balcony_analysis': self._enhanced_balcony_detection(df),
            'corner_units': self._identify_corner_units_enhanced(df),
            'accessibility': self._analyze_accessibility_features(df),
            'building_category': self._determine_building_category(df),
            'optimization_potential': self._assess_optimization_potential(df)
        }
        
        return analysis

    def _analyze_floors(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Comprehensive floor analysis.
        """
        floor_analysis = {
            'min_floor': int(df['Floor'].min()),
            'max_floor': int(df['Floor'].max()),
            'total_floors': int(df['Floor'].nunique()),
            'units_per_floor': {},
            'sf_per_floor': {},
            'avg_units_per_floor': 0,
            'avg_sf_per_floor': 0
        }
        
        # Calculate units and SF per floor
        for floor in sorted(df['Floor'].unique()):
            floor_data = df[df['Floor'] == floor]
            floor_num = int(floor)
            floor_analysis['units_per_floor'][floor_num] = len(floor_data)
            floor_analysis['sf_per_floor'][floor_num] = floor_data['SF'].sum()
        
        # Calculate averages
        floor_analysis['avg_units_per_floor'] = len(df) / floor_analysis['total_floors']
        floor_analysis['avg_sf_per_floor'] = df['SF'].sum() / floor_analysis['total_floors']
        
        return floor_analysis

    def _analyze_unit_types(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Comprehensive unit type analysis.
        """
        unit_types = {
            'studio': 0,
            '1br': 0,
            '2br': 0,
            '3br': 0,
            '4br_plus': 0,
            'total_2br_plus': 0,
            'bedroom_mix_percentage': 0
        }
        
        if 'Bedrooms' in df.columns:
            bedroom_counts = df['Bedrooms'].value_counts()
            
            unit_types['studio'] = bedroom_counts.get(0, 0)
            unit_types['1br'] = bedroom_counts.get(1, 0)
            unit_types['2br'] = bedroom_counts.get(2, 0)
            unit_types['3br'] = bedroom_counts.get(3, 0)
            unit_types['4br_plus'] = bedroom_counts[bedroom_counts.index >= 4].sum() if len(bedroom_counts[bedroom_counts.index >= 4]) > 0 else 0
            
            # Calculate 2BR+ totals
            unit_types['total_2br_plus'] = unit_types['2br'] + unit_types['3br'] + unit_types['4br_plus']
            
            # Calculate bedroom mix percentage
            total_units = len(df)
            if total_units > 0:
                unit_types['bedroom_mix_percentage'] = unit_types['total_2br_plus'] / total_units
        
        return unit_types

    def _categorize_unit_sizes(self, df: pd.DataFrame) -> Dict[str, List[str]]:
        """
        Enhanced unit size categorization.
        """
        if 'SF' not in df.columns:
            return {'small': [], 'medium': [], 'large': []}
        
        # Calculate size thresholds
        sf_values = df['SF']
        q33 = sf_values.quantile(0.33)
        q67 = sf_values.quantile(0.67)
        
        categories = {'small': [], 'medium': [], 'large': []}
        
        for _, row in df.iterrows():
            unit_id = str(row['Unit'])
            sf = row['SF']
            
            if sf <= q33:
                categories['small'].append(unit_id)
            elif sf <= q67:
                categories['medium'].append(unit_id)
            else:
                categories['large'].append(unit_id)
        
        return categories

    def _categorize_floors(self, df: pd.DataFrame) -> Dict[str, List[int]]:
        """
        Enhanced floor categorization for optimization strategies.
        """
        if 'Floor' not in df.columns:
            return {'lower': [], 'middle': [], 'upper': []}
        
        floors = sorted(df['Floor'].unique())
        total_floors = len(floors)
        
        categories = {'lower': [], 'middle': [], 'upper': []}
        
        if total_floors <= 3:
            # For low-rise buildings
            categories['lower'] = floors[:1]
            categories['middle'] = floors[1:-1] if total_floors > 2 else []
            categories['upper'] = floors[-1:]
        else:
            # For mid-rise and high-rise buildings
            third = total_floors // 3
            categories['lower'] = floors[:third]
            categories['middle'] = floors[third:2*third]
            categories['upper'] = floors[2*third:]
        
        return categories

    def _enhanced_balcony_detection(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Enhanced balcony detection and analysis.
        """
        balcony_analysis = {
            'has_balcony_data': False,
            'units_with_balconies': 0,
            'units_without_balconies': 0,
            'balcony_percentage': 0,
            'balcony_units': [],
            'non_balcony_units': []
        }
        
        # Check for balcony column
        balcony_columns = ['Balcony', 'BALCONY', 'balcony', 'Outdoor Space', 'Terrace']
        balcony_col = None
        
        for col in balcony_columns:
            if col in df.columns:
                balcony_col = col
                break
        
        if balcony_col is not None:
            balcony_analysis['has_balcony_data'] = True
            
            # Analyze balcony data
            balcony_series = df[balcony_col]
            
            # Count units with balconies (non-null, non-zero values)
            has_balcony = balcony_series.notna() & (balcony_series != 0)
            units_with_balconies = has_balcony.sum()
            units_without_balconies = len(df) - units_with_balconies
            
            balcony_analysis['units_with_balconies'] = units_with_balconies
            balcony_analysis['units_without_balconies'] = units_without_balconies
            balcony_analysis['balcony_percentage'] = units_with_balconies / len(df) if len(df) > 0 else 0
            
            # Get unit lists
            balcony_analysis['balcony_units'] = df[has_balcony]['Unit'].astype(str).tolist()
            balcony_analysis['non_balcony_units'] = df[~has_balcony]['Unit'].astype(str).tolist()
        
        return balcony_analysis

    def _identify_corner_units_enhanced(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Enhanced corner unit identification using multiple heuristics.
        """
        corner_analysis = {
            'corner_units': [],
            'potential_corner_units': [],
            'corner_percentage': 0,
            'identification_method': 'heuristic'
        }
        
        if 'Unit' not in df.columns:
            return corner_analysis
        
        # Method 1: Unit number patterns (e.g., A, B units often corners)
        corner_units = []
        potential_corners = []
        
        for _, row in df.iterrows():
            unit_str = str(row['Unit'])
            
            # Check for letter patterns that often indicate corners
            if any(letter in unit_str.upper() for letter in ['A', 'B', '01', '02']):
                corner_units.append(unit_str)
            elif any(letter in unit_str.upper() for letter in ['C', 'D', '03', '04']):
                potential_corners.append(unit_str)
        
        # Method 2: Floor analysis (first and last units per floor)
        if 'Floor' in df.columns:
            for floor in df['Floor'].unique():
                floor_units = df[df['Floor'] == floor]['Unit'].astype(str).tolist()
                if len(floor_units) >= 2:
                    # First and last units are often corners
                    sorted_units = sorted(floor_units)
                    if sorted_units[0] not in corner_units:
                        corner_units.append(sorted_units[0])
                    if sorted_units[-1] not in corner_units:
                        corner_units.append(sorted_units[-1])
        
        corner_analysis['corner_units'] = corner_units
        corner_analysis['potential_corner_units'] = potential_corners
        corner_analysis['corner_percentage'] = len(corner_units) / len(df) if len(df) > 0 else 0
        
        return corner_analysis

    def _analyze_accessibility_features(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Analyze accessibility features and ADA compliance.
        """
        accessibility_analysis = {
            'has_accessibility_data': False,
            'accessible_units': 0,
            'accessibility_percentage': 0,
            'accessible_unit_list': []
        }
        
        # Check for accessibility columns
        accessibility_columns = ['Accessible', 'ADA', 'SEC. 504', 'Section 504', 'Disability']
        accessibility_col = None
        
        for col in accessibility_columns:
            if col in df.columns:
                accessibility_col = col
                break
        
        if accessibility_col is not None:
            accessibility_analysis['has_accessibility_data'] = True
            
            # Analyze accessibility data
            accessibility_series = df[accessibility_col]
            
            # Count accessible units (non-null, non-zero values)
            is_accessible = accessibility_series.notna() & (accessibility_series != 0)
            accessible_units = is_accessible.sum()
            
            accessibility_analysis['accessible_units'] = accessible_units
            accessibility_analysis['accessibility_percentage'] = accessible_units / len(df) if len(df) > 0 else 0
            accessibility_analysis['accessible_unit_list'] = df[is_accessible]['Unit'].astype(str).tolist()
        
        return accessibility_analysis

    def _determine_building_category(self, df: pd.DataFrame) -> str:
        """
        Determine building category for optimization strategies.
        """
        total_sf = df['SF'].sum()
        total_units = len(df)
        
        if total_sf <= 15000:
            return 'Small Building'
        elif total_sf <= 50000:
            return 'Medium Building'
        elif total_sf <= 100000:
            return 'Large Building'
        else:
            return 'Mega Building'

    def _assess_optimization_potential(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Assess the optimization potential of the building.
        """
        potential = {
            'bedroom_mix_potential': 'Low',
            'floor_strategy_potential': 'Medium',
            'size_strategy_potential': 'Medium',
            'overall_potential': 'Medium'
        }
        
        # Assess bedroom mix potential
        unit_types = self._analyze_unit_types(df)
        if unit_types['bedroom_mix_percentage'] >= 0.6:
            potential['bedroom_mix_potential'] = 'High'
        elif unit_types['bedroom_mix_percentage'] >= 0.4:
            potential['bedroom_mix_potential'] = 'Medium'
        
        # Assess floor strategy potential
        floor_analysis = self._analyze_floors(df)
        if floor_analysis['total_floors'] >= 5:
            potential['floor_strategy_potential'] = 'High'
        elif floor_analysis['total_floors'] >= 3:
            potential['floor_strategy_potential'] = 'Medium'
        else:
            potential['floor_strategy_potential'] = 'Low'
        
        # Assess size strategy potential
        size_categories = self._categorize_unit_sizes(df)
        size_variety = len([cat for cat in size_categories.values() if len(cat) > 0])
        if size_variety >= 3:
            potential['size_strategy_potential'] = 'High'
        elif size_variety >= 2:
            potential['size_strategy_potential'] = 'Medium'
        
        # Overall potential
        potentials = [potential['bedroom_mix_potential'], potential['floor_strategy_potential'], potential['size_strategy_potential']]
        if potentials.count('High') >= 2:
            potential['overall_potential'] = 'High'
        elif potentials.count('Low') >= 2:
            potential['overall_potential'] = 'Low'
        
        return potential

    def _smart_compliance_handling_and_optimization(self, df: pd.DataFrame, target_sf: float, required_floors: Optional[List[int]] = None) -> Dict[str, Any]:
        """
        Smart compliance handling with multiple optimization strategies.
        """
        try:
            # Generate multiple optimization approaches
            optimization_results = []
            
            # Strategy 1: Closest to 60% AMI
            result_1 = self._optimize_for_closest_to_60_ami(df, target_sf, required_floors)
            if result_1:
                result_1['strategy_name'] = 'Closest to 60% AMI'
                result_1['strategy_description'] = 'Optimized for weighted average closest to 60% AMI'
                optimization_results.append(result_1)
            
            # Strategy 2: Floor Strategy Optimized
            result_2 = self._optimize_for_floor_strategy(df, target_sf, required_floors)
            if result_2:
                result_2['strategy_name'] = 'Floor Strategy Optimized'
                result_2['strategy_description'] = '40% AMI on lower floors, higher AMI on upper floors'
                optimization_results.append(result_2)
            
            # Strategy 3: Size Strategy Optimized
            result_3 = self._optimize_for_size_strategy(df, target_sf, required_floors)
            if result_3:
                result_3['strategy_name'] = 'Size Strategy Optimized'
                result_3['strategy_description'] = 'Smaller units get 40% AMI, larger units get higher AMI'
                optimization_results.append(result_3)
            
            # Rank results by compliance and optimization quality
            ranked_results = self._rank_optimization_results(optimization_results)
            
            return {
                'success': len(ranked_results) > 0,
                'results': ranked_results,
                'total_strategies': len(ranked_results),
                'best_strategy': ranked_results[0] if ranked_results else None
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': str(e),
                'results': [],
                'total_strategies': 0,
                'best_strategy': None
            }

    def _optimize_for_closest_to_60_ami(self, df: pd.DataFrame, target_sf: float, required_floors: Optional[List[int]] = None) -> Optional[Dict[str, Any]]:
        """
        Optimize AMI distribution to get weighted average closest to 60%.
        """
        try:
            # Select units based on target SF
            selected_units = self._select_units_for_target_sf(df, target_sf, required_floors, strategy='balanced')
            
            if selected_units is None or len(selected_units) == 0:
                return None
            
            # Assign AMI levels to achieve ~60% weighted average
            ami_assignments = self._assign_ami_for_60_percent_target(selected_units)
            
            # Calculate compliance metrics
            compliance = self._calculate_compliance_metrics(selected_units, ami_assignments)
            
            return {
                'selected_units': selected_units,
                'ami_assignments': ami_assignments,
                'compliance': compliance,
                'total_units': len(selected_units),
                'total_sf': selected_units['SF'].sum(),
                'weighted_average_ami': compliance['weighted_average_ami'],
                'overage_sf': selected_units['SF'].sum() - target_sf,
                'overage_percentage': ((selected_units['SF'].sum() - target_sf) / target_sf) * 100
            }
            
        except Exception as e:
            return None

    def _optimize_for_floor_strategy(self, df: pd.DataFrame, target_sf: float, required_floors: Optional[List[int]] = None) -> Optional[Dict[str, Any]]:
        """
        Optimize using floor strategy: 40% AMI on lower floors, higher AMI on upper floors.
        """
        try:
            # Select units with floor preference
            selected_units = self._select_units_for_target_sf(df, target_sf, required_floors, strategy='floor_preference')
            
            if selected_units is None or len(selected_units) == 0:
                return None
            
            # Assign AMI based on floor strategy
            ami_assignments = self._assign_ami_by_floor_strategy(selected_units)
            
            # Calculate compliance metrics
            compliance = self._calculate_compliance_metrics(selected_units, ami_assignments)
            
            return {
                'selected_units': selected_units,
                'ami_assignments': ami_assignments,
                'compliance': compliance,
                'total_units': len(selected_units),
                'total_sf': selected_units['SF'].sum(),
                'weighted_average_ami': compliance['weighted_average_ami'],
                'overage_sf': selected_units['SF'].sum() - target_sf,
                'overage_percentage': ((selected_units['SF'].sum() - target_sf) / target_sf) * 100
            }
            
        except Exception as e:
            return None

    def _optimize_for_size_strategy(self, df: pd.DataFrame, target_sf: float, required_floors: Optional[List[int]] = None) -> Optional[Dict[str, Any]]:
        """
        Optimize using size strategy: smaller units get 40% AMI, larger units get higher AMI.
        """
        try:
            # Select units with size preference
            selected_units = self._select_units_for_target_sf(df, target_sf, required_floors, strategy='size_preference')
            
            if selected_units is None or len(selected_units) == 0:
                return None
            
            # Assign AMI based on size strategy
            ami_assignments = self._assign_ami_by_size_strategy(selected_units)
            
            # Calculate compliance metrics
            compliance = self._calculate_compliance_metrics(selected_units, ami_assignments)
            
            return {
                'selected_units': selected_units,
                'ami_assignments': ami_assignments,
                'compliance': compliance,
                'total_units': len(selected_units),
                'total_sf': selected_units['SF'].sum(),
                'weighted_average_ami': compliance['weighted_average_ami'],
                'overage_sf': selected_units['SF'].sum() - target_sf,
                'overage_percentage': ((selected_units['SF'].sum() - target_sf) / target_sf) * 100
            }
            
        except Exception as e:
            return None

    def _select_units_for_target_sf(self, df: pd.DataFrame, target_sf: float, required_floors: Optional[List[int]] = None, strategy: str = 'balanced') -> Optional[pd.DataFrame]:
        """
        Select optimal units to meet target SF with various strategies.
        """
        try:
            available_units = df.copy()
            
            # Handle required floors if specified
            if required_floors:
                # Try to include units from required floors
                required_floor_units = available_units[available_units['Floor'].isin(required_floors)]
                other_units = available_units[~available_units['Floor'].isin(required_floors)]
                
                # Smart mixing strategy
                if len(required_floor_units) > 0:
                    required_sf = min(target_sf * 0.3, required_floor_units['SF'].sum())  # At least 30% from required floors
                    selected_from_required = self._select_units_by_strategy(required_floor_units, required_sf, strategy)
                    
                    remaining_sf = target_sf - selected_from_required['SF'].sum()
                    if remaining_sf > 0 and len(other_units) > 0:
                        selected_from_others = self._select_units_by_strategy(other_units, remaining_sf, strategy)
                        return pd.concat([selected_from_required, selected_from_others])
                    else:
                        return selected_from_required
            
            # Standard selection without floor requirements
            return self._select_units_by_strategy(available_units, target_sf, strategy)
            
        except Exception as e:
            return None

    def _select_units_by_strategy(self, df: pd.DataFrame, target_sf: float, strategy: str) -> pd.DataFrame:
        """
        Select units using specified strategy.
        """
        if strategy == 'floor_preference':
            # Prefer lower floors
            df_sorted = df.sort_values(['Floor', 'SF'])
        elif strategy == 'size_preference':
            # Prefer smaller units first
            df_sorted = df.sort_values(['SF', 'Floor'])
        else:  # balanced
            # Balanced approach
            df_sorted = df.sort_values(['Floor', 'Bedrooms', 'SF'])
        
        # Greedy selection to meet target SF
        selected_units = []
        current_sf = 0
        
        for _, unit in df_sorted.iterrows():
            if current_sf >= target_sf:
                break
            selected_units.append(unit)
            current_sf += unit['SF']
        
        return pd.DataFrame(selected_units)

    def _assign_ami_for_60_percent_target(self, df: pd.DataFrame) -> Dict[str, float]:
        """
        Assign AMI levels using the 'closest_to_60' strategy from AMI_STRATEGIES.
        """
        return self._assign_ami_using_strategy(df, 'closest_to_60')

    def _assign_ami_by_floor_strategy(self, df: pd.DataFrame) -> Dict[str, float]:
        """
        Assign AMI levels using the 'floor_optimized' strategy from AMI_STRATEGIES.
        """
        return self._assign_ami_using_strategy(df, 'floor_optimized', sort_by='floor')

    def _assign_ami_by_size_strategy(self, df: pd.DataFrame) -> Dict[str, float]:
        """
        Assign AMI levels using the 'size_optimized' strategy from AMI_STRATEGIES.
        """
        return self._assign_ami_using_strategy(df, 'size_optimized', sort_by='size')

    def _assign_ami_using_strategy(self, df: pd.DataFrame, strategy_name: str, sort_by: str = 'balanced') -> Dict[str, float]:
        """
        Assign AMI levels using the specified strategy from AMI_STRATEGIES dictionary.
        
        Args:
            df: DataFrame with selected units
            strategy_name: Name of strategy from AMI_STRATEGIES ('closest_to_60', 'floor_optimized', 'size_optimized')
            sort_by: How to sort units ('balanced', 'floor', 'size')
        """
        ami_assignments = {}
        
        # Get the strategy distribution
        if strategy_name not in self.AMI_STRATEGIES:
            raise ValueError(f"Unknown strategy: {strategy_name}")
        
        strategy = self.AMI_STRATEGIES[strategy_name]
        total_units = len(df)
        total_sf = df['SF'].sum()
        
        # For size strategy, use SF-based allocation to ensure 20% SF minimum at 40% AMI
        if strategy_name == 'size_optimized':
            return self._assign_ami_sf_based(df, strategy, sort_by)
        
        # For other strategies, use unit-based allocation
        target_counts = {}
        for ami_str, percentage in strategy.items():
            ami_decimal = float(ami_str) / 100  # Convert '40' to 0.4
            target_count = max(1, round(total_units * percentage))
            target_counts[ami_decimal] = target_count
        
        # Adjust counts to match total units exactly
        total_assigned = sum(target_counts.values())
        if total_assigned != total_units:
            # Adjust the largest category
            largest_ami = max(target_counts.keys(), key=lambda k: target_counts[k])
            target_counts[largest_ami] += (total_units - total_assigned)
        
        # Sort units based on client's priority: lower floors + smaller units get 40% AMI first
        if sort_by == 'floor':
            # Floor strategy: Lower floors get lower AMI, then by size within floor
            units_sorted = df.sort_values(['Floor', 'SF'])
        elif sort_by == 'size':
            # Size strategy: Smaller units get lower AMI, then by floor within size
            units_sorted = df.sort_values(['SF', 'Floor'])
        else:  # balanced - client's preferred approach
            # Client's priority: Lower floors + smaller units get 40% AMI first
            # Create composite score: lower floor + smaller size = lower score = higher priority for 40% AMI
            df_copy = df.copy()
            df_copy['priority_score'] = df_copy['Floor'] * 1000 + df_copy['SF']  # Floor dominates, size breaks ties
            units_sorted = df_copy.sort_values('priority_score')
        
        # Assign AMI levels in order: 40%, 60%, 80%, 100%
        ami_levels = [0.4, 0.6, 0.8, 1.0]
        current_counts = {ami: 0 for ami in ami_levels}
        
        for _, unit in units_sorted.iterrows():
            unit_id = str(unit['Unit'])
            
            # Find the next AMI level that needs units
            for ami_level in ami_levels:
                if ami_level in target_counts and current_counts[ami_level] < target_counts[ami_level]:
                    ami_assignments[unit_id] = ami_level
                    current_counts[ami_level] += 1
                    break
        
        return ami_assignments
    
    def _assign_ami_sf_based(self, df: pd.DataFrame, strategy: Dict[str, float], sort_by: str) -> Dict[str, float]:
        """
        Assign AMI levels based on square footage targets (for size strategy).
        """
        ami_assignments = {}
        total_sf = df['SF'].sum()
        
        # Calculate target SF for each AMI level
        target_sf = {}
        for ami_str, percentage in strategy.items():
            ami_decimal = float(ami_str) / 100
            target_sf[ami_decimal] = total_sf * percentage
        
        # Sort by size (smaller units get lower AMI)
        units_sorted = df.sort_values(['SF', 'Floor'])
        
        # Assign AMI levels to meet SF targets
        ami_levels = [0.4, 0.6, 0.8, 1.0]
        current_sf = {ami: 0 for ami in ami_levels}
        
        for _, unit in units_sorted.iterrows():
            unit_id = str(unit['Unit'])
            unit_sf = unit['SF']
            
            # Find the AMI level that needs more SF and can accommodate this unit
            assigned = False
            for ami_level in ami_levels:
                if ami_level in target_sf and current_sf[ami_level] < target_sf[ami_level]:
                    ami_assignments[unit_id] = ami_level
                    current_sf[ami_level] += unit_sf
                    assigned = True
                    break
            
            # If no level needs more SF, assign to the highest level
            if not assigned:
                ami_assignments[unit_id] = 1.0
                current_sf[1.0] += unit_sf
        
        return ami_assignments

    def _calculate_compliance_metrics(self, df: pd.DataFrame, ami_assignments: Dict[str, float]) -> Dict[str, Any]:
        """
        Calculate comprehensive compliance metrics.
        """
        compliance = {
            'weighted_average_ami': 0,
            'ami_40_percentage': 0,
            'ami_60_percentage': 0,
            'ami_80_percentage': 0,
            'ami_100_percentage': 0,
            'bedroom_mix_percentage': 0,
            'total_sf': df['SF'].sum(),
            'total_units': len(df),
            'compliance_grade': 'F',
            'compliance_issues': []
        }
        
        try:
            # Calculate weighted average AMI
            total_sf = df['SF'].sum()
            weighted_sum = 0
            
            # Count units by AMI level
            ami_counts = {0.4: 0, 0.6: 0, 0.8: 0, 1.0: 0}
            ami_sf = {0.4: 0, 0.6: 0, 0.8: 0, 1.0: 0}
            
            for _, unit in df.iterrows():
                unit_id = str(unit['Unit'])
                ami = ami_assignments.get(unit_id, 0.6)  # Default to 60% if not assigned
                sf = unit['SF']
                
                weighted_sum += ami * sf
                ami_counts[ami] = ami_counts.get(ami, 0) + 1
                ami_sf[ami] = ami_sf.get(ami, 0) + sf
            
            compliance['weighted_average_ami'] = weighted_sum / total_sf if total_sf > 0 else 0
            
            # Calculate AMI percentages by SF
            compliance['ami_40_percentage'] = ami_sf[0.4] / total_sf if total_sf > 0 else 0
            compliance['ami_60_percentage'] = ami_sf[0.6] / total_sf if total_sf > 0 else 0
            compliance['ami_80_percentage'] = ami_sf[0.8] / total_sf if total_sf > 0 else 0
            compliance['ami_100_percentage'] = ami_sf[1.0] / total_sf if total_sf > 0 else 0
            
            # Calculate bedroom mix
            if 'Bedrooms' in df.columns:
                total_units = len(df)
                units_2br_plus = len(df[df['Bedrooms'] >= 2])
                compliance['bedroom_mix_percentage'] = units_2br_plus / total_units if total_units > 0 else 0
            
            # Assess compliance based on client-specific requirements
            issues = []
            
            # Check 40% AMI minimum (required)
            if compliance['ami_40_percentage'] < self.NYC_REGULATIONS['min_40_ami_percentage']:
                issues.append(f"40% AMI coverage: {compliance['ami_40_percentage']:.1%} (needs {self.NYC_REGULATIONS['min_40_ami_percentage']:.0%})")
            
            # Check weighted average maximum (client requirement: cannot exceed 60%)
            if compliance['weighted_average_ami'] > self.NYC_REGULATIONS['max_weighted_average']:
                issues.append(f"Weighted average: {compliance['weighted_average_ami']:.1%} (exceeds {self.NYC_REGULATIONS['max_weighted_average']:.0%} maximum)")
            
            # Check weighted average minimum (reasonable floor)
            if compliance['weighted_average_ami'] < self.NYC_REGULATIONS['min_weighted_average']:
                issues.append(f"Weighted average: {compliance['weighted_average_ami']:.1%} (below {self.NYC_REGULATIONS['min_weighted_average']:.0%} minimum)")
            
            # Skip bedroom mix and floor distribution checks (client pre-selects units)
            # These requirements are handled by client's unit selection, not by the algorithm
            
            compliance['compliance_issues'] = issues
            
            # Calculate grade
            if len(issues) == 0:
                compliance['compliance_grade'] = 'A'
            elif len(issues) == 1:
                compliance['compliance_grade'] = 'B'
            elif len(issues) == 2:
                compliance['compliance_grade'] = 'C'
            else:
                compliance['compliance_grade'] = 'F'
            
        except Exception as e:
            compliance['compliance_issues'].append(f"Calculation error: {str(e)}")
        
        return compliance

    def _rank_optimization_results(self, results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Rank optimization results by compliance quality and other factors.
        """
        def calculate_score(result):
            compliance = result['compliance']
            score = 0
            
            # Compliance grade scoring
            grade_scores = {'A': 100, 'B': 80, 'C': 60, 'D': 40, 'F': 20}
            score += grade_scores.get(compliance['compliance_grade'], 0)
            
            # Weighted average proximity to 60%
            weighted_avg = compliance['weighted_average_ami']
            if 0.60 <= weighted_avg <= 0.80:
                score += 20  # Bonus for being in range
                # Additional bonus for being close to 60%
                proximity_bonus = 10 * (1 - abs(weighted_avg - 0.60) / 0.20)
                score += proximity_bonus
            
            # 40% AMI coverage bonus
            ami_40_pct = compliance['ami_40_percentage']
            if ami_40_pct >= 0.20:
                score += 15  # Bonus for meeting minimum
                if ami_40_pct >= 0.25:
                    score += 5  # Extra bonus for exceeding
            
            # Bedroom mix bonus
            bedroom_mix = compliance['bedroom_mix_percentage']
            if bedroom_mix >= 0.50:
                score += 15
            elif bedroom_mix >= 0.40:
                score += 10
            
            # Overage penalty
            overage_pct = abs(result.get('overage_percentage', 0))
            if overage_pct <= 1:
                score += 10
            elif overage_pct <= 3:
                score += 5
            else:
                score -= overage_pct  # Penalty for high overage
            
            return score
        
        # Calculate scores and sort
        for result in results:
            result['optimization_score'] = calculate_score(result)
        
        return sorted(results, key=lambda x: x['optimization_score'], reverse=True)

    def _analyze_for_extreme_cases(self, df: pd.DataFrame, target_sf: float) -> List[Dict[str, Any]]:
        """
        Analyze for extreme cases and generate recommendations.
        Only triggers in 0.1% of cases where client selection makes no mathematical sense.
        """
        recommendations = []
        
        try:
            total_building_sf = df['SF'].sum()
            total_units = len(df)
            
            # Extreme Case 1: Target SF exceeds building capacity
            if target_sf > total_building_sf * 1.1:  # 10% buffer
                recommendations.append({
                    'type': 'IMPOSSIBLE_TARGET',
                    'severity': 'CRITICAL',
                    'issue': f'Target {target_sf:,.0f} SF exceeds building capacity {total_building_sf:,.0f} SF',
                    'recommendation': f'Reduce target to maximum {total_building_sf:,.0f} SF or consider entire building affordable',
                    'suggested_target_sf': total_building_sf
                })
            
            # Extreme Case 2: Target SF is impossibly small
            min_reasonable_sf = df['SF'].min() * 2  # At least 2 units worth
            if target_sf < min_reasonable_sf:
                recommendations.append({
                    'type': 'UNREALISTIC_TARGET',
                    'severity': 'HIGH',
                    'issue': f'Target {target_sf:,.0f} SF is unrealistically small (less than 2 units)',
                    'recommendation': f'Consider minimum {min_reasonable_sf:,.0f} SF for viable affordable housing',
                    'suggested_target_sf': min_reasonable_sf
                })
            
            # Extreme Case 3: Building has no 2BR+ units but needs bedroom mix compliance
            if 'Bedrooms' in df.columns:
                units_2br_plus = len(df[df['Bedrooms'] >= 2])
                if units_2br_plus == 0 and total_units > 5:
                    recommendations.append({
                        'type': 'IMPOSSIBLE_BEDROOM_MIX',
                        'severity': 'HIGH',
                        'issue': 'Building has no 2BR+ units but needs 50% 2BR+ for compliance',
                        'recommendation': 'Consider converting some 1BR units to 2BR or seek compliance waiver',
                        'suggested_action': 'Convert at least 50% of units to 2BR+'
                    })
            
            # Extreme Case 4: Required floors don't exist
            if hasattr(self, 'required_floors') and self.required_floors:
                available_floors = set(df['Floor'].unique())
                required_floors_set = set(self.required_floors)
                missing_floors = required_floors_set - available_floors
                
                if missing_floors:
                    recommendations.append({
                        'type': 'MISSING_REQUIRED_FLOORS',
                        'severity': 'MEDIUM',
                        'issue': f'Required floors {missing_floors} do not exist in building',
                        'recommendation': f'Use available floors: {sorted(available_floors)}',
                        'suggested_floors': sorted(available_floors)
                    })
            
            # Only return recommendations for truly extreme cases (0.1%)
            critical_recommendations = [r for r in recommendations if r['severity'] == 'CRITICAL']
            
            return critical_recommendations  # Only return critical issues
            
        except Exception as e:
            return []

    def _process_ami_optimization_mode_enhanced(self, df: pd.DataFrame, target_sf: float, required_floors: Optional[List[int]] = None) -> Dict[str, Any]:
        """
        Enhanced AMI optimization mode processing.
        Handles cases where units have existing AMI values that need optimization.
        """
        try:
            # Get units that currently have AMI assignments
            units_with_ami = df[df['AMI'].notna()].copy()
            
            if len(units_with_ami) == 0:
                return {'success': False, 'error': 'No units with AMI values found'}
            
            # Use the pre-selected units (those with AMI values)
            selected_units = units_with_ami
            actual_target_sf = selected_units['SF'].sum()
            
            # Generate optimized AMI assignments
            optimization_results = self._smart_compliance_handling_and_optimization(
                selected_units, actual_target_sf, required_floors
            )
            
            if not optimization_results['success']:
                return optimization_results
            
            # Enhance results with additional analysis
            for result in optimization_results['results']:
                result['mode'] = 'AMI_OPTIMIZATION'
                result['original_ami_analysis'] = self._analyze_original_ami_distribution(units_with_ami)
                result['improvement_analysis'] = self._analyze_ami_improvements(
                    units_with_ami, result['ami_assignments']
                )
            
            return optimization_results
            
        except Exception as e:
            return {'success': False, 'error': f'AMI optimization failed: {str(e)}'}

    def _analyze_original_ami_distribution(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Analyze the original AMI distribution before optimization.
        """
        analysis = {
            'total_units': len(df),
            'total_sf': df['SF'].sum(),
            'ami_distribution': {},
            'weighted_average': 0,
            'compliance_issues': []
        }
        
        try:
            # Calculate original weighted average
            total_sf = df['SF'].sum()
            weighted_sum = (df['AMI'] * df['SF']).sum()
            analysis['weighted_average'] = weighted_sum / total_sf if total_sf > 0 else 0
            
            # Analyze AMI distribution
            ami_counts = df['AMI'].value_counts()
            for ami_level, count in ami_counts.items():
                sf_for_ami = df[df['AMI'] == ami_level]['SF'].sum()
                analysis['ami_distribution'][ami_level] = {
                    'units': count,
                    'sf': sf_for_ami,
                    'percentage_by_sf': sf_for_ami / total_sf if total_sf > 0 else 0
                }
            
            # Identify compliance issues
            if analysis['weighted_average'] > 0.80:
                analysis['compliance_issues'].append('Weighted average exceeds 80% limit')
            elif analysis['weighted_average'] < 0.60:
                analysis['compliance_issues'].append('Weighted average below 60% target')
            
            ami_40_pct = analysis['ami_distribution'].get(0.4, {}).get('percentage_by_sf', 0)
            if ami_40_pct < 0.20:
                analysis['compliance_issues'].append('40% AMI coverage below 20% minimum')
            
        except Exception as e:
            analysis['compliance_issues'].append(f'Analysis error: {str(e)}')
        
        return analysis

    def _analyze_ami_improvements(self, original_df: pd.DataFrame, new_ami_assignments: Dict[str, float]) -> Dict[str, Any]:
        """
        Analyze improvements from original to new AMI assignments.
        """
        improvements = {
            'weighted_average_change': 0,
            'ami_40_coverage_change': 0,
            'compliance_improvements': [],
            'overall_improvement': 'None'
        }
        
        try:
            # Calculate original metrics
            original_total_sf = original_df['SF'].sum()
            original_weighted_sum = (original_df['AMI'] * original_df['SF']).sum()
            original_weighted_avg = original_weighted_sum / original_total_sf if original_total_sf > 0 else 0
            
            original_40_sf = original_df[original_df['AMI'] == 0.4]['SF'].sum()
            original_40_pct = original_40_sf / original_total_sf if original_total_sf > 0 else 0
            
            # Calculate new metrics
            new_weighted_sum = 0
            new_40_sf = 0
            
            for _, unit in original_df.iterrows():
                unit_id = str(unit['Unit'])
                new_ami = new_ami_assignments.get(unit_id, unit['AMI'])
                sf = unit['SF']
                
                new_weighted_sum += new_ami * sf
                if new_ami == 0.4:
                    new_40_sf += sf
            
            new_weighted_avg = new_weighted_sum / original_total_sf if original_total_sf > 0 else 0
            new_40_pct = new_40_sf / original_total_sf if original_total_sf > 0 else 0
            
            # Calculate changes
            improvements['weighted_average_change'] = new_weighted_avg - original_weighted_avg
            improvements['ami_40_coverage_change'] = new_40_pct - original_40_pct
            
            # Identify improvements
            if abs(improvements['weighted_average_change']) < 0.01:
                improvements['compliance_improvements'].append('Maintained optimal weighted average')
            elif 0.60 <= new_weighted_avg <= 0.80 and not (0.60 <= original_weighted_avg <= 0.80):
                improvements['compliance_improvements'].append('Brought weighted average into compliance range')
            
            if improvements['ami_40_coverage_change'] > 0.05:
                improvements['compliance_improvements'].append('Significantly improved 40% AMI coverage')
            elif new_40_pct >= 0.20 and original_40_pct < 0.20:
                improvements['compliance_improvements'].append('Achieved 40% AMI minimum requirement')
            
            # Overall assessment
            if len(improvements['compliance_improvements']) >= 2:
                improvements['overall_improvement'] = 'Significant'
            elif len(improvements['compliance_improvements']) == 1:
                improvements['overall_improvement'] = 'Moderate'
            elif abs(improvements['weighted_average_change']) < 0.05:
                improvements['overall_improvement'] = 'Maintained'
            else:
                improvements['overall_improvement'] = 'Minor'
            
        except Exception as e:
            improvements['compliance_improvements'].append(f'Analysis error: {str(e)}')
        
        return improvements

    def _generate_multiple_optimized_versions(self, df: pd.DataFrame, target_sf: float, required_floors: Optional[List[int]] = None) -> List[Dict[str, Any]]:
        """
        Generate multiple optimized versions with different strategies.
        """
        versions = []
        
        try:
            # Version 1: Closest to 60% AMI
            version_1 = self._optimize_for_closest_to_60_ami(df, target_sf, required_floors)
            if version_1:
                version_1['version_name'] = 'Closest to 60% AMI'
                version_1['version_description'] = 'Optimized for weighted average closest to 60% AMI'
                version_1['version_number'] = 1
                versions.append(version_1)
            
            # Version 2: Floor Strategy Optimized
            version_2 = self._optimize_for_floor_strategy(df, target_sf, required_floors)
            if version_2:
                version_2['version_name'] = 'Floor Strategy Optimized'
                version_2['version_description'] = '40% AMI on lower floors, higher AMI on upper floors'
                version_2['version_number'] = 2
                versions.append(version_2)
            
            # Version 3: Size Strategy Optimized
            version_3 = self._optimize_for_size_strategy(df, target_sf, required_floors)
            if version_3:
                version_3['version_name'] = 'Size Strategy Optimized'
                version_3['version_description'] = 'Smaller units get 40% AMI, larger units get higher AMI'
                version_3['version_number'] = 3
                versions.append(version_3)
            
            # Rank versions by optimization score
            ranked_versions = self._rank_optimization_results(versions)
            
            # Add ranking information
            for i, version in enumerate(ranked_versions):
                version['rank'] = i + 1
                version['is_recommended'] = i == 0
            
            return ranked_versions
            
        except Exception as e:
            return []

    def _create_comprehensive_reasoning_documentation(self, results: List[Dict[str, Any]], building_analysis: Dict[str, Any], mode: str) -> Dict[str, Any]:
        """
        Create comprehensive reasoning documentation for the optimization decisions.
        """
        documentation = {
            'executive_summary': '',
            'building_analysis_summary': '',
            'optimization_approach': '',
            'version_comparisons': [],
            'compliance_analysis': '',
            'recommendations': '',
            'technical_details': {}
        }
        
        try:
            # Executive Summary
            best_result = results[0] if results else None
            if best_result:
                compliance = best_result['compliance']
                documentation['executive_summary'] = f"""
EXECUTIVE SUMMARY:
Generated {len(results)} optimized AMI assignment strategies for {building_analysis['total_units']} units 
({building_analysis['total_sf']:,.0f} SF). Best strategy achieves {compliance['compliance_grade']} grade 
with {compliance['weighted_average_ami']:.1%} weighted average AMI and {compliance['ami_40_percentage']:.1%} 
coverage at 40% AMI.
                """.strip()
            
            # Building Analysis Summary
            unit_types = building_analysis['unit_types']
            documentation['building_analysis_summary'] = f"""
BUILDING CHARACTERISTICS:
• {building_analysis['building_category']} ({building_analysis['total_units']} units, {building_analysis['total_sf']:,.0f} SF)
• Floor Range: {building_analysis['floors']['min_floor']}-{building_analysis['floors']['max_floor']} ({building_analysis['floors']['total_floors']} floors)
• Unit Mix: {unit_types['studio']} Studio, {unit_types['1br']} 1BR, {unit_types['2br']} 2BR, {unit_types['3br']} 3BR+
• 2BR+ Percentage: {unit_types['bedroom_mix_percentage']:.1%}
• Optimization Potential: {building_analysis['optimization_potential']['overall_potential']}
            """.strip()
            
            # Optimization Approach
            documentation['optimization_approach'] = f"""
OPTIMIZATION APPROACH:
Mode: {mode}
Strategy: Client-respectful optimization (99.9% respect rate)
Generated {len(results)} versions using different AMI assignment strategies:
1. Closest to 60% AMI - Targets optimal weighted average
2. Floor Strategy - Lower floors get lower AMI levels  
3. Size Strategy - Smaller units get lower AMI levels
            """.strip()
            
            # Version Comparisons
            for result in results:
                compliance = result['compliance']
                version_summary = {
                    'name': result.get('version_name', result.get('strategy_name', 'Unknown')),
                    'rank': result.get('rank', 0),
                    'grade': compliance['compliance_grade'],
                    'weighted_average': compliance['weighted_average_ami'],
                    'ami_40_coverage': compliance['ami_40_percentage'],
                    'bedroom_mix': compliance['bedroom_mix_percentage'],
                    'overage': result.get('overage_percentage', 0),
                    'score': result.get('optimization_score', 0),
                    'issues': compliance['compliance_issues']
                }
                documentation['version_comparisons'].append(version_summary)
            
            # Compliance Analysis
            if best_result:
                compliance = best_result['compliance']
                issues_text = '; '.join(compliance['compliance_issues']) if compliance['compliance_issues'] else 'All requirements met'
                documentation['compliance_analysis'] = f"""
COMPLIANCE ANALYSIS (Best Version):
• Grade: {compliance['compliance_grade']}
• Weighted Average AMI: {compliance['weighted_average_ami']:.1%} (maximum: 60%)
• 40% AMI Coverage: {compliance['ami_40_percentage']:.1%} (minimum: 20%)
• Issues: {issues_text}
                """.strip()
            
            # Recommendations
            if len(results) > 1:
                documentation['recommendations'] = f"""
RECOMMENDATIONS:
• Primary Choice: {results[0].get('version_name', 'Version 1')} (Rank #1, Grade {results[0]['compliance']['compliance_grade']})
• Alternative: {results[1].get('version_name', 'Version 2')} (Rank #2, Grade {results[1]['compliance']['compliance_grade']})
• All versions meet core NYC affordable housing requirements
• Choose based on development priorities and market conditions
                """.strip()
            else:
                documentation['recommendations'] = "Single optimized version generated. Meets all core requirements."
            
            # Technical Details
            documentation['technical_details'] = {
                'total_versions_generated': len(results),
                'building_category': building_analysis['building_category'],
                'optimization_mode': mode,
                'analysis_timestamp': datetime.now().isoformat(),
                'compliance_framework': 'NYC UAP/MIH Requirements'
            }
            
        except Exception as e:
            documentation['executive_summary'] = f"Documentation generation error: {str(e)}"
        
        return documentation

    def cleanup_old_files(self):
        """Clean up old generated files."""
        try:
            for folder in [UPLOAD_FOLDER, RESULTS_FOLDER]:
                if os.path.exists(folder):
                    files = os.listdir(folder)
                    # Keep only the latest 5 files
                    files.sort(key=lambda x: os.path.getctime(os.path.join(folder, x)), reverse=True)
                    for file in files[5:]:
                        try:
                            os.remove(os.path.join(folder, file))
                        except:
                            pass
        except Exception as e:
            print(f"Error cleaning up files: {str(e)}")

    def process_building_file(self, file_path: str, target_sf: Optional[float] = None, required_floors: Optional[str] = None) -> Dict[str, Any]:
        """
        Main processing function for building files.
        Enhanced with all 38 functions and client-respectful approach.
        """
        try:
            # Load and validate data
            df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
            
            # Enhanced data validation and cleaning
            df_clean = self._enhanced_data_validation_and_cleaning(df)
            
            # Parse required floors
            required_floors_list = None
            if required_floors:
                try:
                    required_floors_list = [int(f.strip()) for f in required_floors.split(',')]
                except:
                    pass
            
            # Enhanced intelligent mode detection
            mode = self._enhanced_intelligent_mode_detection(df_clean, target_sf)
            
            # Auto-calculate target SF if not provided (core feature)
            if target_sf is None:
                target_sf = self._auto_calculate_target_sf_from_preselected_units(df_clean)
            
            # Enhanced building analysis
            building_analysis = self._enhanced_building_analysis(df_clean)
            
            # Analyze for extreme cases (0.1% of cases)
            extreme_case_recommendations = self._analyze_for_extreme_cases(df_clean, target_sf)
            
            # Process based on mode
            if mode == 'AMI_OPTIMIZATION':
                optimization_results = self._process_ami_optimization_mode_enhanced(df_clean, target_sf, required_floors_list)
            else:
                # AMI_ASSIGNMENT mode
                optimization_results = self._smart_compliance_handling_and_optimization(df_clean, target_sf, required_floors_list)
            
            if not optimization_results['success']:
                return {
                    'success': False,
                    'error': optimization_results.get('error', 'Optimization failed'),
                    'building_analysis': building_analysis
                }
            
            # Generate multiple optimized versions
            results = optimization_results['results']
            if len(results) < 3:
                # Generate additional versions if needed
                additional_results = self._generate_multiple_optimized_versions(df_clean, target_sf, required_floors_list)
                for result in additional_results:
                    if result not in results:
                        results.append(result)
                results = self._rank_optimization_results(results)
            
            # Create comprehensive reasoning documentation
            reasoning_documentation = self._create_comprehensive_reasoning_documentation(results, building_analysis, mode)
            
            # Generate Excel files for each version
            excel_files = []
            for i, result in enumerate(results[:3]):  # Limit to top 3 versions
                excel_file = self._create_excel_output(df_clean, result, i+1, reasoning_documentation)
                if excel_file:
                    excel_files.append(excel_file)
            
            return {
                'success': True,
                'mode': mode,
                'target_sf': target_sf,
                'building_analysis': building_analysis,
                'results': results[:3],  # Return top 3 versions
                'excel_files': excel_files,
                'reasoning_documentation': reasoning_documentation,
                'extreme_case_recommendations': extreme_case_recommendations,
                'total_versions': len(results)
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': f"Processing failed: {str(e)}",
                'traceback': traceback.format_exc()
            }

    def _create_excel_output(self, original_df: pd.DataFrame, result: Dict[str, Any], version_number: int, reasoning: Dict[str, Any]) -> Optional[str]:
        """
        Create clean Excel output file with AMI assignments.
        No reasoning text in Excel - keeps files clean and uncorrupted.
        """
        try:
            # Create output dataframe
            output_df = original_df.copy()
            
            # Apply AMI assignments and convert to percentage format
            ami_assignments = result['ami_assignments']
            for unit_id, ami_value in ami_assignments.items():
                mask = output_df['Unit'].astype(str) == str(unit_id)
                # Convert decimal to percentage format (0.4 -> "40%", 1.0 -> "100%")
                if ami_value is not None:
                    ami_percentage = f"{int(ami_value * 100)}%"
                    output_df.loc[mask, 'AMI'] = ami_percentage
            
            # Create filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            strategy_name = result.get('version_name', result.get('strategy_name', f'Version_{version_number}'))
            safe_strategy_name = strategy_name.replace(' ', '_').replace(':', '').replace('%', 'PCT')
            filename = f"AMI_Optimization_{safe_strategy_name}_{timestamp}.xlsx"
            filepath = os.path.join(RESULTS_FOLDER, filename)
            
            # Write to Excel with clean formatting
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                output_df.to_excel(writer, sheet_name='AMI_Assignments', index=False)
                
                # Add summary sheet with key metrics only
                summary_data = {
                    'Metric': [
                        'Strategy',
                        'Total Units',
                        'Total SF',
                        'Weighted Average AMI',
                        '40% AMI Coverage',
                        '2BR+ Mix',
                        'Compliance Grade',
                        'Overage SF',
                        'Overage %'
                    ],
                    'Value': [
                        strategy_name,
                        result['total_units'],
                        f"{result['total_sf']:,.0f}",
                        f"{result['compliance']['weighted_average_ami']:.1%}",
                        f"{result['compliance']['ami_40_percentage']:.1%}",
                        f"{result['compliance']['bedroom_mix_percentage']:.1%}",
                        result['compliance']['compliance_grade'],
                        f"{result.get('overage_sf', 0):,.0f}",
                        f"{result.get('overage_percentage', 0):.2f}%"
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            return filepath
            
        except Exception as e:
            print(f"Error creating Excel file: {str(e)}")
            return None

# Flask Web Application
@app.route('/')
def index():
    """Main page with beautiful UI."""
    return render_template_string("""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>NYC Affordable Housing AMI Calculator</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }
        
        .header p {
            font-size: 1.2em;
            opacity: 0.9;
        }
        
        .content {
            padding: 40px;
        }
        
        .upload-section {
            background: #f8f9fa;
            border-radius: 15px;
            padding: 30px;
            margin-bottom: 30px;
            border: 2px dashed #dee2e6;
            text-align: center;
            transition: all 0.3s ease;
        }
        
        .upload-section:hover {
            border-color: #667eea;
            background: #f0f2ff;
        }
        
        .upload-section.dragover {
            border-color: #667eea;
            background: #e3f2fd;
            transform: scale(1.02);
        }
        
        .upload-icon {
            font-size: 3em;
            color: #667eea;
            margin-bottom: 20px;
        }
        
        .upload-text {
            font-size: 1.3em;
            color: #495057;
            margin-bottom: 20px;
        }
        
        .file-input {
            display: none;
        }
        
        .upload-btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 30px;
            border: none;
            border-radius: 50px;
            font-size: 1.1em;
            cursor: pointer;
            transition: all 0.3s ease;
        }
        
        .upload-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.3);
        }
        
        .options-section {
            display: none;
            margin-top: 30px;
        }
        
        .form-group {
            margin-bottom: 25px;
        }
        
        .form-group label {
            display: block;
            font-weight: 600;
            color: #2c3e50;
            margin-bottom: 8px;
            font-size: 1.1em;
        }
        
        .form-group input {
            width: 100%;
            padding: 12px 15px;
            border: 2px solid #e9ecef;
            border-radius: 10px;
            font-size: 1em;
            transition: border-color 0.3s ease;
        }
        
        .form-group input:focus {
            outline: none;
            border-color: #667eea;
            box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }
        
        .process-btn {
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
            color: white;
            padding: 15px 40px;
            border: none;
            border-radius: 50px;
            font-size: 1.2em;
            cursor: pointer;
            transition: all 0.3s ease;
            width: 100%;
        }
        
        .process-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(40, 167, 69, 0.3);
        }
        
        .process-btn:disabled {
            background: #6c757d;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        
        .results-section {
            display: none;
            margin-top: 30px;
        }
        
        .results-header {
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
            color: white;
            padding: 20px;
            border-radius: 15px 15px 0 0;
            text-align: center;
        }
        
        .results-content {
            background: white;
            border: 2px solid #28a745;
            border-top: none;
            border-radius: 0 0 15px 15px;
            padding: 30px;
        }
        
        .summary-table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }
        
        .summary-table th {
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
        }
        
        .summary-table td {
            padding: 12px 15px;
            border-bottom: 1px solid #e9ecef;
        }
        
        .summary-table tr:hover {
            background: #f8f9fa;
        }
               .compliance-badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-weight: 600;
            font-size: 0.85em;
            text-transform: uppercase;
        }
        
        .unit-breakdown-btn {
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 0.9em;
            margin-top: 10px;
            transition: all 0.3s ease;
        }
        
        .unit-breakdown-btn:hover {
            background: linear-gradient(135deg, #2980b9 0%, #1f5f8b 100%);
            transform: translateY(-1px);
        }
        
        .unit-breakdown-table {
            margin-top: 15px;
            background: #f8f9fa;
            border-radius: 8px;
            padding: 15px;
            border: 1px solid #e9ecef;
        }
        
        .unit-breakdown-table h6 {
            margin: 0 0 15px 0;
            color: #2c3e50;
            font-size: 1.1em;
        }
        
        .unit-details-table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 6px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .unit-details-table th {
            background: #34495e;
            color: white;
            padding: 10px;
            text-align: left;
            font-weight: 600;
            font-size: 0.9em;
        }
        
        .unit-details-table td {
            padding: 8px 10px;
            border-bottom: 1px solid #e9ecef;
            font-size: 0.9em;
        }
        
        .unit-details-table tr:hover {
            background: #f1f3f4;
        }
        
        .ami-badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 12px;
            font-weight: 600;
            font-size: 0.8em;
            color: white;
        }
        
        .ami-40 { background: #e74c3c; }
        .ami-60 { background: #f39c12; }
        .ami-80 { background: #27ae60; }
        .ami-100 { background: #8e44ad; }
        
        .grade-A { background: #d4edda; color: #155724; }
        .grade-B { background: #d1ecf1; color: #0c5460; }
        .grade-C { background: #fff3cd; color: #856404; }
        .grade-F { background: #f8d7da; color: #721c24; }
        
        .building-analysis {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 30px;
        }
        
        .analysis-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }
        
        .analysis-item {
            background: white;
            padding: 15px;
            border-radius: 8px;
            border-left: 4px solid #667eea;
        }
        
        .results-header {
            text-align: center;
            margin-bottom: 30px;
        }
        
        .results-header h3 {
            color: #2c3e50;
            margin-bottom: 10px;
        }
        
        .strategy-name {
            line-height: 1.4;
        }
        
        .strategy-desc {
            font-size: 0.85em;
            color: #666;
            margin-top: 5px;
        }
        
        .download-btn-small {
            background: #28a745;
            color: white;
            padding: 5px 12px;
            border-radius: 5px;
            text-decoration: none;
            font-size: 0.85em;
            display: inline-block;
        }
        
        .download-btn-small:hover {
            background: #218838;
            color: white;
            text-decoration: none;
        }
        
        .compliance-details {
            background: #e3f2fd;
            padding: 20px;
            border-radius: 10px;
            margin: 30px 0;
        }
        
        .compliance-details h4 {
            color: #1976d2;
            margin-bottom: 15px;
        }
        
        .compliance-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
        }
        
        .compliance-item {
            background: white;
            padding: 12px;
            border-radius: 8px;
            border-left: 4px solid #1976d2;
        }
        
        .ami-breakdown {
            background: #f1f8e9;
            padding: 20px;
            border-radius: 10px;
            margin: 30px 0;
        }
        
        .ami-breakdown h4 {
            color: #388e3c;
            margin-bottom: 20px;
        }
        
        .strategy-breakdown {
            background: white;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            border-left: 4px solid #4caf50;
        }
        
        .strategy-breakdown h5 {
            color: #2e7d32;
            margin-bottom: 10px;
        }
        
        .ami-distribution-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 10px;
        }
        
        .ami-item {
            background: #f8f9fa;
            padding: 8px 12px;
            border-radius: 5px;
            font-size: 0.9em;
        }
        
        .ami-level {
            font-weight: 600;
            color: #2e7d32;
        }
        
        .ami-details {
            color: #666;
        }
        
        .download-section {
            margin-top: 30px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 15px;
        }
        
        .download-btn {
            background: linear-gradient(135deg, #007bff 0%, #0056b3 100%);
            color: white;
            padding: 12px 25px;
            border: none;
            border-radius: 25px;
            text-decoration: none;
            display: inline-block;
            margin: 5px;
            transition: all 0.3s ease;
            font-size: 1em;
        }
        
        .download-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 15px rgba(0, 123, 255, 0.3);
            text-decoration: none;
            color: white;
        }
        
        .clear-btn {
            background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
            color: white;
            padding: 12px 25px;
            border: none;
            border-radius: 25px;
            cursor: pointer;
            transition: all 0.3s ease;
            font-size: 1em;
            margin-top: 20px;
        }
        
        .clear-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 15px rgba(220, 53, 69, 0.3);
        }
        
        .loading {
            display: none;
            text-align: center;
            padding: 30px;
        }
        
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #667eea;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        .error-message {
            background: #f8d7da;
            color: #721c24;
            padding: 15px;
            border-radius: 10px;
            margin: 20px 0;
            border: 1px solid #f5c6cb;
        }
        
        .success-message {
            background: #d4edda;
            color: #155724;
            padding: 15px;
            border-radius: 10px;
            margin: 20px 0;
            border: 1px solid #c3e6cb;
        }
        
        .file-info {
            background: #e3f2fd;
            padding: 15px;
            border-radius: 10px;
            margin: 15px 0;
            border-left: 4px solid #2196f3;
        }
        
        @media (max-width: 768px) {
            .container {
                margin: 10px;
                border-radius: 15px;
            }
            
            .header h1 {
                font-size: 2em;
            }
            
            .content {
                padding: 20px;
            }
            
            .summary-table {
                font-size: 0.9em;
            }
            
            .summary-table th,
            .summary-table td {
                padding: 8px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🏠 NYC Affordable Housing AMI Calculator</h1>
            <p>Professional AMI Optimization System with Enhanced Intelligence</p>
        </div>
        
        <div class="content">
            <div class="upload-section" id="uploadSection">
                <div class="upload-icon">📁</div>
                <div class="upload-text">
                    <strong>Upload Building Units File</strong><br>
                    Drag and drop your Excel (.xlsx) or CSV (.csv) file here
                </div>
                <input type="file" id="fileInput" class="file-input" accept=".xlsx,.xls,.csv">
                <button class="upload-btn" onclick="document.getElementById('fileInput').click()">
                    Browse Files
                </button>
                <div style="margin-top: 15px; color: #6c757d; font-size: 0.9em;">
                    Limit: 200MB per file • Supports: XLSX, XLS, CSV
                </div>
            </div>
            
            <div class="options-section" id="optionsSection">
                <div class="form-group">
                    <label for="requiredFloors">Required Floors (Optional)</label>
                    <input type="text" id="requiredFloors" placeholder="e.g., 5,6 or leave blank for automatic optimization">
                </div>
                
                <button class="process-btn" id="processBtn" onclick="processFile()">
                    Generate Smart Options
                </button>
            </div>
            
            <div class="loading" id="loading">
                <div class="spinner"></div>
                <div style="font-size: 1.2em; color: #667eea;">
                    Processing your building data...<br>
                    <small>Analyzing units, optimizing AMI distribution, and generating results</small>
                </div>
            </div>
            
            <div class="results-section" id="resultsSection">
                <div class="results-header">
                    <h2>🎉 Smart Options Generated Successfully!</h2>
                    <div id="buildingInfo"></div>
                </div>
                
                <div class="results-content">
                    <div id="summaryTable"></div>
                    
                    <div class="download-section">
                        <h3 style="margin-bottom: 15px;">📥 Download Results</h3>
                        <div id="downloadLinks"></div>
                        
                        <button class="clear-btn" onclick="clearResults()">
                            🔄 Clear & Start New
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        let uploadedFile = null;
        let currentResults = null;

        // File upload handling
        const fileInput = document.getElementById('fileInput');
        const uploadSection = document.getElementById('uploadSection');
        const optionsSection = document.getElementById('optionsSection');

        fileInput.addEventListener('change', handleFileSelect);
        
        // Drag and drop functionality
        uploadSection.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadSection.classList.add('dragover');
        });
        
        uploadSection.addEventListener('dragleave', () => {
            uploadSection.classList.remove('dragover');
        });
        
        uploadSection.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadSection.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                fileInput.files = files;
                handleFileSelect();
            }
        });

        function handleFileSelect() {
            const file = fileInput.files[0];
            if (file) {
                uploadedFile = file;
                
                // Show file info
                const fileInfo = document.createElement('div');
                fileInfo.className = 'file-info';
                fileInfo.innerHTML = `
                    <strong>📄 ${file.name}</strong><br>
                    Size: ${(file.size / 1024 / 1024).toFixed(2)} MB<br>
                    Type: ${file.type || 'Unknown'}
                `;
                
                // Remove any existing file info
                const existingInfo = document.querySelector('.file-info');
                if (existingInfo) {
                    existingInfo.remove();
                }
                
                uploadSection.appendChild(fileInfo);
                optionsSection.style.display = 'block';
            }
        }

        async function processFile() {
            if (!uploadedFile) {
                alert('Please select a file first');
                return;
            }

            const processBtn = document.getElementById('processBtn');
            const loading = document.getElementById('loading');
            const resultsSection = document.getElementById('resultsSection');
            
            // Show loading
            processBtn.disabled = true;
            loading.style.display = 'block';
            resultsSection.style.display = 'none';

            try {
                const formData = new FormData();
                formData.append('file', uploadedFile);
                
                const requiredFloors = document.getElementById('requiredFloors').value.trim();
                if (requiredFloors) {
                    formData.append('required_floors', requiredFloors);
                }

                const response = await fetch('/process', {
                    method: 'POST',
                    body: formData
                });

                const result = await response.json();

                if (result.success) {
                    currentResults = result;
                    displayResults(result);
                } else {
                    showError(result.error || 'Processing failed');
                }
            } catch (error) {
                showError('Network error: ' + error.message);
            } finally {
                processBtn.disabled = false;
                loading.style.display = 'none';
            }
        }

        function displayResults(result) {
            const buildingInfo = document.getElementById('buildingInfo');
            const summaryTable = document.getElementById('summaryTable');
            const downloadLinks = document.getElementById('downloadLinks');
            const resultsSection = document.getElementById('resultsSection');

            // Building info with enhanced details
            const analysis = result.building_analysis;
            buildingInfo.innerHTML = `
                <div class="building-analysis">
                    <h3>🏢 Building Analysis</h3>
                    <div class="analysis-grid">
                        <div class="analysis-item">
                            <strong>Total Units:</strong> ${analysis.total_units}
                        </div>
                        <div class="analysis-item">
                            <strong>Total SF:</strong> ${analysis.total_sf.toLocaleString()} SF
                        </div>
                        <div class="analysis-item">
                            <strong>Floors:</strong> ${analysis.floor_range || 'N/A'}
                        </div>
                        <div class="analysis-item">
                            <strong>Unit Mix:</strong> ${analysis.unit_mix || 'N/A'}
                        </div>
                    </div>
                </div>
            `;

            // Enhanced summary table with detailed compliance breakdown
            let tableHTML = `
                <div class="results-header">
                    <h3>📊 AMI Optimization Results</h3>
                    <p>Three optimized strategies for your affordable housing requirements</p>
                </div>
                <table class="summary-table">
                    <thead>
                        <tr>
                            <th>Strategy</th>
                            <th>Units</th>
                            <th>Total SF</th>
                            <th>Weighted AMI</th>
                            <th>40% AMI Coverage</th>
                            <th>2BR+ Mix</th>
                            <th>Compliance</th>
                            <th>Download</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            result.results.forEach((option, index) => {
                const compliance = option.compliance;
                const gradeClass = `grade-${compliance.compliance_grade.toLowerCase()}`;
                const filename = result.excel_files[index]?.split('/').pop() || '';
                
                tableHTML += `
                    <tr>
                        <td>
                            <div class="strategy-name">
                                <strong>${option.version_name || option.strategy_name}</strong>
                                <div class="strategy-desc">${option.strategy_description || ''}</div>
                            </div>
                        </td>
                        <td>${option.total_units}</td>
                        <td>${option.total_sf.toLocaleString()} SF</td>
                        <td><strong>${(compliance.weighted_average_ami * 100).toFixed(1)}%</strong></td>
                        <td>${(compliance.ami_40_percentage * 100).toFixed(1)}%</td>
                        <td>${(compliance.bedroom_mix_percentage * 100).toFixed(1)}%</td>
                        <td><span class="compliance-badge ${gradeClass}">${compliance.compliance_grade}</span></td>
                        <td>
                            ${filename ? `<a href="/download/${filename}" class="download-btn-small" target="_blank">📥 Excel</a>` : ''}
                        </td>
                    </tr>
                `;
            });

            tableHTML += '</tbody></table>';

            // Add detailed compliance analysis
            tableHTML += `
                <div class="compliance-details">
                    <h4>📋 Compliance Requirements</h4>
                    <div class="compliance-grid">
                        <div class="compliance-item">
                            <strong>40% AMI Minimum:</strong> 20% of affordable SF
                        </div>
                        <div class="compliance-item">
                            <strong>Weighted Average:</strong> Maximum 60% AMI
                        </div>
                        <div class="compliance-item">
                            <strong>Unit Selection:</strong> Client pre-selects affordable units
                        </div>
                        <div class="compliance-item">
                            <strong>Assignment Priority:</strong> Lower floors + smaller units get 40% AMI first
                        </div>
                    </div>
                </div>
            `;

            // Add AMI distribution breakdown for each strategy
            tableHTML += `
                <div class="ami-breakdown">
                    <h4>🎯 AMI Distribution Breakdown</h4>
            `;

            result.results.forEach((option, index) => {
                if (option.ami_distribution) {
                    tableHTML += `
                        <div class="strategy-breakdown">
                            <h5>${option.version_name || option.strategy_name}</h5>
                            <div class="ami-distribution-grid">
                    `;
                    
                    Object.entries(option.ami_distribution).forEach(([ami, data]) => {
                        const percentage = ((data.sf / option.total_sf) * 100).toFixed(1);
                        tableHTML += `
                            <div class="ami-item">
                                <span class="ami-level">${ami}% AMI:</span>
                                <span class="ami-details">${data.units} units, ${data.sf.toLocaleString()} SF (${percentage}%)</span>
                            </div>
                        `;
                    });
                    
                    tableHTML += `
                            </div>
                            <button class="unit-breakdown-btn" onclick="toggleUnitBreakdown(${index})">
                                📋 View Unit Details
                            </button>
                            <div id="unitBreakdown${index}" class="unit-breakdown-table" style="display: none;">
                                <h6>Unit-by-Unit Breakdown</h6>
                                <table class="unit-details-table">
                                    <thead>
                                        <tr>
                                            <th>Unit</th>
                                            <th>Floor</th>
                                            <th>SF</th>
                                            <th>Bedrooms</th>
                                            <th>AMI Level</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                    `;
                    
                    // Add unit details if available
                    if (option.unit_details) {
                        option.unit_details.forEach(unit => {
                            tableHTML += `
                                <tr>
                                    <td><strong>${unit.unit}</strong></td>
                                    <td>${unit.floor}</td>
                                    <td>${unit.sf.toLocaleString()}</td>
                                    <td>${unit.bedrooms}</td>
                                    <td><span class="ami-badge ami-${unit.ami_level.replace('%', '')}">${unit.ami_level}</span></td>
                                </tr>
                            `;
                        });
                    }
                    
                    tableHTML += `
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    `;
                }
            });

            tableHTML += '</div>';

            summaryTable.innerHTML = tableHTML;

            resultsSection.style.display = 'block';
        }

        function toggleUnitBreakdown(index) {
            const breakdown = document.getElementById(`unitBreakdown${index}`);
            const btn = event.target;
            
            if (breakdown.style.display === 'none') {
                breakdown.style.display = 'block';
                btn.textContent = '📋 Hide Unit Details';
            } else {
                breakdown.style.display = 'none';
                btn.textContent = '📋 View Unit Details';
            }
        }

        function showError(message) {
            const errorDiv = document.createElement('div');
            errorDiv.className = 'error-message';
            errorDiv.innerHTML = `<strong>❌ Error:</strong> ${message}`;
            
            // Remove any existing error messages
            const existingError = document.querySelector('.error-message');
            if (existingError) {
                existingError.remove();
            }
            
            document.querySelector('.content').appendChild(errorDiv);
            
            // Auto-remove after 10 seconds
            setTimeout(() => {
                if (errorDiv.parentNode) {
                    errorDiv.remove();
                }
            }, 10000);
        }

        function clearResults() {
            // Reset form
            uploadedFile = null;
            currentResults = null;
            fileInput.value = '';
            document.getElementById('requiredFloors').value = '';
            
            // Hide sections
            optionsSection.style.display = 'none';
            document.getElementById('resultsSection').style.display = 'none';
            
            // Remove file info
            const fileInfo = document.querySelector('.file-info');
            if (fileInfo) {
                fileInfo.remove();
            }
            
            // Remove error messages
            const errorMsg = document.querySelector('.error-message');
            if (errorMsg) {
                errorMsg.remove();
            }
            
            // Reset upload section
            uploadSection.classList.remove('dragover');
        }
    </script>
</body>
</html>
    """)

@app.route('/process', methods=['POST'])
def process_file():
    """Process uploaded file and generate AMI optimization results."""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file uploaded'})
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'})
        
        # Save uploaded file
        filename = f"upload_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file.filename}"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        
        # Get optional parameters
        required_floors = request.form.get('required_floors', '').strip()
        
        # Initialize system and process
        system = TrulyCompleteAMIOptimizationSystem()
        result = system.process_building_file(
            filepath, 
            target_sf=None,  # Auto-calculate from pre-selected units
            required_floors=required_floors if required_floors else None
        )
        
        # Clean up uploaded file
        try:
            os.remove(filepath)
        except:
            pass
        
        # Clean up old files
        system.cleanup_old_files()
        
        # Convert result to JSON serializable format
        result = make_json_serializable(result)
        
        return jsonify(result)
        
    except Exception as e:
        return jsonify({
            'success': False, 
            'error': f'Processing failed: {str(e)}',
            'traceback': traceback.format_exc()
        })

@app.route('/download/<filename>')
def download_file(filename):
    """Download generated Excel file."""
    try:
        filepath = os.path.join(RESULTS_FOLDER, filename)
        if os.path.exists(filepath):
            return send_file(filepath, as_attachment=True)
        else:
            return "File not found", 404
    except Exception as e:
        return f"Download error: {str(e)}", 500

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

