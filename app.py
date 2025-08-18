#!/usr/bin/env python3
"""
Complete Universal Guaranteed AMI Calculator Web Application
===========================================================

Enhanced version with full calculation results and detailed strategies.
"""

from flask import Flask, render_template_string, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
import numpy as np
import json
from datetime import datetime
import warnings
import io
import os
from itertools import combinations
warnings.filterwarnings('ignore')

app = Flask(__name__)
CORS(app)

class EnhancedAMICalculator:
    def __init__(self):
        self.ami_levels = [0.40, 0.50, 0.60, 0.70, 0.80, 0.90, 1.00]
        
        self.scenario_presets = {
            'conservative': {
                'weighted_ami_range': (57.0, 58.0),
                '40_ami_range': (22.0, 23.0),
                'description': 'Safe, compliant approach'
            },
            'balanced': {
                'weighted_ami_range': (58.0, 60.0),
                '40_ami_range': (20.0, 22.0),
                'description': 'Optimal balance of revenue and compliance'
            },
            'aggressive': {
                'weighted_ami_range': (59.0, 60.0),
                '40_ami_range': (20.0, 21.0),
                'description': 'Maximum revenue optimization'
            },
            'ultra_precise': {
                'weighted_ami_range': (59.5, 60.0),
                '40_ami_range': (20.0, 20.5),
                'description': 'Extremely tight tolerances'
            }
        }
    
    def _clean_data(self, df):
        """Clean and validate the input data."""
        try:
            df_clean = df.copy()
            
            # Find SF column with better detection
            sf_columns = []
            for col in df_clean.columns:
                col_upper = str(col).upper().strip()
                if any(sf_term in col_upper for sf_term in ['SF', 'SQFT', 'SQUARE', 'SQ FT', 'NET SF', ' SF']):
                    sf_columns.append(col)
            
            if not sf_columns:
                raise ValueError("No square footage column found. Expected columns like 'SF', 'NET SF', 'SQFT'")
            
            sf_col = sf_columns[0]
            
            # Find other columns
            floor_columns = []
            for col in df_clean.columns:
                col_upper = str(col).upper().strip()
                if any(floor_term in col_upper for floor_term in ['FLOOR', 'LEVEL', 'FLR', 'STORY']):
                    floor_columns.append(col)
            
            unit_columns = []
            for col in df_clean.columns:
                col_upper = str(col).upper().strip()
                if any(unit_term in col_upper for unit_term in ['APT', 'UNIT', 'APARTMENT', 'SUITE']):
                    unit_columns.append(col)
            
            floor_col = floor_columns[0] if floor_columns else 'Floor'
            unit_col = unit_columns[0] if unit_columns else 'Apartment'
            
            # Standardize column names
            df_clean = df_clean.rename(columns={
                sf_col: 'SF',
                floor_col: 'Floor',
                unit_col: 'Apartment'
            })
            
            # Create missing columns
            if 'Floor' not in df_clean.columns:
                df_clean['Floor'] = range(1, len(df_clean) + 1)
            if 'Apartment' not in df_clean.columns:
                df_clean['Apartment'] = [f"Unit_{i+1}" for i in range(len(df_clean))]
            
            # Clean data
            df_clean['SF'] = pd.to_numeric(df_clean['SF'], errors='coerce')
            df_clean['Floor'] = pd.to_numeric(df_clean['Floor'], errors='coerce').fillna(1)
            
            # Remove invalid rows
            df_clean = df_clean.dropna(subset=['SF'])
            df_clean = df_clean[df_clean['SF'] > 0]
            
            if len(df_clean) == 0:
                raise ValueError("No valid units found after data cleaning")
            
            return df_clean
            
        except Exception as e:
            raise ValueError(f"Data validation failed: {str(e)}")
    
    def _calculate_weighted_ami(self, df):
        """Calculate weighted average AMI."""
        if 'AMI_Level' not in df.columns or len(df) == 0:
            return 0.0
        
        total_sf = df['SF'].sum()
        if total_sf == 0:
            return 0.0
        
        weighted_sum = (df['SF'] * df['AMI_Level']).sum()
        return (weighted_sum / total_sf) * 100
    
    def _calculate_40_ami_coverage(self, df):
        """Calculate percentage of SF at 40% AMI."""
        if 'AMI_Level' not in df.columns or len(df) == 0:
            return 0.0
        
        total_sf = df['SF'].sum()
        if total_sf == 0:
            return 0.0
        
        ami_40_sf = df[df['AMI_Level'] == 0.40]['SF'].sum()
        return (ami_40_sf / total_sf) * 100
    
    def _is_compliant(self, df, weighted_range, ami_40_range):
        """Check if assignment meets compliance requirements."""
        weighted_ami = self._calculate_weighted_ami(df)
        ami_40_coverage = self._calculate_40_ami_coverage(df)
        
        # STRICT 20% MINIMUM ENFORCEMENT
        min_40_ami = max(20.0, ami_40_range[0])
        
        weighted_ok = weighted_range[0] <= weighted_ami <= weighted_range[1]
        ami_40_ok = min_40_ami <= ami_40_coverage <= ami_40_range[1]
        
        return weighted_ok and ami_40_ok
    
    def _generate_all_combinations(self, df, weighted_range, ami_40_range):
        """Generate and test all possible AMI combinations."""
        results = []
        total_units = len(df)
        
        # Test different numbers of units at 40% AMI
        for num_40_units in range(1, min(total_units, 12) + 1):
            # Test different combinations of units for 40% AMI
            for units_40_combo in combinations(range(total_units), num_40_units):
                # Test different distributions for remaining units
                remaining_units = [i for i in range(total_units) if i not in units_40_combo]
                
                if len(remaining_units) == 0:
                    continue
                
                # Try different numbers of 80% AMI units
                for num_80_units in range(0, min(len(remaining_units), 6) + 1):
                    if num_80_units > 0:
                        for units_80_combo in combinations(remaining_units, num_80_units):
                            df_test = df.copy()
                            
                            # Assign AMI levels
                            df_test['AMI_Level'] = 0.60  # Default
                            
                            # Assign 40% AMI
                            for idx in units_40_combo:
                                df_test.iloc[idx, df_test.columns.get_loc('AMI_Level')] = 0.40
                            
                            # Assign 80% AMI
                            for idx in units_80_combo:
                                df_test.iloc[idx, df_test.columns.get_loc('AMI_Level')] = 0.80
                            
                            # Calculate metrics
                            weighted_ami = self._calculate_weighted_ami(df_test)
                            ami_40_coverage = self._calculate_40_ami_coverage(df_test)
                            is_compliant = self._is_compliant(df_test, weighted_range, ami_40_range)
                            
                            # Store result
                            result = {
                                'df': df_test.copy(),
                                'weighted_ami': weighted_ami,
                                '40_ami_coverage': ami_40_coverage,
                                'units_40_ami': num_40_units,
                                'units_80_ami': num_80_units,
                                'units_60_ami': total_units - num_40_units - num_80_units,
                                'compliant': is_compliant,
                                'score': weighted_ami if is_compliant else 0
                            }
                            results.append(result)
                    else:
                        # No 80% AMI units
                        df_test = df.copy()
                        df_test['AMI_Level'] = 0.60  # Default
                        
                        # Assign 40% AMI
                        for idx in units_40_combo:
                            df_test.iloc[idx, df_test.columns.get_loc('AMI_Level')] = 0.40
                        
                        # Calculate metrics
                        weighted_ami = self._calculate_weighted_ami(df_test)
                        ami_40_coverage = self._calculate_40_ami_coverage(df_test)
                        is_compliant = self._is_compliant(df_test, weighted_range, ami_40_range)
                        
                        # Store result
                        result = {
                            'df': df_test.copy(),
                            'weighted_ami': weighted_ami,
                            '40_ami_coverage': ami_40_coverage,
                            'units_40_ami': num_40_units,
                            'units_80_ami': 0,
                            'units_60_ami': total_units - num_40_units,
                            'compliant': is_compliant,
                            'score': weighted_ami if is_compliant else 0
                        }
                        results.append(result)
        
        return results
    
    def _assign_ami_strategy(self, df, strategy_type, weighted_range, ami_40_range):
        """Assign AMI levels using comprehensive strategy testing."""
        # Sort based on strategy
        if strategy_type == 'floor_optimized':
            df_sorted = df.sort_values(['Floor', 'SF'])
        elif strategy_type == 'size_optimized':
            df_sorted = df.sort_values(['SF', 'Floor'])
        else:  # optimal_revenue
            df_sorted = df.sort_values(['SF', 'Floor'])
        
        # Generate all possible combinations
        all_results = self._generate_all_combinations(df_sorted, weighted_range, ami_40_range)
        
        if not all_results:
            return None
        
        # Filter compliant results
        compliant_results = [r for r in all_results if r['compliant']]
        
        if compliant_results:
            # Return best compliant result (highest weighted AMI)
            best_result = max(compliant_results, key=lambda x: x['score'])
            return best_result['df']
        else:
            # Return best non-compliant result for analysis
            best_result = max(all_results, key=lambda x: x['weighted_ami'])
            return best_result['df']
    
    def _generate_detailed_recommendations(self, df, weighted_range, ami_40_range, all_attempts):
        """Generate detailed, specific recommendations."""
        recommendations = []
        
        # Analyze what we achieved vs what we need
        best_attempt = max(all_attempts, key=lambda x: x.get('weighted_ami', 0))
        
        current_weighted = best_attempt.get('weighted_ami', 0)
        current_40_coverage = best_attempt.get('40_ami_coverage', 0)
        
        # Specific gap analysis
        weighted_gap = weighted_range[0] - current_weighted
        ami_40_gap = 20.0 - current_40_coverage  # 20% minimum
        
        if ami_40_gap > 0:
            recommendations.append({
                'title': f'Need {ami_40_gap:.1f}% More 40% AMI Coverage',
                'description': f'Current: {current_40_coverage:.1f}%, Required: 20.0% minimum. Gap: {ami_40_gap:.1f}%',
                'expected_result': f'Add approximately {int(ami_40_gap * len(df) / 100)} more units at 40% AMI or increase unit sizes.'
            })
        
        if weighted_gap > 0:
            recommendations.append({
                'title': f'Need {weighted_gap:.1f}% Higher Weighted AMI',
                'description': f'Current: {current_weighted:.1f}%, Target: {weighted_range[0]:.1f}%. Gap: {weighted_gap:.1f}%',
                'expected_result': f'Replace some 60% AMI units with 80% AMI units, or increase unit sizes for higher AMI units.'
            })
        
        # Unit-specific recommendations based on building analysis
        total_sf = df['SF'].sum()
        target_40_sf = total_sf * 0.20  # 20% minimum
        
        # Calculate how many units needed
        avg_unit_sf = df['SF'].mean()
        units_needed_40 = int(target_40_sf / avg_unit_sf) + 1
        
        if units_needed_40 > len(df):
            recommendations.append({
                'title': 'Insufficient Units for Compliance',
                'description': f'Need {units_needed_40} units at 40% AMI, but only have {len(df)} total affordable units.',
                'expected_result': 'Add more affordable units to the building or modify existing unit sizes.'
            })
        else:
            # Specific unit size recommendations
            smallest_units = df.nsmallest(units_needed_40, 'SF')
            largest_units = df.nlargest(min(3, len(df) - units_needed_40), 'SF')
            
            recommendations.append({
                'title': 'Optimal Unit Assignment Strategy',
                'description': f'Assign the {units_needed_40} smallest units ({smallest_units["SF"].sum():.0f} SF total) to 40% AMI.',
                'expected_result': f'This would achieve {(smallest_units["SF"].sum()/total_sf)*100:.1f}% at 40% AMI.'
            })
            
            if len(largest_units) > 0:
                recommendations.append({
                    'title': 'Revenue Optimization Strategy',
                    'description': f'Assign the {len(largest_units)} largest units ({largest_units["SF"].sum():.0f} SF total) to 80% AMI.',
                    'expected_result': f'This would help maximize the weighted average AMI while maintaining compliance.'
                })
        
        return recommendations
    
    def process_optimization(self, df, mode='cascading', **kwargs):
        """Process comprehensive AMI optimization with detailed results."""
        try:
            df_clean = self._clean_data(df)
            
            results = {
                'success': True,
                'strategies': {},
                'recommendations': [],
                'mode': mode,
                'building_analysis': {
                    'total_units': len(df_clean),
                    'total_sf': int(df_clean['SF'].sum()),
                    'avg_unit_sf': int(df_clean['SF'].mean()),
                    'min_unit_sf': int(df_clean['SF'].min()),
                    'max_unit_sf': int(df_clean['SF'].max())
                }
            }
            
            # Define scenarios to test
            if mode == 'cascading':
                scenarios = [
                    ('Perfect Scenario', (59.5, 60.0), (20.0, 20.5)),
                    ('Excellent Scenario', (58.5, 59.5), (20.5, 21.5)),
                    ('Great Scenario', (57.0, 58.5), (21.5, 23.0)),
                    ('Conservative Fallback', (55.0, 57.0), (23.0, 25.0))
                ]
            elif mode == 'custom':
                weighted_range = kwargs.get('custom_weighted_range', (58, 60))
                ami_40_range = kwargs.get('custom_40_ami_range', (20, 22))
                scenarios = [('Custom Scenario', weighted_range, ami_40_range)]
            else:
                preset = self.scenario_presets.get(mode, self.scenario_presets['balanced'])
                scenarios = [(f'{mode.title()} Scenario', preset['weighted_ami_range'], preset['40_ami_range'])]
            
            # Test each scenario
            all_attempts = []
            
            for scenario_name, weighted_range, ami_40_range in scenarios:
                scenario_results = {}
                scenario_attempts = []
                
                strategies = {
                    'optimal_revenue': 'Optimal Revenue Strategy',
                    'floor_optimized': 'Floor Optimized Strategy',
                    'size_optimized': 'Size Optimized Strategy'
                }
                
                for strategy_key, strategy_name in strategies.items():
                    result_df = self._assign_ami_strategy(df_clean, strategy_key, weighted_range, ami_40_range)
                    
                    if result_df is not None:
                        weighted_ami = self._calculate_weighted_ami(result_df)
                        ami_40_coverage = self._calculate_40_ami_coverage(result_df)
                        is_compliant = self._is_compliant(result_df, weighted_range, ami_40_range)
                        
                        # Count units by AMI level
                        units_40 = len(result_df[result_df['AMI_Level'] == 0.40])
                        units_60 = len(result_df[result_df['AMI_Level'] == 0.60])
                        units_80 = len(result_df[result_df['AMI_Level'] == 0.80])
                        
                        strategy_result = {
                            'df_with_ami': result_df,
                            'weighted_ami': weighted_ami,
                            '40_ami_coverage': ami_40_coverage,
                            'units_40_ami': units_40,
                            'units_60_ami': units_60,
                            'units_80_ami': units_80,
                            'total_units': len(result_df),
                            'compliant': is_compliant,
                            'scenario': scenario_name,
                            'ami_distribution': {
                                '40%': units_40,
                                '60%': units_60,
                                '80%': units_80
                            }
                        }
                        
                        scenario_results[strategy_name] = strategy_result
                        scenario_attempts.append(strategy_result)
                
                all_attempts.extend(scenario_attempts)
                
                # Use first scenario with compliant strategies
                compliant_strategies = [s for s in scenario_results.values() if s.get('compliant', False)]
                if compliant_strategies and not results['strategies']:
                    results['strategies'] = scenario_results
                    results['scenario_used'] = scenario_name
                    break
                elif not results['strategies']:  # Keep first attempt
                    results['strategies'] = scenario_results
                    results['scenario_used'] = scenario_name
            
            # Generate detailed recommendations if needed
            if not any(s.get('compliant', False) for s in results['strategies'].values()):
                # Use the most restrictive scenario for recommendations
                target_scenario = scenarios[0] if scenarios else ('Default', (58, 60), (20, 22))
                results['recommendations'] = self._generate_detailed_recommendations(
                    df_clean, target_scenario[1], target_scenario[2], all_attempts
                )
            
            return results
            
        except Exception as e:
            return {
                'success': False,
                'error': str(e),
                'strategies': {},
                'recommendations': []
            }

# HTML Template (same as before but with enhanced results display)
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Universal Guaranteed AMI Calculator</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
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
        .header h1 { font-size: 2.5em; margin-bottom: 10px; font-weight: 300; }
        .header p { font-size: 1.1em; opacity: 0.9; }
        .content { padding: 40px; }
        .upload-section {
            background: #f8f9fa;
            border-radius: 15px;
            padding: 30px;
            margin-bottom: 30px;
            border: 2px dashed #dee2e6;
            text-align: center;
            transition: all 0.3s ease;
        }
        .upload-section:hover { border-color: #667eea; background: #f0f4ff; }
        .file-input { display: none; }
        .file-label {
            display: inline-block;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 30px;
            border-radius: 50px;
            cursor: pointer;
            font-size: 1.1em;
            transition: all 0.3s ease;
            margin-bottom: 15px;
        }
        .file-label:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.3);
        }
        .scenario-section { margin-bottom: 30px; }
        .scenario-title {
            font-size: 1.5em;
            color: #2c3e50;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #ecf0f1;
        }
        .scenario-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            margin-bottom: 20px;
        }
        .scenario-card {
            background: white;
            border: 2px solid #ecf0f1;
            border-radius: 15px;
            padding: 20px;
            transition: all 0.3s ease;
            cursor: pointer;
        }
        .scenario-card:hover {
            border-color: #667eea;
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(0,0,0,0.1);
        }
        .scenario-card.selected { border-color: #667eea; background: #f0f4ff; }
        .scenario-name { font-weight: 600; color: #2c3e50; margin-bottom: 10px; }
        .scenario-details { font-size: 0.9em; color: #7f8c8d; }
        .calculate-btn {
            background: linear-gradient(135deg, #27ae60 0%, #2ecc71 100%);
            color: white;
            border: none;
            padding: 15px 40px;
            border-radius: 50px;
            font-size: 1.2em;
            cursor: pointer;
            transition: all 0.3s ease;
            width: 100%;
            margin-top: 20px;
        }
        .calculate-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(39, 174, 96, 0.3);
        }
        .calculate-btn:disabled {
            background: #bdc3c7;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        .loading { display: none; text-align: center; padding: 40px; }
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #667eea;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
        }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .results { display: none; margin-top: 30px; }
        .building-analysis {
            background: #e8f5e8;
            border-radius: 15px;
            padding: 20px;
            margin-bottom: 20px;
        }
        .building-analysis h3 { color: #27ae60; margin-bottom: 15px; }
        .analysis-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
        }
        .analysis-item {
            text-align: center;
            background: white;
            padding: 15px;
            border-radius: 10px;
        }
        .analysis-value { font-size: 1.3em; font-weight: 600; color: #27ae60; }
        .analysis-label { font-size: 0.9em; color: #666; margin-top: 5px; }
        .result-card {
            background: white;
            border-radius: 15px;
            padding: 25px;
            margin-bottom: 20px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }
        .result-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
        }
        .result-title { font-size: 1.3em; font-weight: 600; color: #2c3e50; }
        .result-status {
            padding: 8px 16px;
            border-radius: 20px;
            font-weight: 600;
            font-size: 0.9em;
        }
        .status-success { background: #d4edda; color: #155724; }
        .status-failed { background: #f8d7da; color: #721c24; }
        .result-metrics {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
        }
        .metric {
            text-align: center;
            padding: 15px;
            background: #f8f9fa;
            border-radius: 10px;
        }
        .metric-value { font-size: 1.3em; font-weight: 600; color: #2c3e50; }
        .metric-label { font-size: 0.9em; color: #7f8c8d; margin-top: 5px; }
        .ami-distribution {
            background: #f0f8ff;
            border-radius: 10px;
            padding: 15px;
            margin-bottom: 15px;
        }
        .ami-distribution h4 { color: #2c3e50; margin-bottom: 10px; }
        .distribution-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 10px;
        }
        .distribution-item {
            text-align: center;
            background: white;
            padding: 10px;
            border-radius: 8px;
        }
        .download-btn {
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
            color: white;
            border: none;
            padding: 12px 25px;
            border-radius: 25px;
            cursor: pointer;
            transition: all 0.3s ease;
            text-decoration: none;
            display: inline-block;
        }
        .download-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 16px rgba(52, 152, 219, 0.3);
        }
        .recommendations {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 10px;
            padding: 20px;
            margin-top: 20px;
        }
        .recommendations h4 { color: #856404; margin-bottom: 15px; }
        .recommendation-item {
            background: white;
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 10px;
            border-left: 4px solid #f39c12;
        }
        @media (max-width: 768px) {
            .container { margin: 10px; border-radius: 15px; }
            .content { padding: 20px; }
            .scenario-grid { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Universal Guaranteed AMI Calculator</h1>
            <p>Advanced Multi-Tier Optimization with Universal Recommendations</p>
        </div>
        
        <div class="content">
            <div class="upload-section">
                <input type="file" id="fileInput" class="file-input" accept=".xlsx,.xls" />
                <label for="fileInput" class="file-label">📁 Choose Excel File</label>
                <p>Upload your building's unit schedule (Excel format)</p>
                <div id="fileName" style="margin-top: 10px; font-weight: 600; color: #27ae60;"></div>
            </div>
            
            <div class="scenario-section">
                <h2 class="scenario-title">🎯 Choose Optimization Scenario</h2>
                
                <div class="scenario-grid">
                    <div class="scenario-card" data-mode="cascading">
                        <div class="scenario-name">🚀 Cascading Optimization (Recommended)</div>
                        <div class="scenario-details">
                            Automatically tests Perfect → Excellent → Great scenarios<br>
                            Finds the best possible solution with universal guarantee
                        </div>
                    </div>
                </div>
                
                <h3 style="margin: 30px 0 15px 0; color: #2c3e50;">📋 Preset Scenarios</h3>
                <div class="scenario-grid">
                    <div class="scenario-card" data-mode="conservative">
                        <div class="scenario-name">🛡️ Conservative</div>
                        <div class="scenario-details">
                            Weighted AMI: 57-58%<br>
                            40% AMI Coverage: 22-23%<br>
                            Safe, compliant approach
                        </div>
                    </div>
                    
                    <div class="scenario-card" data-mode="balanced">
                        <div class="scenario-name">⚖️ Balanced</div>
                        <div class="scenario-details">
                            Weighted AMI: 58-60%<br>
                            40% AMI Coverage: 20-22%<br>
                            Optimal balance of revenue and compliance
                        </div>
                    </div>
                    
                    <div class="scenario-card" data-mode="aggressive">
                        <div class="scenario-name">💰 Aggressive Revenue</div>
                        <div class="scenario-details">
                            Weighted AMI: 59-60%<br>
                            40% AMI Coverage: 20-21%<br>
                            Maximum revenue optimization
                        </div>
                    </div>
                    
                    <div class="scenario-card" data-mode="ultra_precise">
                        <div class="scenario-name">🎯 Ultra-Precise</div>
                        <div class="scenario-details">
                            Weighted AMI: 59.5-60%<br>
                            40% AMI Coverage: 20-20.5%<br>
                            Extremely tight tolerances
                        </div>
                    </div>
                </div>
            </div>
            
            <button id="calculateBtn" class="calculate-btn" disabled>🚀 Calculate AMI Strategies</button>
            
            <div id="loading" class="loading">
                <div class="spinner"></div>
                <p>Analyzing building and optimizing strategies...</p>
            </div>
            
            <div id="results" class="results"></div>
        </div>
    </div>

    <script>
        let selectedMode = null;
        let uploadedFile = null;
        
        document.getElementById('fileInput').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (file) {
                uploadedFile = file;
                document.getElementById('fileName').textContent = `✅ ${file.name}`;
                updateCalculateButton();
            }
        });
        
        document.querySelectorAll('.scenario-card').forEach(card => {
            card.addEventListener('click', function() {
                document.querySelectorAll('.scenario-card').forEach(c => c.classList.remove('selected'));
                this.classList.add('selected');
                selectedMode = this.dataset.mode;
                updateCalculateButton();
            });
        });
        
        function updateCalculateButton() {
            const btn = document.getElementById('calculateBtn');
            if (uploadedFile && selectedMode) {
                btn.disabled = false;
                btn.textContent = '🚀 Calculate AMI Strategies';
            } else {
                btn.disabled = true;
                btn.textContent = uploadedFile ? 'Select a scenario above' : 'Upload a file first';
            }
        }
        
        document.getElementById('calculateBtn').addEventListener('click', function() {
            if (!uploadedFile || !selectedMode) return;
            
            document.getElementById('loading').style.display = 'block';
            document.getElementById('results').style.display = 'none';
            
            const formData = new FormData();
            formData.append('file', uploadedFile);
            formData.append('mode', selectedMode);
            
            fetch('/calculate', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                document.getElementById('loading').style.display = 'none';
                displayResults(data);
            })
            .catch(error => {
                document.getElementById('loading').style.display = 'none';
                alert('Error: ' + error.message);
            });
        });
        
        function displayResults(data) {
            const resultsDiv = document.getElementById('results');
            let html = '';
            
            if (data.success) {
                html += `<h2 style="color: #27ae60; margin-bottom: 20px;">✅ Optimization Results</h2>`;
                
                // Building Analysis
                if (data.building_analysis) {
                    const analysis = data.building_analysis;
                    html += `
                        <div class="building-analysis">
                            <h3>🏢 Building Analysis</h3>
                            <div class="analysis-grid">
                                <div class="analysis-item">
                                    <div class="analysis-value">${analysis.total_units}</div>
                                    <div class="analysis-label">Total Units</div>
                                </div>
                                <div class="analysis-item">
                                    <div class="analysis-value">${analysis.total_sf.toLocaleString()}</div>
                                    <div class="analysis-label">Total SF</div>
                                </div>
                                <div class="analysis-item">
                                    <div class="analysis-value">${analysis.avg_unit_sf}</div>
                                    <div class="analysis-label">Avg Unit SF</div>
                                </div>
                                <div class="analysis-item">
                                    <div class="analysis-value">${analysis.min_unit_sf}</div>
                                    <div class="analysis-label">Min Unit SF</div>
                                </div>
                                <div class="analysis-item">
                                    <div class="analysis-value">${analysis.max_unit_sf}</div>
                                    <div class="analysis-label">Max Unit SF</div>
                                </div>
                            </div>
                        </div>
                    `;
                }
                
                // Display each strategy
                let strategyIndex = 0;
                for (const [strategyName, strategy] of Object.entries(data.strategies)) {
                    if (strategy && strategy.weighted_ami !== undefined) {
                        const statusClass = strategy.compliant ? 'status-success' : 'status-failed';
                        const statusText = strategy.compliant ? 'COMPLIANT' : 'FAILED';
                        
                        html += `
                            <div class="result-card">
                                <div class="result-header">
                                    <div class="result-title">${strategyName}</div>
                                    <div class="result-status ${statusClass}">${statusText}</div>
                                </div>
                                
                                <div class="ami-distribution">
                                    <h4>AMI Level Distribution</h4>
                                    <div class="distribution-grid">
                                        <div class="distribution-item">
                                            <div style="font-weight: 600; color: #e74c3c;">${strategy.units_40_ami || 0}</div>
                                            <div style="font-size: 0.8em;">40% AMI</div>
                                        </div>
                                        <div class="distribution-item">
                                            <div style="font-weight: 600; color: #f39c12;">${strategy.units_60_ami || 0}</div>
                                            <div style="font-size: 0.8em;">60% AMI</div>
                                        </div>
                                        <div class="distribution-item">
                                            <div style="font-weight: 600; color: #27ae60;">${strategy.units_80_ami || 0}</div>
                                            <div style="font-size: 0.8em;">80% AMI</div>
                                        </div>
                                    </div>
                                </div>
                                
                                <div class="result-metrics">
                                    <div class="metric">
                                        <div class="metric-value">${strategy.weighted_ami.toFixed(2)}%</div>
                                        <div class="metric-label">Weighted AMI</div>
                                    </div>
                                    <div class="metric">
                                        <div class="metric-value">${strategy['40_ami_coverage'].toFixed(1)}%</div>
                                        <div class="metric-label">40% AMI Coverage</div>
                                    </div>
                                    <div class="metric">
                                        <div class="metric-value">${strategy.units_40_ami}</div>
                                        <div class="metric-label">Units at 40% AMI</div>
                                    </div>
                                    <div class="metric">
                                        <div class="metric-value">${strategy.total_units}</div>
                                        <div class="metric-label">Total Units</div>
                                    </div>
                                </div>
                                
                                ${strategy.compliant ? `
                                    <a href="/download/${strategyIndex}" class="download-btn">
                                        📥 Download Excel File
                                    </a>
                                ` : ''}
                            </div>
                        `;
                        strategyIndex++;
                    }
                }
                
                // Show recommendations
                if (data.recommendations && data.recommendations.length > 0) {
                    html += `
                        <div class="recommendations">
                            <h4>💡 Detailed Recommendations</h4>
                            <p>Specific guidance to achieve your targets:</p>
                    `;
                    
                    data.recommendations.forEach(rec => {
                        html += `
                            <div class="recommendation-item">
                                <strong>${rec.title}</strong><br>
                                ${rec.description}<br>
                                <em>Expected Result: ${rec.expected_result}</em>
                            </div>
                        `;
                    });
                    
                    html += '</div>';
                }
                
            } else {
                html += `
                    <div class="result-card">
                        <div class="result-header">
                            <div class="result-title">❌ Analysis Failed</div>
                            <div class="result-status status-failed">ERROR</div>
                        </div>
                        <p>${data.error}</p>
                    </div>
                `;
            }
            
            resultsDiv.innerHTML = html;
            resultsDiv.style.display = 'block';
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/calculate', methods=['POST'])
def calculate():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file uploaded'})
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'})
        
        mode = request.form.get('mode', 'cascading')
        
        # Read Excel file
        df = pd.read_excel(file)
        
        # Initialize calculator
        calculator = EnhancedAMICalculator()
        
        # Run optimization
        result = calculator.process_optimization(df, mode=mode)
        
        # Store results for download
        app.config['LAST_RESULTS'] = result
        
        return jsonify(result)
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/download/<int:strategy_index>')
def download_strategy(strategy_index):
    try:
        results = app.config.get('LAST_RESULTS')
        if not results or not results.get('success'):
            return "No results available", 404
        
        strategy_names = list(results['strategies'].keys())
        if strategy_index >= len(strategy_names):
            return "Strategy not found", 404
        
        strategy_name = strategy_names[strategy_index]
        strategy_data = results['strategies'][strategy_name]
        
        if not strategy_data or 'df_with_ami' not in strategy_data:
            return "Strategy data not available", 404
        
        # Create Excel file
        output = io.BytesIO()
        df_result = strategy_data['df_with_ami']
        df_result.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        
        # Generate filename
        safe_name = strategy_name.replace(' ', '_').replace('/', '_')
        filename = f"{safe_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        return f"Error generating download: {str(e)}", 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=True)

