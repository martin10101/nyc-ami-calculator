# NYC AMI Calculator

A comprehensive tool for optimizing Affordable Housing AMI (Area Median Income) distribution in NYC development projects.

## Features

- **Client-Respectful Approach**: Works with pre-selected affordable units
- **Maximum 60% Weighted AMI**: Complies with specific NYC program requirements
- **Minimum 20% at 40% AMI**: Ensures deep affordability requirements
- **Smart Assignment Logic**: Lower floors + smaller units get 40% AMI first
- **Multiple Strategies**: Provides 3 optimized AMI distribution options
- **Comprehensive Compliance**: Validates against NYC regulations
- **Professional Reports**: Generates detailed Excel files and web interface

## Requirements

- Python 3.8+
- Flask
- pandas
- numpy
- openpyxl

## Installation

1. Clone this repository
2. Install dependencies: `pip install -r requirements.txt`
3. Run the application: `python client_specific_ami_calculator.py`
4. Open your browser to `http://localhost:5000`

## Usage

1. Upload an Excel file with your building data and pre-selected affordable units
2. The system will generate 3 optimized AMI distribution strategies
3. Download Excel files with detailed unit assignments
4. View compliance analysis and recommendations

## Deployment

This application is configured for deployment on Railway.app with automatic builds and deployments.

## License

Proprietary - All rights reserved

