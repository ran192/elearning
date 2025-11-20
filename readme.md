# Elearning Data Analysis Tool

## Project Overview
This tool is for homework purposes. It provides a GUI-based Python application for analyzing and generating reports from Excel, Word, and PowerPoint files. Main functionalities include:

1. **Data Transformation**
   - Reliability Analysis (Cronbach’s Alpha) for Excel files
   - Automatically detects column prefixes (E, CE, PP, FWS, R)
   - Calculates Cronbach’s Alpha and item-deleted effects
   - Outputs Excel, Word, or PowerPoint files in the same format as input with `_converted` suffix
   - Users can optionally customize the output filename

2. **Charts & Report Generation**
   - Supports single or multiple Excel file input
   - Automatically generates charts: GDP, Birth Rate, Death Rate, Gender Ratio, Gini Coefficient
   - Generates PDF report per country and a global overview report
   - Charts and reports can display Chinese characters

## Project Structure

### Install required packages
```bash
pip install -r requirements.txt
```
## Usage

### Run GUI
```bash
python main.py
```
