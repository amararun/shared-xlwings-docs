# Excel Data Analysis & ML Workflows with xlwings Lite

This repository demonstrates how to run full Python workflows directly inside Excel using xlwings Lite. It includes examples of data transformations, automations, statistical analysis, machine learning models, and external system integrations.

## Overview

xlwings Lite is a powerful addition to the data scientist's toolkit that allows you to run Python code directly within Excel. This repository contains practical examples showing how to leverage this capability for various data analysis and machine learning tasks.

## Key Features Demonstrated

- 🔄 **Data Transformations**: Complex Excel operations using Python
- 🤖 **Automations**: Excel workflow automation without VBA
- 📊 **Statistical Analysis**: Advanced statistical operations
- 🧠 **Machine Learning**: XGBoost model implementation
- 📈 **Visualizations**: ROC curves, gain charts, and more
- 🔍 **Feature Engineering**: Data preparation and analysis
- 🔌 **External Connections**: API calls and remote database integration
- 📱 **Web Integration**: Connect to any system with API access

## Example Workflows

### 1. Monte Carlo Simulation & Distribution Analysis
- Generate random samples from multiple probability distributions (Uniform, Triangular, Normal, Log-Normal)
- Interactive input parameters with P10, P50, P90 values
- Dynamic visualization of distributions with histograms
- Statistical analysis including mean, standard deviation, and percentiles
- Cumulative distribution plots for risk analysis
- Automated dashboard generation with multiple charts
- Real-time updates of statistical measures
- Support for oil and gas reserve calculations
- Customizable input parameters and distribution types

### 2. XGBoost Response Model
- Loads data from Excel tables
- Performs feature engineering
- Trains XGBoost model
- Generates decile tables
- Creates ROC and gain curves
- Exports results back to Excel

### 3. Credit Card Segment Analysis
- Table formatting and data preparation
- Weighted scoring implementation
- Multicollinearity analysis
- Results visualization

### 4. Remote Database Integration
- Connect to remote PostgreSQL database
- Explore tables and execute custom SQL queries
- Pull and analyze data directly from Excel
- Generate descriptive statistics and visualizations
- Build and evaluate ML models on remote data

## Advanced Features

### API Integration
- Connect to external systems via API
- Remote database connections through web layer
- FastAPI server integration for secure database access
- Support for various API endpoints (databases, LLMs)

### LLM Integration
- API calls to OpenAI, Gemini, and Claude Sonnet
- Structured output using JSON schemas
- Automation capabilities with LLM responses

### Excel Object Manipulation
- Dynamic sheet creation and management
- Custom table formatting
- Automated report generation
- Interactive chart creation

## Getting Started

1. **Installation**
   ```bash
   # Install xlwings Add-in in Excel
   
   ```

2. **Usage**
   - Open the Excel files in this repository
   - Use the xlwings interface to run the scripts
   - View results in new sheets within the workbook


## Resources

### Documentation & Guides
- [Monte Carlo Simulation & Dashboards with xlwings from Franco Silvia](https://github.com/FrancoSivila/OGIP_Py/blob/3adb423a17c2170d2d976bfb3b5e3975c36c0ad3/14%20OGIPpy%20xlwings%20(26%20mar%2025).xlsm)
- [xlwings Documentation](https://www.xlwings.org/)
- [xlwings Lite Documentation](https://lite.xlwings.org/)
- [Running Python in Excel - Part 1](https://lnkd.in/gjdqpqXK)
- [API Calls, Database Connect, EDA, Machine Learning - Part 2](https://www.linkedin.com/feed/update/urn:li:activity:7308863086529040385/)
- [Post & walkthrough video by Felix Zumstein (creator of xlwings)](https://lnkd.in/geKzDK45)


