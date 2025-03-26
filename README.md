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

### 1. XGBoost Response Model
- Loads data from Excel tables
- Performs feature engineering
- Trains XGBoost model
- Generates decile tables
- Creates ROC and gain curves
- Exports results back to Excel

### 2. Credit Card Segment Analysis
- Table formatting and data preparation
- Weighted scoring implementation
- Multicollinearity analysis
- Results visualization

### 3. Remote Database Integration
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
- Secure credential management using environment variables
- API key rotation and access control

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
   # Install required Python packages
   pip install -r requirements.txt
   ```

2. **Usage**
   - Open the Excel files in this repository
   - Use the xlwings interface to run the scripts
   - View results in new sheets within the workbook

## Best Practices

- Format Excel ranges as tables for optimal results
- Use `.options(index=False)` when writing back to Excel tables
- Leverage the console and `print()` for debugging
- Share code samples with AI coders for efficient development
- Use FastAPI server for secure remote database connections

## When to Use xlwings Lite

### Great For:
- Excel automation (Python over VBA)
- Complex Excel transformations
- Quick EDA and validation
- In-sheet diagnostics
- Basic ML model implementation
- External system integration
- API-based workflows
- Remote database analysis

### When to Stick with Jupyter/Colab:
- Complex data warehouse operations
- Extensive data cleaning workflows
- Multiple iteration cycles
- Large-scale ML model development

## Requirements

- Python 3.x
- xlwings
- pandas
- numpy
- scikit-learn
- xgboost
- matplotlib
- seaborn
- fastapi (for remote database connections)
- requests (for API calls)

## Resources

### Documentation & Guides
- [xlwings Documentation](https://www.xlwings.org/)
- [xlwings Lite Documentation](https://lite.xlwings.org/)
- [Running Python in Excel - Part 1](https://lnkd.in/gjdqpqXK)
- [Build with xlwings Lite - Part 2](https://www.linkedin.com/feed/update/urn:li:activity:7308863086529040385/)

### Example Files
- [Example Excel Files](*.xlsx)
- [Scripts](SCRIPTS.MD)

## Contributing

Feel free to submit issues and enhancement requests!

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Security Best Practices

### Credential Management
- Never commit sensitive credentials to the repository
- Use environment variables for API keys and passwords
- Store credentials in `.env` files (listed in .gitignore)
- Rotate credentials regularly
- Use secure credential storage services when possible

### API Security
- Use HTTPS for all API communications
- Implement proper authentication and authorization
- Use API key rotation
- Follow the principle of least privilege
- Monitor API usage and access patterns

### Database Security
- Use connection pooling
- Implement proper access controls
- Encrypt sensitive data in transit
- Use parameterized queries to prevent SQL injection
- Regular security audits and updates

### General Security
- Keep dependencies updated
- Regular security scanning
- Follow security best practices for your specific use case
- Monitor for suspicious activities
- Implement proper error handling

