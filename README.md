# 📦 Inventory Analysis Suite

A professional Python utility for automated inventory classification and analysis. This tool helps businesses optimize their supply chain by providing two core analytical methodologies: ABC and XYZ analysis.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Pandas](https://img.shields.io/badge/Pandas-2.2.0-orange.svg)
![NumPy](https://img.shields.io/badge/NumPy-1.26.3-blue.svg)

## 🎯 Features

### 1. **ABC Analysis** (Value Importance)
Classifies items based on their cumulative monetary value contribution using the Pareto Principle (80/20 rule).

- **Group A**: Top ~80% of total value (Critical items)
- **Group B**: Next ~15% of total value (Medium importance)
- **Group C**: Remaining ~5% of total value (Low priority)

**Required Input:**
- `Stock.xlsx` (Date, SKU, Stock quantity)
- `COGS.xlsx` (SKU, Cost of Goods Sold)

### 2. **XYZ Analysis** (Demand Stability)
Classifies items by the stability of their stock levels or demand using the Coefficient of Variation (CV).

- **Group X**: Stable demand/stock (CV ≤ 33rd percentile)
- **Group Y**: Moderate variability (33rd < CV ≤ 66th percentile)
- **Group Z**: High volatility (CV > 66th percentile)

**Required Input:**
- `Stock.xlsx` (Date, SKU, Stock/Value quantity)

## 🚀 Installation & Setup

### Prerequisites
- Python 3.8 or higher
- pip (Python package manager)

### Setup Instructions

1. **Clone the repository:**
```bash
git clone <your-repository-url>
cd ABC_PROJECT
Install dependencies:

Bash

pip install pandas numpy openpyxl xlsxwriter
Prepare Data Folders: The script automatically creates these, but ensure your files are placed correctly:

Put input files in: data/input/

Results will appear in: data/output/

📂 Project Structure
ABC_PROJECT/
├── abc_analyzer.py      # Main logic for ABC classification
├── xyz_analyzer.py      # Main logic for XYZ classification
├── .gitignore           # Prevents uploading private .xlsx data
├── README.md            # Project documentation
└── data/
    ├── input/           # Place Stock.xlsx and COGS.xlsx here
    └── output/          # Generated Excel reports
📋 Input File Formats
Required Columns
Stock File (Stock.xlsx):

Date       | SKU      | Stock
2024-01-15 | SKU001   | 100
2024-01-15 | SKU002   | 250
COGS File (COGS.xlsx):

SKU      | COGS
SKU001   | 25.50
SKU002   | 15.75
📈 Analysis Workflow
Data Loading

Scripts read Excel files from data/input/.

Automatic validation of required columns and data types.

Data Transformation

Dates are converted to YYYY-MM periods.

Values are aggregated per SKU and Period.

Classification

ABC: Calculates cumulative percentages and assigns groups.

XYZ: Calculates standard deviation and mean to find the CV.

Output

Saves final results to data/output/ as .xlsx files.

Includes original data + new classification columns (ABC_Group or XYZ).

🔧 Technical Details
Calculation Logic
ABC Classification:

Python

# Cumulative Value % = Cumulative Sum / Total Sum
# Thresholds: <= 0.8 (A), <= 0.95 (B), else (C)
XYZ Classification:

Python

# CV (Coefficient of Variation) = Standard Deviation / Mean
# Automatically determines thresholds using 33% and 66% quantiles
🛡️ Security & Privacy
⚠️ Data Safety: This project is configured with a .gitignore file to ensure that your business spreadsheets (.xlsx, .csv) are never uploaded to GitHub.

Keep your data/ folder local!

🤝 Contributing
Contributions are welcome!

Fork the repository

Create a feature branch

Submit a Pull Request

📄 License
This project is licensed under the MIT License.

👨‍💻 Author
Created with ❤️ for better inventory management.

Happy Analyzing! 📊