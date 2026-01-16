# 📦 Inventory Analysis Suite

A professional Python utility for automated inventory classification and analysis.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Pandas](https://img.shields.io/badge/Pandas-2.2.0-orange.svg)

## 🎯 Features

### 1. ABC Analysis
Classifies items based on their cumulative monetary value contribution (80/20 rule).
- **Group A**: Top ~80% of value
- **Group B**: Next ~15% of value
- **Group C**: Remaining ~5% of value

### 2. XYZ Analysis
Classifies items by the stability of their demand using the Coefficient of Variation (CV).
- **Group X**: Stable demand
- **Group Y**: Moderate variability
- **Group Z**: High volatility

## 🚀 Installation & Setup

1. **Clone the repository:**
```bash
git clone <your-repository-url>
cd ABC_project
Install dependencies:

Bash

pip install pandas numpy openpyxl xlsxwriter
Prepare Data Folders: Put input files in: data/input/ Results will appear in: data/output/

📂 Project Structure
abc_analyzer.py: ABC classification logic.

xyz_analyzer.py: XYZ classification logic.

.gitignore: Prevents uploading private data.