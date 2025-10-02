# Library Data Projects

This repository contains multiple projects developed to support Manatee County Public Library System data reporting, visualization, and patron engagement. Each project is organized into its own folder with documentation and code.

## 📂 Projects
### 1. Library Data Dashboard
https://app.powerbigov.us/view?r=eyJrIjoiMmU2MmVhNzEtY2Y4Yi00NWUzLTg2NTgtNmFiYjg4MDU3MmVkIiwidCI6ImNiZjE4NTg3LTc0MjItNDBmMi1hOGYyLWVhYTNhNGVhNDI0MCJ9&pageName=36bcde6e6c9f7f8b1db1

A Python-based ETL automation that consolidates messy monthly Excel reports into a clean, standardized dataset (MasterDataset.xlsx). The dataset powers Power BI dashboards (internal and public-facing) for system-wide performance insights.

* **Tools & Tech:** Python (`openpyxl`), Power BI, Excel
* **Key Features:**

  * Automated ETL for multi-branch, multi-year data

  * Consistent schema for dashboard integration

  * Error handling and safe data validation

### 2. Library Programming Matrix Map

A system that processes LibCal event exports into structured datasets, enabling a Power BI Matrix Map for programming analysis, cost-benefit evaluation, and strategic planning.

* **Tools & Tech:** Python, Excel, Power BI, LibCal

* **Key Features:**

  * Automated CSV processing and Excel dataset generation

  * Configurable cost/impact scoring (`MatrixMapSettings.xlsx`)

  * Intelligent program classification (Story Time, Tech Support, Book Clubs, etc.)

  * Staff time and attendance tracking

### 3. Library Value Use Calculator

A web-based calculator that helps patrons estimate the monetary value of their library usage. The system uses an Excel-driven configuration file to generate production-ready HTML, CSS, and JavaScript.

* **Tools & Tech:** Python (`openpyxl`), HTML5, CSS3, JavaScript

* **Key Features:**

  * Automated code generation from Excel settings

  * Responsive design with mobile optimization

  * Real-time calculations with smooth animations

  * Dynamic fiscal year data support

## 🚀 Why These Projects Matter

* **Data Accessibility:** Transforms raw reports into clean datasets for analysis.

* **Strategic Insights:** Supports decision-making with detailed programming and event impact analysis.

* **Patron Engagement:** Provides tools like the Value Use Calculator to communicate the library’s value.

* **Sustainable Workflows:** Non-technical staff can update data in Excel without editing code.

## 🛠️ Common Tools & Practices

* **Python 3.9+** with `openpyxl` for Excel processing

* **Power BI** for interactive dashboards

* **Excel (.xlsx)** as configuration and data source format

* **Modular, maintainable code** with strong error handling and reusable functions
