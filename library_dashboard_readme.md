# Manatee County Library Data Dashboard

A comprehensive data analytics solution that transforms manual library statistics reporting into an automated, interactive Power BI dashboard system.

## 🏛️ Project Overview

This project was developed during a summer internship at **Manatee County Library** under the supervision of **Tammy Parrott**. The initiative replaced time-consuming manual data processes with a centralized, dynamic, and user-friendly dashboard that improves how the library system views and interacts with performance data.

### The Problem
Previously, library staff relied on:
- Manually searching through folders of Excel files
- Entering data into static templates
- Time-intensive processes to assess monthly and annual performance
- Limited ability to quickly reveal data insights

### The Solution
A modern Data Dashboard that provides:
- **Centralized Performance Overview** - All key metrics in one place
- **Interactive Filtering and User Guidance** - Easy data exploration
- **Automated Data Processing** - Python scripts handle data cleaning and preparation
- **Public Transparency** - Public-facing version planned for the library website

## 📊 Dashboard Features

The dashboard consists of **six informative tabs**:

1. **📈 Library by the Numbers** - Key performance data by branch
2. **📚 Collection Data** - Circulation, inter-library loans, digital/physical collection usage
3. **👥 About Our Patrons** - Registration patterns and visitor analytics
4. **🏆 Top Performers** - Highlighting high-performing branches
5. **📖 Library Material Usage** - Enhanced version of previous internal dashboard
6. **📖 Value Calculator** - Estimated Monetary Value of Library Services

## 🛠️ Technical Architecture

### Technologies Used
- **Python 3.x** - Data processing and automation
- **Power BI** - Interactive dashboard and visualization
- **Microsoft Excel** - Source data format
- **openpyxl** - Python library for Excel file manipulation

### Key Components

#### 1. Data Processing Pipeline (`Refresh Dashboard Dataset.py`)
- **Automated Data Extraction** - Processes multiple Excel files across fiscal years
- **Data Cleaning & Standardization** - Handles missing values, type conversion, and formatting
- **Multi-source Integration** - Combines data from:
  - Branch statistics files
  - Inter-Library Loan (ILL) data
  - Digital information reports
  - Technology statistics
  - Computer and study room usage

#### 2. Template Management System (`Create New Branch.py`)
- **Dynamic Branch Creation** - Automatically generates new branch templates
- **Consistency Maintenance** - Ensures uniform data structure across all branches
- **User-Friendly Interface** - Simple command-line prompts for easy operation

### Data Processing Features

#### Smart File Detection
```python
def detect_branch_files(folder_path):
    """Detect all Excel files ending with 'Branch.xlsx' and create branch legend data."""
```
- Automatically discovers branch files across multiple fiscal years
- Creates dynamic mapping for consistent data processing

#### Robust Error Handling
```python
def safe_append_library_row(worksheet, row, location, month_name, year, data_type, skip_columns=4):
    """Safely append a row to worksheet with error handling."""
```
- Comprehensive error logging for data quality assurance
- Graceful handling of missing or corrupted data

#### Multi-Year Processing
- Processes data across multiple fiscal years (October-September cycles)
- Maintains data integrity across different time periods
- Handles year transitions automatically

## 📁 Project Structure

```
Statistics Folder/
└── October 2023 - September 2024/
    ├── Braden River Branch.xlsx
    ├── Central Library Branch.xlsx
    ├── Digital Information Branch.xlsx
    └── ...
└── October 2024 - September 2025/
    ├── Braden River Branch.xlsx
    ├── Central Library Branch.xlsx
    ├── Digital Information Branch.xlsx
    └── ...
└── Template/
    ├── Braden River Branch.xlsx
    ├── Central Library Branch.xlsx
    ├── Digital Information Branch.xlsx
    └── ...
└── Library Data Dashboard/               # Source data directories
    ├── Create New Branch.py              # Branch template creation tool
    ├── MasterDataset.xlsx                # Generated consolidated dataset
    ├── Refresh Dashboard Dataset.py      # Main data processing script  
    ├── Internal Library Dashboard.pbix
    └── Public Library Dashboard.pbix
```

## 🚀 Getting Started

### Prerequisites
- Python 3.x
- Required Python packages:
  ```bash
  pip install openpyxl
  ```
- Microsoft Power BI Desktop
- Access to Manatee County Library statistics files


### Usage

#### Creating New Branch Templates
```bash
python "Create New Branch.py"
```
- Follow prompts to create new branch template
- Automatically populates standard cells with branch name

#### Refreshing Dashboard Data
```bash
python "Refresh Dashboard Dataset.py"
```
- Processes all fiscal year folders
- Generates `MasterDataset.xlsx` for Power BI import
- Provides detailed processing logs

#### Power BI Integration
1. Open Power BI Desktop
2. Import `MasterDataset.xlsx`
3. Refresh data connections
4. Dashboard automatically updates with new data

## 📈 Data Sources Processed

The system handles the following Excel file types:

| File Type | Content | Processing Method |
|-----------|---------|-------------------|
| `[Branch] Branch.xlsx` | Individual branch statistics | Dynamic branch detection |
| `ILL.xlsx` | Inter-Library Loan data | System-wide aggregation |
| `Digital Information.xlsx` | Digital collection usage | Monthly processing |
| `Tech Statistics.xlsx` | Technology usage metrics | Branch-level breakdown |
| `Summary Usage Report.xlsx` | Computer & study room usage | Location-based analysis |

## 🔧 Configuration

### Customizable Settings

The script includes easily modifiable configuration sections:

#### Month Mapping
```python
MONTHS = ["October", "November", "December", "January", "February", "March", 
          "April", "May", "June", "July", "August", "September"]
```

#### Cell Mappings
```python
GENERAL_STATISTICS_CELLS = {
    'Total Patrons': 'F8',
    'Volunteers Hours': 'F10',
    # ... additional mappings
}
```

#### File Expectations
```python
EXPECTED_FILES = ['ILL.xlsx', 'Digital Information.xlsx', 'Tech Statistics.xlsx', 'Summary Usage Report.xlsx']
```

## 📊 Sample Output

The processed dataset includes worksheets for:
- **General Statistics** - Branch-level operational data
- **Programming** - Event and program attendance
- **Digital Information** - Digital resource usage
- **ILL** - Inter-library loan transactions
- **Tech Statistics** - Technology platform usage
- **Computer & Study Room Usage** - Facility utilization
- **Branch Legend** - Location reference data

## 🌟 Key Achievements

- **Eliminated Manual Processing** - Reduced hours of manual work to minutes of automation
- **Improved Data Quality** - Consistent formatting and error handling
- **Enhanced Accessibility** - User-friendly dashboard interface
- **Scalable Solution** - Easily handles additional branches and time periods
- **Public Transparency** - Supports community engagement through data visibility

## 🔮 Future Enhancements

- **Real-time Data Integration** - Direct database connections
- **Advanced Analytics** - Predictive modeling and trend analysis
- **Mobile Optimization** - Responsive dashboard design
- **Automated Scheduling** - Scheduled data refreshes
- **Additional Visualizations** - Enhanced charts and interactive elements

## 🙏 Acknowledgments

- **Tammy Parrott** - Project Supervisor, Manatee County Library
- **Manatee County Library System** - Data and resources
- **Library Staff** - Feedback and testing support

*This project demonstrates the power of combining traditional library science with modern data analytics to improve public service delivery and operational efficiency.*