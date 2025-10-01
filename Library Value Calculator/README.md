# Library Value Use Calculator

This project supports the Manatee County Public Library's Value Use Calculator by allowing library statisticians to easily update calculator settings while enabling IT support to deploy updated code seamlessly.

## Overview

The Library Value Use Calculator helps library patrons understand the value of the services they use. The system consists of an Excel-based configuration file and a Python script that generates production-ready web code.

## Features

- **Automated Code Generation**: Generates complete HTML, CSS, and JavaScript from Excel configurations
- **Dynamic Fiscal Year Support**: Automatically includes buttons for fiscal years with complete data
- **Responsive Design**: Mobile-friendly interface with professional styling
- **Interactive Elements**: Smooth animations, input formatting, and real-time calculations
- **Modular Architecture**: Separates configuration from code generation for easy maintenance

## File Structure

```
├── Library Programming Matrix Map/
│     ├── Refresh Dashboard Dataset.py
│     ├── Calculator Settings.xlsx          # Configuration file for calculator data
│     ├── library_calculator_html_builder.py  # Main code generation script
│     ├── style_and_script.html           # Generated head section code
│     ├── content_box_code.html           # Generated body content code
│     └── prototype.html                  # Complete test file
```

## Configuration Files

### Calculator Settings.xlsx
Contains two main worksheets:
- **Values Per Service**: Service names, values, and explanations
- **Fiscal Year Values**: Historical usage data by fiscal year

## Core Functions

### `generate_templates()`
Main orchestration function that:
- Loads Excel configuration data
- Processes service information and fiscal year data
- Generates all output files

### Data Processing Functions
- **Excel Data Extraction**: Reads service values and fiscal year statistics
- **Dictionary Building**: Creates structured data objects for JavaScript generation
- **Column Letter Conversion**: `get_column_letter()` handles Excel column indexing

### Code Generation Functions

#### `get_style()`
Generates comprehensive CSS including:
- Responsive table styling
- Manatee County color scheme (#415364 primary, #d15e14 accent)
- Mobile-optimized layouts
- Interactive button states

#### `get_script()`
Creates JavaScript functionality:
- **`formatCurrency()`**: Formats monetary values with proper locale
- **`formatNumberWithCommas()`**: Adds thousand separators to large numbers
- **`calculate()`**: Real-time value calculations as users input data
- **`loadFiscalYear()`**: Populates fields with historical data
- **`animateValue()`**: Smooth number transitions with easing effects
- **`toggleExplanation()`**: Shows/hides service value explanations
- **`initializeCalculator()`**: Sets up event listeners and initial state

#### Template Generation Functions
- **`get_table()`**: Builds HTML table structure with dynamic service rows
- **`get_buttons()`**: Creates fiscal year and control buttons
- **`generate_style_and_script()`**: Outputs head section code
- **`generate_content_box_code()`**: Outputs body content code
- **`generate_prototype()`**: Creates complete test file

## Key Technical Features

### Smart Data Validation
- Automatically excludes fiscal years with incomplete data
- Validates service names across worksheets
- Handles missing or zero values gracefully

### User Experience Enhancements
- **Input Formatting**: Automatic comma insertion and number parsing
- **Focus Behavior**: Removes formatting during editing for easier input
- **Animation System**: Smooth value transitions when loading fiscal year data
- **Responsive Design**: Adapts to mobile and desktop screens

### Code Quality
- Well-formatted, readable output code
- Comprehensive error handling
- Efficient DOM manipulation
- Modern JavaScript practices

## Usage Workflow

### For Library Statisticians
1. Update service data in `Calculator Settings.xlsx`
2. Run `library_calculator_html_builder.py`
3. Test functionality using `prototype.html`

### For IT/Web Support
1. Execute the Python script to generate fresh code
2. Copy `style_and_script.html` content to website head section
3. Copy `content_box_code.html` content to calculator content area
4. Publish updated page

## Technical Specifications

- **Backend**: Python with openpyxl for Excel processing
- **Frontend**: Vanilla JavaScript, CSS3, HTML5
- **Styling**: Custom CSS with mobile-first responsive design
- **Data Format**: Excel (.xlsx) configuration files
- **Output**: Production-ready web components

## Benefits

- **Separation of Concerns**: Non-technical staff can update data without touching code
- **Automated Deployment**: Reduces manual coding errors and deployment time
- **Consistent Formatting**: Ensures professional appearance across updates
- **Scalable Architecture**: Easy to add new services or modify calculations
- **Quality Assurance**: Built-in validation and testing capabilities

## Browser Compatibility

- Modern browsers with ES6+ support
- Mobile-responsive design
- Graceful fallbacks for older browsers

---

*This project demonstrates automated web development workflows, Excel-to-web data integration, and responsive JavaScript application development.*