# Library Programming Matrix Map System Documentation

## Overview

The Library Programming Matrix Map System is a comprehensive data processing and analysis tool designed to transform LibCal event data into strategic programming insights. The system processes library event information and creates structured datasets that feed into a Power BI matrix map for cost-benefit analysis and strategic decision-making.

## System Architecture

```
LibCal Events Export → Refresh Matrix Dataset.py → Excel Datasets → Power BI Matrix Map
```

### Components
- **Data Source**: LibCal event management system
- **Processing Engine**: `Refresh Matrix Dataset.py`
- **Output Files**: `MatrixMapDataset.xlsx` and `MatrixMapSettings.xlsx`
- **Visualization**: Power BI Matrix Map (`Matrix Map.pbix`)

## File Structure

```
├── Library Programming Matrix Map/
│     ├── Refresh Dashboard Dataset.py
│     ├── Refresh Matrix Dataset.py     # Main processing script
│     ├── Matrix Map.pbix              # Power BI dashboard
│     ├── lc_events_[date].csv        # LibCal export (input)
│     ├── MatrixMapDataset.xlsx       # Processed event data (output)
│     └── MatrixMapSettings.xlsx      # Configuration settings (output)

```

## Usage Workflow

### Step 1: Export Data from LibCal

1. **Access LibCal**
   - Log into your LibCal account
   - Navigate to: **Events Tab > Manatee Library Events > Event Explorer**

2. **Configure Export Filters**
   - **Date Range**: From 2022-01-01 To 2025-09-23 (or current date)
   - **Show Registration Responses**: No
   - Click **"Submit"** to apply filters

3. **Export Data**
   - Click **"Export Data"**
   - Save file with naming convention: `lc_events_[timestamp].csv`

### Step 2: File Management

- Move the exported CSV file to the Matrix Map directory:
  ```
  \Library Programming Matrix Map
  ```
- Ensure the file name starts with `lc_events_` for automatic detection

### Step 3: Process Data

1. Navigate to the Matrix Map directory
2. Run `Refresh Matrix Dataset.py`
3. The script will automatically:
   - Locate the LibCal CSV file
   - Process all event data
   - Generate output Excel files
   - Display processing status and completion time

### Step 4: Configure Settings (Optional)

1. Open `Matrix Map Settings.xlsx`
2. Adjust cost and impact scores for:
   - Program types
   - Room utilization
   - Event categories
   - Target audiences
   - Funding sources
3. Save and close the file

### Step 5: Update Dashboard

1. Open `Matrix Map.pbix` in Power BI
2. Refresh the data connections
3. Review updated matrix visualizations and insights

## Technical Specifications

### Input Data Structure

The system expects LibCal CSV exports with the following columns:
- Event ID, Title, Description
- Start/End dates and times
- Setup and teardown times
- Location and branch information
- Organizer and presenter details
- Audience and category classifications
- Registration and attendance data

### Output Data Structure

#### MatrixMapDataset.xlsx
Contains multiple worksheets:

- **EventInformation**: Core event details
- **EventAudiences**: Normalized audience data
- **EventCategories**: Event classification data
- **EventInternalTags**: Internal tagging system
- **EventTimes**: Time analysis and staff calculations
- **EventParticipation**: Registration and attendance metrics
- **EventProgram**: Automated program type classification

#### MatrixMapSettings.xlsx
Configuration worksheets:

- **Program Options**: Cost/impact scoring for program types
- **Room Utilization Options**: Location-based metrics
- **Category Options**: Event category scoring
- **Audience Options**: Target audience analysis
- **Funding Options**: Internal tag classifications
- **Branch Operational Hours By Year**: Operational metrics

### Program Classification System

The system automatically categorizes events into strategic program types:

1. **Story Time** - Literacy and storytelling programs
2. **Makerspace and Workshop** - Creative and hands-on activities
3. **Tech Support** - Technology assistance and training
4. **Book Club** - Reading discussion groups
5. **Reader's Advisory** - Literature guidance services
6. **Discovery Center** - Children's programming and STEAM activities
7. **Genealogy Services** - Family history research support
8. **Language and Culture** - ESL, cultural, and travel programs
9. **Local History & Archives** - Community heritage programs
10. **Fitness and Wellness** - Health and wellness initiatives
11. **Entertainment** - Games, movies, and social activities
12. **Nature and Home** - Environmental and domestic programs
13. **Music and Film** - Performing arts and cinema
14. **Important Meeting** - Board meetings and advisory sessions
15. **Life Skills and Community Resource** - Financial literacy, career services

## Key Features

### Automated Processing
- **File Detection**: Automatically locates LibCal export files
- **Data Validation**: Handles cancelled events and data inconsistencies
- **Time Calculations**: Computes event duration and total staff time
- **Classification**: Intelligent program type assignment using keyword matching

### Data Enrichment
- **Staff Time Analysis**: Includes setup, event, and teardown time
- **Fiscal Year Mapping**: Converts dates to fiscal year periods
- **Attendance Tracking**: Processes registration and actual attendance data
- **Resource Planning**: Calculates operational metrics

### Error Handling
- **Missing Files**: Clear error messages for missing data sources
- **Invalid Data**: Graceful handling of incomplete time information
- **Duplicate Prevention**: Removes existing files before regeneration

## Performance Metrics

- **Processing Speed**: Typical runtime under 5 seconds for annual datasets
- **Data Volume**: Handles thousands of events efficiently
- **Memory Usage**: Optimized for standard office computers
- **Compatibility**: Works with Excel 2016+ and Power BI Desktop

## Maintenance and Updates

### Regular Tasks
- **Monthly**: Update date filters in LibCal export
- **Quarterly**: Review program classification accuracy
- **Annually**: Audit MatrixMapSettings.xlsx configurations

### System Updates
- **Classification Rules**: Add new keywords to program type definitions
- **Data Structure**: Modify worksheets as reporting needs evolve
- **Performance**: Monitor processing times and optimize as needed

## Troubleshooting

### Common Issues

**"No lc_events_*.csv file found"**
- Ensure LibCal export file is in the correct directory
- Verify file naming convention starts with `lc_events_`

**Missing or Invalid Time Data**
- Check LibCal export for complete time information
- Review events with "All Day Event" settings

**Classification Accuracy**
- Update keyword lists in `classify_event_program_by_title()` function
- Add new program types as library services expand

## Version History

- **Current Version**: Supports LibCal exports through September 2025
- **Key Features**: Automated program classification, staff time calculation
- **Recent Updates**: Enhanced error handling, improved processing speed
