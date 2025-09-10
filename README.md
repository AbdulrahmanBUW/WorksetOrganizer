# Workset Organizer for Autodesk Revit

**Version 1.0.0** | **DEAXO GmbH**

A Revit add-in for MEP workset management, model extraction, and project deliverable generation.

## Overview

The Workset Organizer handles BIM coordination workflows with two operation modes: quality-controlled element organization using Excel mapping, and direct workset extraction. Built for projects requiring consistent deliverables and workset management across disciplines.

## Features

### Operation Modes

**QC Check & Extraction**
- Excel-based element categorization using pattern matching
- Workset assignment based on system names
- Quality control for unassigned elements
- Export generation for subcontractor packages

**Extract Worksets**
- Direct extraction of existing worksets
- Maintains current workset structure
- Applies iFLS codes for standard naming
- No Excel mapping required

### Element Handling

Processes MEP systems (piping, ducting, electrical), structural components, generic models, and specialty equipment. Pattern recognition handles system naming conventions while maintaining element relationships during extraction.

### Template Integration

Integrates extracted elements into project templates while preserving workset assignments and system relationships.

## Requirements

### Software
- Autodesk Revit 2023 or later
- Microsoft .NET Framework 4.8
- Windows 10/11 (64-bit)
- Microsoft Excel (for QC mode)

### Hardware
- 8GB RAM minimum (16GB recommended)
- 1GB free disk space
- Multi-core processor recommended

## Installation

### Setup
1. Copy files to Revit add-ins folder:
   ```
   C:\ProgramData\Autodesk\Revit\Addins\2023\
   ```
2. Required files:
   - `WorksetOrganizer.dll`
   - `WorksetOrganizer.addin`
3. Restart Revit

### Access
Find the tool in **Workset Tools** ribbon or **External Tools** → **Workset Orchestrator**

## Configuration

### QC Mode: Excel Mapping File

Create Excel file with "Mapping" worksheet:

| Workset Name | System Name in Model File | System Description | Model iFLS/Package Code |
|--------------|---------------------------|-------------------|------------------------|
| DX_HW_Supply_01 | HWS-xxx | Hot Water Supply | HWS |
| DX_CW_Supply_01 | CWS-xxx | Cold Water Supply | CWS |
| DX_ELT | ELT | Electrical Systems | ELT |

**Pattern Options:**
- `x` = single digit
- `xx` = 1-3 digits  
- `xxx` = 2-3 digits
- `*` = any characters
- `NO EXPORT` = organize but don't export

### Extract Mode: Built-in iFLS Codes

| Workset | iFLS Code | Description |
|---------|-----------|-------------|
| DX_ELT | E-X | Electrical |
| DX_CHM | C-S | Chemical Systems |
| DX_PAW | S-D | Process Air & Water |
| DX_STB | S-T | Structural |

## Usage

### QC Check & Extraction

1. **Setup**
   - Open workshared Revit model
   - Prepare Excel mapping file
   - Set destination folder

2. **Run Process**
   - Start add-in, select QC mode
   - Browse to Excel file
   - Choose export options
   - Click "QC Check & Extraction"

3. **Review Results**
   - Check processing log
   - Verify exported files
   - Validate element assignments

### Extract Worksets

1. **Preparation**
   - Switch to Extract mode
   - Set destination folder
   - Confirm workset detection

2. **Execute**
   - Click "Extract Worksets"
   - Monitor progress
   - Review generated files

3. **Validation**
   - Check extraction log
   - Confirm file naming
   - Verify workset preservation

### Template Integration

After export completion:
1. Click "Integrate into Template"
2. Select template file
3. Review integrated files in "In Template" folder

## File Naming

Format: `{ProjectPrefix}_{iFLSCode}_MO_Part_001_DX.rvt`

Examples:
- `CMPP64-47_HWS_MO_Part_001_DX.rvt`
- `CMPP64-47_E-X_MO_Part_001_DX.rvt`
- `CMPP64-47_QC_MO_Part_001_DX.rvt`

## Performance

### Processing Time
- Small models (<1,000 elements): 2-5 minutes
- Medium models (1,000-5,000 elements): 5-15 minutes
- Large models (>5,000 elements): 15-30 minutes
- Template integration: adds 50% to processing time

### Memory Usage
- Base: 100-200MB
- Per export: 50-100MB additional
- Large models: up to 2GB peak usage

## Troubleshooting

### Common Problems

**Add-in Not Found**
- Check file location in correct Revit add-ins folder
- Verify both .dll and .addin files present
- Run Revit as administrator
- Look under External Tools menu

**Excel Errors**
- Use .xlsx format only
- Include "Mapping" worksheet
- Match required column headers exactly
- Close Excel before running

**Export Failures**
- Ensure model has worksharing enabled
- Check disk space (1GB+ needed)
- Verify folder write permissions
- Close other Revit sessions

**Memory Issues**
- Close unnecessary programs
- Use 64-bit Revit version
- Process smaller batches for large models

### Log Files

Check these files in destination folder:
- `WorksetOrchestrationLog.txt` (QC mode)
- `WorksetExtractionLog.txt` (Extract mode)

**Log Types:**
- ERROR: Problems requiring action
- WARNING: Issues to review
- INFO: Normal operation messages

## Technical Details

### Architecture
- **WorksetOrchestrator.cs**: Main processing logic
- **ExcelReader.cs**: Excel file handling
- **WorksetEventHandler.cs**: Revit API integration
- **MainForm.xaml**: User interface

### Dependencies
- EPPlus 4.5.3.3 (Excel processing)
- Revit API 2023
- WPF (user interface)
- .NET Framework 4.8

### Pattern Matching
The system uses regular expressions to match system names against Excel patterns. For electrical elements without system names, it identifies elements by category (cable trays, electrical equipment, etc.).

### Workset Management
Creates new worksets as needed and moves elements using Revit's workset parameter. Maintains element relationships during copy operations between documents.

## Development

### Build Requirements
- Visual Studio 2019/2022
- .NET Framework 4.8 SDK
- Revit 2023 SDK

### Building
```bash
git clone [repository]
cd workset-organizer
nuget restore
msbuild WorksetOrganizer.sln /p:Configuration=Release
```

Copy output files to Revit add-ins folder for testing.

## Support

### Contact
- BIM Coordination Team: Technical support
- Development Team: Code issues and enhancements

### Resources
- Check log files for error details
- Review Excel mapping format
- Verify workset naming conventions
- Confirm file permissions

---

**DEAXO GmbH**  
*Developed by the BIM Team with <3*
