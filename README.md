# Workset Organizer for Autodesk Revit

A powerful Revit add-in that automates MEP workset organization, model export, and template integration for streamlined project handover and standardization.


## 🚀 Features

### Core Functionality
- **Automated Workset Organization**: Intelligently assigns MEP elements to worksets based on Excel-defined patterns
- **System-Specific Exports**: Creates clean, package-specific Revit files for subcontractor handover
- **Template Integration**: Merges exported MEP elements into standardized project templates
- **Quality Control**: Automatically manages orphaned elements in dedicated QC worksets
- **Comprehensive Logging**: Real-time process monitoring with detailed operation logs

### Supported MEP Systems
- **Piping**: Fittings, Accessories, Curves & Systems
- **HVAC**: Ducts, Fittings, Terminals & Equipment
- **Electrical**: Equipment, Fixtures, Lighting, Cable Trays
- **Plumbing**: Fixtures, Sprinklers & Accessories
- **Mechanical**: All MEP equipment categories

## 📋 Requirements

### Software
- **Autodesk Revit**: 2023 or later
- **.NET Framework**: 4.8+
- **Windows**: 10/11 (64-bit)
- **Microsoft Excel**: For mapping configuration

### Hardware
- **RAM**: 8GB minimum, 16GB+ recommended
- **Storage**: 1GB+ free space for exports
- **CPU**: Multi-core processor recommended

## 🔧 Installation

### Quick Install
1. Download the latest release from the [Releases](../../releases) page
2. Extract files to: `C:\ProgramData\Autodesk\Revit\Addins\2023\`
3. Restart Revit
4. Look for **Workset Tools** tab or **External Tools** → **Workset Orchestrator**

### Manual Installation
```bash
# Download required files
WorksetOrganizer.dll
WorksetOrganizer.addin

# Copy to Revit add-ins directory
C:\ProgramData\Autodesk\Revit\Addins\2023\

# Verify installation
# Launch Revit and check for Workset Tools ribbon tab
```

## 📖 Quick Start

### 1. Prepare Excel Mapping File
Create an Excel file with a "Mapping" worksheet containing:

| Workset Name | System Name in Model File | System Description | Model iFLS/Package Code |
|--------------|---------------------------|-------------------|------------------------|
| HW_Supply_01 | HWS-xxx | Hot Water Supply | HWS |
| CW_Supply_01 | CWS-xxx | Cold Water Supply | CWS |
| HVAC_Supply_01 | HVAC-SUP-xx | HVAC Supply Air | HVAC-SUP |

### 2. Launch Plugin
- Open workshared Revit document
- Navigate to **Workset Tools** → **Workset Organizer**
- Or use **External Tools** → **Workset Orchestrator**

### 3. Configure and Run
1. Select Excel mapping file
2. Choose export destination folder
3. Set options (overwrite files, export QC)
4. Click **Run**

### 4. Template Integration (Optional)
1. After successful export, click **Integrate into Template**
2. Select template file
3. Integration creates files in "In Template" subfolder

## 🔍 Pattern Matching

### Wildcard Patterns
- `x` → Single digit (0-9)
- `xx` → 1-3 digits  
- `xxx` → 2-3 digits
- `*` → Any characters

### Examples
```
HWS-xxx     → Matches: HWS-001, HWS-012, HWS-123
CWS-xx      → Matches: CWS-1, CWS-12, CWS-123  
HVAC-*      → Matches: HVAC-Supply-A, HVAC-Return-B
```

### Special Codes
- `NO EXPORT` → Organize elements but skip export

## 📁 Project Structure

```
WorksetOrganizer/
├── App.cs                    # Revit application entry point
├── Command.cs                # External command implementation
├── MainForm.xaml            # WPF user interface
├── MainForm.xaml.cs         # UI logic and event handling
├── WorksetOrchestrator.cs   # Core orchestration logic
├── ExcelReader.cs           # Excel file processing
├── MappingRecord.cs         # Data model for mapping
├── WorksetEventHandler.cs   # External event handling
├── Properties/
│   └── AssemblyInfo.cs      # Assembly information
├── packages.config          # NuGet package references
├── WorksetOrganizer.csproj  # Project file
└── app.config              # Application configuration
```

## 🔄 Workflow Process

### Phase 1: Organization
1. **Element Collection**: Gathers all MEP elements from model
2. **Pattern Matching**: Matches elements to patterns using system names
3. **Workset Assignment**: Moves elements to appropriate worksets
4. **QC Processing**: Assigns orphaned elements to DX_QC workset

### Phase 2: Export
1. **Synchronization**: Syncs with central and relinquishes ownership
2. **Package Creation**: Creates new documents per package
3. **Element Transfer**: Copies specific elements using Revit API
4. **File Generation**: Saves clean package-specific files

### Phase 3: Template Integration
1. **Template Processing**: Opens template for each export file
2. **Element Integration**: Copies MEP elements into template
3. **Workset Preservation**: Maintains workset structure
4. **Output Generation**: Saves integrated files with standards

## 🛠️ Configuration

### Excel Mapping File Format
```excel
Required Headers:
- Workset Name (text)
- System Name in Model File (pattern)
- System Description (optional text)
- Model iFLS/Package Code (export identifier)

Worksheet Name: "Mapping" (case-insensitive)
File Format: .xlsx
```

### Export File Naming
```
{ProjectPrefix}_{PackageCode}_MO_Part_001_DX.rvt

Examples:
INFINEON_HWS_MO_Part_001_DX.rvt
INFINEON_CWS_MO_Part_001_DX.rvt
INFINEON_QC_MO_Part_001_DX.rvt
```

## 📊 Performance

### Typical Processing Times
- **Small Models** (< 1000 elements): 2-5 minutes
- **Medium Models** (1000-5000 elements): 5-15 minutes  
- **Large Models** (5000+ elements): 15-30 minutes
- **Template Integration**: +50% additional time

### Memory Usage
- **Base Usage**: 100-200MB
- **Per Export**: 50-100MB additional
- **Peak Usage**: Up to 2GB for very large models

## 🐛 Troubleshooting

### Common Issues

#### Plugin Not Visible
- Verify installation path: `C:\ProgramData\Autodesk\Revit\Addins\2023\`
- Check both `.dll` and `.addin` files are present
- Run Revit as administrator
- Try **External Tools** → **Workset Orchestrator**

#### Excel File Errors
- Use `.xlsx` format (not `.xls`)
- Ensure "Mapping" worksheet exists
- Verify required column headers
- Close Excel before running plugin

#### Export Failures
- Ensure document is workshared
- Check available disk space (1GB+ required)
- Verify write permissions for destination folder
- Close other Revit sessions to free memory

#### Template Integration Issues
- Verify template file is accessible and not corrupted
- Ensure sufficient system memory (8GB+ recommended)
- Check that exported files exist before integration
- Monitor log for specific error messages

### Log File Analysis
Check `WorksetOrchestrationLog.txt` in destination folder for detailed error information:
- **"ERROR"** entries indicate issues requiring attention
- **"WARNING"** entries are non-critical but worth reviewing
- **"PROCESS COMPLETED SUCCESSFULLY"** confirms successful completion

## 📝 Development

### Building from Source
```bash
# Prerequisites
Visual Studio 2019/2022
.NET Framework 4.8 SDK
Revit 2023 SDK

# Clone repository
git clone https://github.com/your-org/workset-organizer.git
cd workset-organizer

# Restore NuGet packages
nuget restore

# Build solution
msbuild WorksetOrganizer.sln /p:Configuration=Release
```

### Development Setup
1. Install Revit 2023 and SDK
2. Clone repository to local machine
3. Update Revit API references in project file
4. Build and copy output to Revit add-ins folder
5. Launch Revit for testing

### Dependencies
- **EPPlus 4.5.3.3**: Excel file processing
- **Revit API 2023**: Core Revit integration
- **WPF**: User interface framework
- **.NET Framework 4.8**: Runtime platform

## 🤝 Contributing

### Bug Reports
Please include:
- Revit version and build number
- Plugin version
- Excel mapping file (if relevant)
- Complete log file from failed operation
- Steps to reproduce the issue

### Feature Requests
1. Open an issue describing the feature
2. Include use case and business justification
3. Provide mockups or examples if applicable
4. Discuss implementation approach

### Pull Requests
1. Fork the repository
2. Create feature branch (`git checkout -b feature/new-feature`)
3. Implement changes with appropriate tests
4. Update documentation if needed
5. Submit pull request with clear description

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🏢 Company Information

**DEAXO GmbH**
- **Product**: Workset Organizer for Autodesk Revit
- **Version**: 1.0.0
- **Support**: Contact BIM Coordination team

## 📚 Documentation

### Complete Documentation
- [Workflow Document](docs/XXX_BIM_BA_Workflow_007_DX.md) - Complete workflow procedures
- [Installation Guide](docs/installation.md) - Detailed installation instructions
- [Configuration Guide](docs/configuration.md) - Excel mapping setup
- [API Documentation](docs/api.md) - Developer reference

### Quick Reference
- [Excel Template](templates/WorkSet_Mapping.xlsx) - Standard mapping file template
- [Pattern Examples](docs/pattern-examples.md) - Common naming patterns
- [Troubleshooting](docs/troubleshooting.md) - Common issues and solutions

## 🔗 Related Projects

- [Revit API Documentation](https://www.revitapidocs.com/)
- [EPPlus Library](https://github.com/EPPlusSoftware/EPPlus)
- [Revit Add-in Guidelines](https://help.autodesk.com/view/RVT/2023/ENU/?guid=Revit_API_Revit_API_Developers_Guide_html)

## 📈 Changelog

### Version 1.0.0 (September 2025)
- ✨ **New**: Template integration functionality
- ✨ **New**: Element-based export engine (replaces view-based)
- ✨ **New**: Enhanced progress tracking with timeout protection
- 🔧 **Improved**: Better error handling and recovery
- 🔧 **Improved**: More intuitive user interface
- 🔧 **Improved**: Comprehensive logging system
- 🐛 **Fixed**: Memory issues with large models
- 🐛 **Fixed**: Synchronization problems with central files

### Version 0.1.0 (Initial Release)
- ✨ Initial workset organization functionality
- ✨ Basic export capabilities
- ✨ Excel mapping integration

## 🎯 Roadmap

### Planned Features
- **Multi-language Support**: Interface localization
- **Advanced Filtering**: Custom element selection criteria
- **Batch Processing**: Multiple project processing
- **Cloud Integration**: SharePoint/BIM 360 connectivity
- **Reporting Dashboard**: Analytics and insights
- **Custom Templates**: User-defined export templates

### Performance Improvements
- **Parallel Processing**: Multi-threaded operations
- **Memory Optimization**: Reduced memory footprint
- **Incremental Updates**: Process only changed elements
- **Background Processing**: Non-blocking operations

## 🆘 Support

### Internal Support Contacts
- **BIM Coordinator**: Abdul Rahman
- **BIM Manager**: Vital Kavaliou

### Self-Service Resources
- Check log files in destination folder
- Review troubleshooting section above
- Consult workflow documentation
- Verify Excel mapping file format

### Emergency Support
For critical project issues:
1. Collect all relevant log files
2. Document exact error messages
3. Note Revit version and model size
4. Contact BIM coordination team immediately

---

## 📞 Quick Contact

**Issues?** Open a [GitHub Issue](../../issues)  
**Questions?** Check [Documentation](docs/)  
**Updates?** Watch this repository for releases

---

*Built with ❤️ by the DEAXO GmbH BIM Team*
