# AquaExport Pro 2.0

A modern, high-performance tool for exporting water quality data from PostgreSQL databases to Excel workbooks. Designed specifically for water treatment facilities with a beautiful GUI and blazing-fast performance.

## 🌟 Features

- **Beautiful Modern GUI** with intuitive date pickers and progress tracking
- **Blazing Fast Performance** using optimized bulk database operations
- **Idempotent File Handling** - creates or updates yearly workbooks seamlessly
- **Multi-Location Support** for different water treatment facilities
- **Comprehensive Error Handling** with detailed logging
- **Single-File Executable** support for easy deployment
- **Template-Based Export** using customizable Excel templates

## 📋 Supported Water Quality Parameters

- **Mutnoca** (Turbidity)
- **Klor** (Chlorine)
- **Temperature**
- **pH**
- **Redox** (Oxidation-Reduction Potential)

## 🏭 Supported Locations

- **PK Barbat** - Primary water treatment facility
- **VS Lopar** - Water station Lopar
- **VS Perici** - Water station Perici

## 🚀 Quick Start

### Prerequisites

- **Windows 10/11** (primary platform)
- **PostgreSQL** database with water quality data
- **Python 3.8+** (for development)

### Installation

#### Option 1: Download Executable (Recommended)
1. Download the latest release from the releases page
2. Extract the ZIP file to your desired location
3. Ensure `template.xlsx` is in the same directory
4. Run `AquaExport Pro 2.0.exe`

#### Option 2: Build from Source
```bash
# Clone the repository
git clone <repository-url>
cd AquaExport

# Install dependencies
pip install -r requirements.txt

# Run the application
python exporter.py
```

### Configuration

1. **First Run**: The application will create a `config.toml` file
2. **Edit Configuration**: Open `config.toml` and update database settings:

```toml
[database]
host = "your-database-host"
port = 5432
name = "your-database-name"
user = "your-username"
password = "your-password"

[export]
directory = "./exports"
template_path = "./template.xlsx"
```

## 📖 Usage

### Basic Export
1. **Launch** the application
2. **Select Date Range** using the date pickers
3. **Choose Export Options**:
   - Quick export (last 7/30/90 days)
   - Custom date range
4. **Click Export** and monitor progress
5. **Find Results** in the exports directory

### Advanced Features
- **Yearly Workbooks**: Data is organized by year in separate Excel files
- **Template Support**: Uses customizable Excel templates for consistent formatting
- **Progress Tracking**: Real-time progress updates during export
- **Error Recovery**: Automatic retry and detailed error logging

## 🛠️ Development

### Project Structure
```
AquaExport/
├── exporter.py          # Main application logic
├── build.py            # Build script for executable
├── config.toml         # Configuration file
├── requirements.txt    # Python dependencies
├── template.xlsx       # Excel template (not included)
└── README.md          # This file
```

### Building Executable
```bash
# Install build dependencies
pip install pyinstaller

# Build executable
python build.py

# Build with optimization
python build.py --optimize
```

### Database Schema
The application expects a PostgreSQL database with the following structure:
- **Table**: `floattable`
- **Columns**: 
  - `dateandtime` (timestamp)
  - `tagindex` (integer)
  - `val` (float)

### Tag Mappings
Water quality parameters are mapped to database tag indices:

| Location | Mutnoca | Klor | Temp | pH | Redox |
|----------|---------|------|------|----|-------|
| PK Barbat | 3 | 21 | 134 | 132 | 133 |
| VS Lopar | 151 | 155 | 156 | - | - |
| VS Perici | - | 72 | 82 | - | 81 |

## 🔧 Configuration Options

### Database Settings
- `host`: PostgreSQL server address
- `port`: Database port (default: 5432)
- `name`: Database name
- `user`: Database username
- `password`: Database password

### Export Settings
- `directory`: Output directory for Excel files
- `template_path`: Path to Excel template file

### Tag Mappings (Optional)
You can override default tag mappings in `config.toml`:
```toml
[tag_mappings.pk_barbat]
mutnoca = 3
klor = 21
temp = 134
pH = 132
redox = 133
```

## 📊 Output Format

### Excel Structure
- **Yearly Workbooks**: One file per year (e.g., `2024_water_quality.xlsx`)
- **Monthly Sheets**: Data organized by month in Croatian
- **Location Blocks**: Separate sections for each facility
- **Daily Aggregates**: Min, Max, and Average values per day

### File Naming Convention
- Format: `YYYY_water_quality.xlsx`
- Example: `2024_water_quality.xlsx`

## 🐛 Troubleshooting

### Common Issues

#### Database Connection Failed
- Verify database credentials in `config.toml`
- Ensure PostgreSQL is running and accessible
- Check firewall settings

#### Template File Missing
- Ensure `template.xlsx` is in the application directory
- Verify template path in configuration

#### Permission Errors
- Run as administrator if needed
- Check write permissions for export directory

#### Windows Defender Warning
- This is normal for new executables
- Click "More info" → "Run anyway"

### Logs
- Log files are created in the export directory
- Check `exporter.log` for detailed error information
- Logs rotate automatically (max 10MB, 5 backups)

## 📝 Changelog

### Version 2.0.0
- Complete GUI redesign with modern interface
- Improved performance with bulk database operations
- Enhanced error handling and logging
- Template-based Excel export
- Multi-location support
- Single-file executable

### Version 1.x
- Basic command-line export functionality
- Simple Excel output


## 📄 License

This project is proprietary software developed for water quality management systems.

## 📞 Support

For technical support or questions:
- **Email**: neven.vrancic@fornax-automatika.hr
- **Company**: Fornax Automatika

## 🔒 Security Notes

- Database credentials are stored in plain text in `config.toml`
- Consider using environment variables for production deployments
- Ensure proper access controls on configuration files

---

**AquaExport Pro 2.0** - Professional water quality data export solution 