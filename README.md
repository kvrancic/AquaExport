# AquaExport Pro 2.1 - Dual Mode Water Data Exporter

A modern, high-performance tool for exporting water quality and quantity data from PostgreSQL to Excel. Features dual-mode operation with dynamic UI theming.

![Python](https://img.shields.io/badge/python-3.8+-blue.svg)
![Version](https://img.shields.io/badge/version-2.1.0-brightgreen.svg)

## 🚀 New in Version 2.1

- **Dual Mode Operation**: Switch between Water Quality and Water Quantities
- **Dynamic Color Themes**: Blue theme for quality, green theme for quantities
- **Organized File Structure**: Separate directories for templates and export types
- **Enhanced Data Queries**: Optimized time-window queries for cumulative counters
- **Improved Error Handling**: Better detection of open Excel files

## 📋 Features

### Water Quality Mode (Kvaliteta vode)
- **Parameters**: Turbidity, Chlorine, Temperature, pH, Redox
- **Locations**: PK Barbat, VS Lopar, VS Perici
- **Output**: `kvaliteta_vode_YYYY.xlsx`
- **Theme**: Blue interface

### Water Quantities Mode (Zahvaćene količine vode)
- **Parameters**: Daily volume (m³), Max flow rate (l/s)
- **Locations**: Hrvatsko primorje južni ogranak, Perići, Gvačići I, Mlinica
- **Output**: `zahvacene_kolicine_YYYY.xlsx`
- **Theme**: Green interface

## 🏗️ Architecture

### Directory Structure
```
AquaExport/
├── templates/
│   ├── kvaliteta_vode_template.xlsx
│   └── zahvacene_kolicine_vode_template.xlsx
├── exports/
│   ├── kvaliteta_vode/
│   │   └── kvaliteta_vode_2024.xlsx
│   └── zahvacene_kolicine_vode/
│       └── zahvacene_kolicine_2024.xlsx
├── exporter.py
├── config.toml
└── README.md
```

### Database Schema
- **floattable**: Time-series measurements
  - `dateandtime`: Timestamp (UTC)
  - `tagindex`: Sensor identifier
  - `val`: Measured value
- **tagtable**: Sensor definitions
  - `tagindex`: Unique ID
  - `tagname`: Human-readable name

## 🚀 Quick Start

### Installation

#### Option 1: Pre-built Executable
1. Download the latest release
2. Extract to desired location
3. Ensure Excel templates are in `templates/` directory
4. Run `AquaExport Pro 2.1.exe`

#### Option 2: From Source
```bash
# Clone repository
git clone <repository-url>
cd aquaexport-pro

# Install dependencies
pip install -r requirements.txt

# Run application
python exporter.py
```

### Configuration

Edit `config.toml`:

```toml
[database]
host = "localhost"
port = 5432
name = "SCADA_arhiva_rab"
user = "postgres"
password = "your_password"

[export]
directory = "./exports"
template_dir = "./templates"

# Water quality tag mappings
[tag_mappings.pk_barbat]
mutnoca = 3
klor = 21
temp = 134
pH = 132
redox = 133
```

## 📊 Data Processing Logic

### Water Quality
- Standard daily aggregation (MIN/MAX/AVG)
- Query window: 00:00 - 23:59 each day

### Water Quantities
- **Daily Volume**: Cumulative counter that resets nightly
  - Query window: 22:00 today to 03:00 tomorrow
  - Takes MAX value in this window
- **Max Flow**: Instantaneous readings
  - Query window: 00:00 - 23:59
  - Takes MAX value during the day

## 🛠️ Development

### Building from Source

**Default Behavior:** The executable will be created in `C:\Program Files\AquaExport\` by default, giving it a professional installation location.

```bash
# Basic build (default: C:\Program Files\AquaExport)
python build.py

# Custom output directory
python build.py -o ./my_output
python build.py --output-dir C:\MyExecutables

# Optimized build
python build.py --optimize

# With installer (requires Inno Setup)
python build.py --installer

# Full custom build
python build.py -o ./custom_output -w ./custom_build --optimize
```

#### Build Options
- `-o, --output-dir DIR`: Specify where the executable will be created (default: `C:\Program Files\AquaExport`)
- `-w, --work-dir DIR`: Specify directory for temporary build files (default: `./build`)
- `-s, --spec-dir DIR`: Specify directory for .spec files (default: current directory)
- `--optimize`: Enable size optimization
- `--installer`: Create Inno Setup installer (Windows only)

### Tag Mappings

#### Water Quality Tags
| Location | Turbidity | Chlorine | Temperature | pH | Redox |
|----------|-----------|----------|-------------|-----|-------|
| PK Barbat | 3 | 21 | 134 | 132 | 133 |
| VS Lopar | - | 151 | 155 | - | 156 |
| VS Perici | - | 72 | 82 | - | 81 |

#### Water Quantity Tags
| Location | Daily In | Daily Out | Max Flow In | Max Flow Out |
|----------|----------|-----------|-------------|--------------|
| Hr. primorje južni ogranak | 14 | 13 | 18 | 16 |
| Perići | 67 | - | 68 | - |
| Gvačići I | 103 | - | 0 | - |
| Mlinica | 51 | - | 52 | - |

## 🔧 Troubleshooting

### Common Issues

1. **Missing Templates**
   - Ensure both Excel templates are in `templates/` directory
   - Templates must be named exactly as specified

2. **Excel File Locked**
   - Close any open Excel files before exporting
   - The app will warn if files are potentially open

3. **Database Connection**
   - Verify PostgreSQL is running
   - Check credentials in `config.toml`
   - Ensure network connectivity

4. **Wrong Data Values**
   - For daily volumes, check if counters reset as expected
   - Verify tag mappings match your SCADA configuration
   - Check timezone settings (data is stored in UTC)

### Log Files
- Location: `exports/exporter.log`
- Rotation: 5 files × 10MB each
- Contains detailed debug information

## 📝 Migration from v1.x

When upgrading from version 1.x:

1. The app will automatically migrate your old `template.xlsx`
2. Place the new quantities template in the templates directory
3. Old exports remain in their original location
4. Update `config.toml` to use `template_dir` instead of `template_path`

## 🔒 Security

- Database credentials stored in `config.toml` (keep secure)
- No passwords in GUI - all automated
- Consider using environment variables in production
- Restrict file permissions appropriately

## 📄 License

Proprietary software for water management systems.

## 📞 Support

**Technical Support**: neven.vrancic@fornax-automatika.hr  
**Company**: FORNAX d.o.o.  
**Website**: www.fornax-automatika.hr

---

**AquaExport Pro 2.1** - Professional dual-mode water data export solution