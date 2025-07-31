"""
Build script for AquaExport Pro 2.1
Creates a standalone executable with all dependencies and proper file structure.
"""

import os
import sys
import shutil
import subprocess
import argparse
from pathlib import Path
import PyInstaller.__main__

def create_icon_if_missing():
    """Create a simple icon file if it doesn't exist."""
    icon_path = Path("icon.ico")
    if not icon_path.exists():
        print("  ℹ️  No icon.ico found - building without icon")
        return None
    return str(icon_path)

def build_executable(output_dir=None, work_dir=None, spec_dir=None):
    """Build the standalone executable."""
    
    # Parse command line arguments for output directories
    parser = argparse.ArgumentParser(description='Build AquaExport Pro 2.1 executable')
    parser.add_argument('--output-dir', '-o', 
                       help='Directory where the executable will be created (default: ./dist)')
    parser.add_argument('--work-dir', '-w',
                       help='Directory for temporary build files (default: ./build)')
    parser.add_argument('--spec-dir', '-s',
                       help='Directory for .spec files (default: current directory)')
    parser.add_argument('--optimize', action='store_true',
                       help='Enable size optimization')
    parser.add_argument('--installer', action='store_true',
                       help='Create Inno Setup installer (Windows only)')
    
    # Parse known args to avoid conflicts with PyInstaller
    args, unknown = parser.parse_known_args()
    
    # Use provided arguments or defaults
    output_dir = args.output_dir or output_dir or r'C:\Program Files\AquaExport'
    work_dir = args.work_dir or work_dir or 'build'
    spec_dir = args.spec_dir or spec_dir or '.'
    
    # Convert to Path objects
    output_path = Path(output_dir)
    work_path = Path(work_dir)
    spec_path = Path(spec_dir)

    print("🔨 Building AquaExport Pro 2.1 (Dual Mode)...")
    print("=" * 50)
    print(f"📁 Output directory: {output_path.absolute()}")
    print(f"🔧 Work directory: {work_path.absolute()}")
    print(f"📋 Spec directory: {spec_path.absolute()}")

    # Clean previous builds
    for dir_name in [work_path, output_path]:
        if dir_name.exists():
            shutil.rmtree(dir_name)
            print(f"  ✓ Cleaned {dir_name}/")

    # Check for icon
    icon_path = create_icon_if_missing()
    icon_args = ['--icon=' + icon_path] if icon_path else []

    # PyInstaller arguments
    pyinstaller_args = [
        'exporter.py',
        '--onefile',
        '--windowed',
        '--name=AquaExport Pro 2.1',
        '--add-data=templates:templates',  # Include templates directory
        '--hidden-import=tkinter',
        '--hidden-import=tkcalendar',
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=psycopg2',
        '--hidden-import=openpyxl',
        '--hidden-import=tomli',
        '--hidden-import=babel',
        '--hidden-import=babel.numbers',
        '--clean',
        '--noconfirm',
        f'--distpath={output_path}',
        f'--workpath={work_path}',
        f'--specpath={spec_path}',
    ] + icon_args

    # Add optimization flags if requested
    if args.optimize:
        pyinstaller_args.extend(['--optimize=2'])
        print("  ✓ Optimization enabled")

    # Run PyInstaller
    print("\n  → Running PyInstaller...")
    print("  ⏳ This may take a few minutes...")

    try:
        PyInstaller.__main__.run(pyinstaller_args)
    except Exception as e:
        print(f"\n❌ PyInstaller failed: {e}")
        return False

    # Prepare distribution directory
    if not output_path.exists() or not (output_path / 'AquaExport Pro 2.1.exe').exists():
        print("\n❌ Build failed - executable not created")
        return False

    print("\n✓ Executable built successfully!")

    # Create proper directory structure in output directory
    print("\n📁 Setting up distribution structure...")

    # Create templates directory
    templates_dir = output_path / 'templates'
    templates_dir.mkdir(exist_ok=True)

    # Copy templates if they exist
    template_files = [
        ('template.xlsx', 'kvaliteta_vode_template.xlsx'),  # Old name -> new name
        ('kvaliteta_vode_template.xlsx', 'kvaliteta_vode_template.xlsx'),
        ('dnevni_ocevidnik_template.xlsx', 'zahvacene_kolicine_vode_template.xlsx'),
        ('zahvacene_kolicine_vode_template.xlsx', 'zahvacene_kolicine_vode_template.xlsx')
    ]

    templates_found = 0
    for src_name, dst_name in template_files:
        # Check in root directory
        if Path(src_name).exists():
            shutil.copy2(src_name, templates_dir / dst_name)
            print(f"  ✓ Copied {src_name} → templates/{dst_name}")
            templates_found += 1
        # Check in templates directory
        elif (Path('templates') / src_name).exists():
            shutil.copy2(Path('templates') / src_name, templates_dir / dst_name)
            print(f"  ✓ Copied templates/{src_name} → templates/{dst_name}")
            templates_found += 1

    if templates_found == 0:
        print("  ⚠️  WARNING: No templates found! Users will need to add them manually.")

    # Copy default config
    if Path('config.toml').exists():
        shutil.copy2('config.toml', output_path / 'config.toml.default')
        print("  ✓ Copied config.toml.default")

    # Create README for distribution
    readme_content = """AquaExport Pro 2.1 - Upute za instalaciju
==========================================

🎯 BRZI START:
1. Pokrenite "AquaExport Pro 2.1.exe"
2. Pri prvom pokretanju će se stvoriti config.toml
3. Uredite config.toml sa svojim postavkama baze podataka
4. Ponovno pokrenite aplikaciju

📁 STRUKTURA MAPA:
Nakon prvog pokretanja, aplikacija će stvoriti:
  ./templates/              - Predlošci za izvoz
  ./exports/                - Izvezene datoteke
    ├── kvaliteta_vode/     - Excel datoteke kvalitete vode
    └── zahvacene_kolicine_vode/ - Excel datoteke količina

📋 PREDLOŠCI:
Aplikacija traži sljedeće predloške u mapi 'templates':
  • kvaliteta_vode_template.xlsx
  • zahvacene_kolicine_vode_template.xlsx

Ako nedostaju, molimo ih dodajte prije izvoza.

⚙️ KONFIGURACIJA:
Uredite config.toml sa svojim postavkama:
  • Podatke za spajanje na bazu
  • Putanje do mapa (ako želite promijeniti)

🛡️ WINDOWS DEFENDER:
Pri prvom pokretanju Windows može upozoriti na nepoznatu aplikaciju.
Kliknite "More info" → "Run anyway"

📞 PODRŠKA:
Email: neven.vrancic@fornax-automatika.hr
Tvrtka: FORNAX d.o.o.

verzija 2.1.0
"""

    with open(output_path / 'README.txt', 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print("  ✓ Created README.txt")

    # Create example exports directory structure
    exports_dir = output_path / 'exports'
    (exports_dir / 'kvaliteta_vode').mkdir(parents=True, exist_ok=True)
    (exports_dir / 'zahvacene_kolicine_vode').mkdir(parents=True, exist_ok=True)
    print("  ✓ Created example exports directory structure")

    # Calculate final size
    exe_path = output_path / 'AquaExport Pro 2.1.exe'
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)

        print("\n" + "=" * 50)
        print("✅ Build completed successfully!")
        print(f"\n📍 Location: {exe_path.absolute()}")
        print(f"📏 Size: {size_mb:.1f} MB")
        print(f"📁 Total files in output: {len(list(output_path.rglob('*')))}")

        if templates_found < 2:
            print("\n⚠️  IMPORTANT: Remember to add missing Excel templates!")

        print("\n🚀 Ready for distribution!")
        return True

    return False

def create_installer(output_dir=r'C:\Program Files\AquaExport'):
    """Create an optional installer using Inno Setup (if available)."""
    # This is optional - only if Inno Setup is installed
    iss_script = r"""[Setup]
AppName=AquaExport Pro
AppVersion=2.1.0
AppPublisher=FORNAX d.o.o.
DefaultDirName={pf}\AquaExport Pro
DefaultGroupName=AquaExport Pro
UninstallDisplayIcon={app}\AquaExport Pro 2.1.exe
Compression=lzma2
SolidCompression=yes
OutputDir=..\installer
OutputBaseFilename=AquaExportPro_2.1_Setup

[Files]
Source: "AquaExport Pro 2.1.exe"; DestDir: "{app}"
Source: "templates\*"; DestDir: "{app}\templates"; Flags: recursesubdirs
Source: "config.toml.default"; DestDir: "{app}"
Source: "README.txt"; DestDir: "{app}"; Flags: isreadme

[Icons]
Name: "{group}\AquaExport Pro 2.1"; Filename: "{app}\AquaExport Pro 2.1.exe"
Name: "{group}\Uninstall AquaExport Pro"; Filename: "{uninstallexe}"
Name: "{commondesktop}\AquaExport Pro 2.1"; Filename: "{app}\AquaExport Pro 2.1.exe"

[Run]
Filename: "{app}\AquaExport Pro 2.1.exe"; Description: "Launch AquaExport Pro"; Flags: nowait postinstall skipifsilent
"""

    # Check if Inno Setup is available
    inno_path = r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    if os.path.exists(inno_path):
        print("\n📦 Creating installer with Inno Setup...")

        # Write script
        with open(f'{output_dir}/setup.iss', 'w') as f:
            f.write(iss_script)

        # Create installer directory
        Path('installer').mkdir(exist_ok=True)

        # Run Inno Setup
        try:
            subprocess.run([inno_path, f'{output_dir}/setup.iss'], check=True)
            print("  ✓ Installer created successfully!")
        except subprocess.CalledProcessError:
            print("  ❌ Installer creation failed")

if __name__ == "__main__":
    print("\n🚀 AquaExport Pro 2.1 Build Tool")
    print("   Dual-mode water data exporter")
    print("\nUsage:")
    print("  python build.py [options]")
    print("\nOptions:")
    print("  -o, --output-dir DIR    Directory for executable (default: C:\\Program Files\\AquaExport)")
    print("  -w, --work-dir DIR      Directory for build files (default: ./build)")
    print("  -s, --spec-dir DIR      Directory for .spec files (default: current)")
    print("  --optimize              Enable size optimization")
    print("  --installer             Create Inno Setup installer (Windows only)")
    print("\nExamples:")
    print("  python build.py")
    print("  python build.py -o ./my_output")
    print("  python build.py --output-dir C:\\MyExecutables --optimize")
    print("")

    # Parse arguments for output directory
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--output-dir', '-o')
    parser.add_argument('--work-dir', '-w')
    parser.add_argument('--spec-dir', '-s')
    parser.add_argument('--optimize', action='store_true')
    parser.add_argument('--installer', action='store_true')
    
    args, _ = parser.parse_known_args()
    
    success = build_executable(
        output_dir=args.output_dir,
        work_dir=args.work_dir,
        spec_dir=args.spec_dir
    )

    if success and args.installer:
        create_installer(args.output_dir or r'C:\Program Files\AquaExport')

    print("\n" + ("✅ All done!" if success else "❌ Build failed!"))