"""
Build script for AquaExport Pro 2.1
Creates a standalone executable with all dependencies and proper file structure.
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path
import PyInstaller.__main__

def create_icon_if_missing():
    """Create a simple icon file if it doesn't exist."""
    icon_path = Path("icon.ico")
    if not icon_path.exists():
        print("  ℹ️  No icon.ico found - building without icon")
        return None
    return str(icon_path)

def build_executable():
    """Build the standalone executable."""

    print("🔨 Building AquaExport Pro 2.1 (Dual Mode)...")
    print("=" * 50)

    # Clean previous builds
    for dir_name in ['build', 'dist']:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"  ✓ Cleaned {dir_name}/")

    # Check for icon
    icon_path = create_icon_if_missing()
    icon_args = ['--icon=' + icon_path] if icon_path else []

    # PyInstaller arguments
    args = [
        'exporter.py',
        '--onefile',
        '--windowed',
        '--name=AquaExport Pro 2.1',
        '--add-data=templates;templates',  # Include templates directory
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
    ] + icon_args

    # Add optimization flags if requested
    if '--optimize' in sys.argv:
        args.extend(['--optimize=2'])
        print("  ✓ Optimization enabled")

    # Run PyInstaller
    print("\n  → Running PyInstaller...")
    print("  ⏳ This may take a few minutes...")

    try:
        PyInstaller.__main__.run(args)
    except Exception as e:
        print(f"\n❌ PyInstaller failed: {e}")
        return False

    # Prepare distribution directory
    dist_dir = Path('dist')
    if not dist_dir.exists() or not (dist_dir / 'AquaExport Pro 2.1.exe').exists():
        print("\n❌ Build failed - executable not created")
        return False

    print("\n✓ Executable built successfully!")

    # Create proper directory structure in dist
    print("\n📁 Setting up distribution structure...")

    # Create templates directory
    templates_dir = dist_dir / 'templates'
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
        shutil.copy2('config.toml', dist_dir / 'config.toml.default')
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

    with open(dist_dir / 'README.txt', 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print("  ✓ Created README.txt")

    # Create example exports directory structure
    exports_dir = dist_dir / 'exports'
    (exports_dir / 'kvaliteta_vode').mkdir(parents=True, exist_ok=True)
    (exports_dir / 'zahvacene_kolicine_vode').mkdir(parents=True, exist_ok=True)
    print("  ✓ Created example exports directory structure")

    # Calculate final size
    exe_path = dist_dir / 'AquaExport Pro 2.1.exe'
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)

        print("\n" + "=" * 50)
        print("✅ Build completed successfully!")
        print(f"\n📍 Location: {exe_path.absolute()}")
        print(f"📏 Size: {size_mb:.1f} MB")
        print(f"📁 Total files in dist: {len(list(dist_dir.rglob('*')))}")

        if templates_found < 2:
            print("\n⚠️  IMPORTANT: Remember to add missing Excel templates!")

        print("\n🚀 Ready for distribution!")
        return True

    return False

def create_installer():
    """Create an optional installer using Inno Setup (if available)."""
    # This is optional - only if Inno Setup is installed
    iss_script = """
[Setup]
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
    if os.path.exists(inno_path) and '--installer' in sys.argv:
        print("\n📦 Creating installer with Inno Setup...")

        # Write script
        with open('dist/setup.iss', 'w') as f:
            f.write(iss_script)

        # Create installer directory
        Path('installer').mkdir(exist_ok=True)

        # Run Inno Setup
        try:
            subprocess.run([inno_path, 'dist/setup.iss'], check=True)
            print("  ✓ Installer created successfully!")
        except subprocess.CalledProcessError:
            print("  ❌ Installer creation failed")

if __name__ == "__main__":
    print("\n🚀 AquaExport Pro 2.1 Build Tool")
    print("   Dual-mode water data exporter")
    print("\nOptions:")
    print("  --optimize    Enable size optimization")
    print("  --installer   Create Inno Setup installer (Windows only)")
    print("")

    success = build_executable()

    if success and '--installer' in sys.argv:
        create_installer()

    print("\n" + ("✅ All done!" if success else "❌ Build failed!"))