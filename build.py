"""
Build script for AquaExport Pro 2.0
Creates a standalone executable with all dependencies.
"""

import os
import sys
import shutil
import PyInstaller.__main__

def build_executable():
    """Build the standalone executable."""
    
    print("🔨 Building AquaExport Pro 2.0...")
    
    # Clean previous builds
    for dir in ['build', 'dist']:
        if os.path.exists(dir):
            shutil.rmtree(dir)
            print(f"  ✓ Cleaned {dir}/")
    
    # Create icon if it doesn't exist (placeholder)
    if not os.path.exists('icon.ico'):
        print("  ⚠️  No icon.ico found, building without icon")
        icon_arg = []
    else:
        icon_arg = ['--icon=icon.ico']
    
    # PyInstaller arguments
    args = [
        'exporter.py',
        '--onefile',
        '--windowed',
        '--name=AquaExport Pro 2.0',
        '--add-data=template.xlsx;.',
        '--hidden-import=tkinter',
        '--hidden-import=tkcalendar', 
        '--hidden-import=PIL',
        '--hidden-import=psycopg2',
        '--hidden-import=openpyxl',
        '--hidden-import=tomli',
        '--hidden-import=babel',
        '--hidden-import=babel.numbers',
        '--clean',
        '--noconfirm',
    ] + icon_arg
    
    # Add optimization flags
    if '--optimize' in sys.argv:
        args.extend(['--optimize=2'])
        print("  ✓ Optimization enabled")
    
    # Run PyInstaller
    print("  → Running PyInstaller...")
    PyInstaller.__main__.run(args)
    
    # Copy additional files to dist
    dist_dir = 'dist'
    if os.path.exists(dist_dir):
        # Copy template
        if os.path.exists('template.xlsx'):
            shutil.copy2('template.xlsx', dist_dir)
            print(f"  ✓ Copied template.xlsx to {dist_dir}/")
        
        # Copy default config
        if os.path.exists('config.toml'):
            shutil.copy2('config.toml', f'{dist_dir}/config.toml.default')
            print(f"  ✓ Copied config.toml.default to {dist_dir}/")
        
        # Create README for distribution
        with open(f'{dist_dir}/README.txt', 'w', encoding='utf-8') as f:
            f.write("""AquaExport Pro 2.0 - Upute za instalaciju
=========================================

1. INSTALACIJA:
   - Kopirajte sve datoteke u željenu mapu
   - Provjerite da imate template.xlsx u istoj mapi

2. KONFIGURACIJA:
   - Pri prvom pokretanju će se stvoriti config.toml
   - Uredite config.toml sa svojim postavkama baze podataka

3. POKRETANJE:
   - Dvostruki klik na "AquaExport Pro 2.0.exe"
   - Ili pokrenite iz command line-a

4. NAPOMENE:
   - Windows Defender može upozoriti pri prvom pokretanju
   - Kliknite "More info" → "Run anyway"
   - Ovo je normalno za nove .exe datoteke

Za pomoć: neven.vrancic@fornax-automatika.hr
""")
        print(f"  ✓ Created README.txt in {dist_dir}/")
        
        print("\n✅ Build completed successfully!")
        print(f"   Executable location: {dist_dir}/AquaExport Pro 2.0.exe")
        
        # Calculate size
        exe_path = f"{dist_dir}/AquaExport Pro 2.0.exe"
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"   Size: {size_mb:.1f} MB")
    else:
        print("\n❌ Build failed - dist directory not created")

if __name__ == "__main__":
    build_executable()