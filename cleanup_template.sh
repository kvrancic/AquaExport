#!/bin/bash

echo "========================================"
echo "AquaExport Template Cleanup Script"
echo "========================================"
echo

# Check if Python is available
if ! command -v python3 &> /dev/null; then
    if ! command -v python &> /dev/null; then
        echo "ERROR: Python is not installed or not in PATH"
        echo "Please install Python 3.8+ and try again"
        exit 1
    else
        PYTHON_CMD="python"
    fi
else
    PYTHON_CMD="python3"
fi

# Check if template.xlsx exists
if [ ! -f "template.xlsx" ]; then
    echo "ERROR: template.xlsx not found in current directory"
    echo "Please ensure template.xlsx is in the same folder as this script"
    exit 1
fi

echo "Starting template cleanup..."
echo

# Run the cleanup script
$PYTHON_CMD cleanup_template.py

if [ $? -ne 0 ]; then
    echo
    echo "ERROR: Cleanup failed. Please check the error messages above."
    exit 1
fi

echo
echo "Cleanup completed successfully!"
echo "The template has been cleaned and is ready for use."
echo