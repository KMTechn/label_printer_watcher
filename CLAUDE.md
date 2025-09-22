# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Windows label printer watcher application written in Python. It monitors specific folders for new PNG label files and automatically sends them to designated printers. The application features a Tkinter GUI, system tray functionality, and automatic updates from GitHub releases.

## Development Commands

### Environment Setup
```bash
# Install dependencies
pip install -r requirements.txt
```

### Running the Application
```bash
# Run the main application
python label_printer_watcher.py
```

### Building Executable
```bash
# Build standalone .exe with PyInstaller (includes assets folder)
pyinstaller --name "Label_Printer_Watcher" --onefile --add-data "assets;assets" label_printer_watcher.py
```

### Release Process
- Tag with version format `v*.*.*` (e.g., `v1.0.1`) to trigger GitHub Actions workflow
- The workflow automatically builds the executable and creates a release with ZIP file
- Built executable will be in `dist/` folder locally

## Architecture Overview

### Core Components

1. **Configuration Dictionary (CONFIG)**: Centralized configuration at the top of the file containing:
   - GitHub repository details for auto-updates
   - Application version
   - Printer names for remnant and defective labels
   - Base folder paths for monitoring

2. **File Monitoring System**:
   - Uses `watchdog` library with `LabelPrintHandler` class
   - Monitors date-specific subfolders (YYYY-MM-DD format)
   - Automatically creates daily folders if they don't exist
   - Implements debouncing to prevent duplicate prints

3. **GUI Application (App class)**:
   - Tkinter-based main window with scrolled text log display
   - System tray integration using `pystray`
   - Status bar showing current monitoring state
   - Redirects stdout/stderr to GUI log display

4. **Auto-Update System**:
   - Checks GitHub API for latest releases
   - Downloads and applies updates using batch script
   - Prompts user before updating
   - Handles application restart after update

5. **Printing System**:
   - Uses Windows `mspaint /p` command for printing
   - Threaded execution to prevent GUI blocking
   - Error handling for missing files or printer issues

### Key Features

- **Daily Folder Monitoring**: Automatically switches to monitor new date-specific folders at midnight
- **Duplicate Prevention**: Tracks recent prints to avoid duplicate printing within 2-second window
- **Resource Path Handling**: `resource_path()` function handles both development and PyInstaller executable environments
- **Threaded Operations**: Background threads for monitoring, update checking, and printing to keep GUI responsive

### File Structure

- `label_printer_watcher.py`: Single-file application containing all functionality
- `requirements.txt`: Python dependencies (requests, watchdog, pyinstaller)
- `assets/`: Contains logo files (logo.png, logo.ico) for GUI and tray icon
- `.github/workflows/release.yml`: GitHub Actions workflow for automated releases

### Configuration Notes

When modifying printer settings or folder paths, update the CONFIG dictionary at the top of `label_printer_watcher.py`. The application expects:
- Folder structure: `{base_folder}/{YYYY-MM-DD}/` for daily organization
- PNG files for label images
- Windows printer names exactly as they appear in system settings