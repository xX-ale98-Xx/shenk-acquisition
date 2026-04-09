# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Python tkinter GUI application that interfaces with an Arduino-based Shenk material testing machine over serial. It reads force and displacement data in real-time, displays a live force-displacement plot, and exports results to `.xlsx`.

## Commands

### Run the application
```bash
source .venv/Scripts/activate  # or .venv\Scripts\activate.bat on Windows cmd
python shenk-acquisition-v4.py
```

### Install dependencies
```bash
pip install -r requirements.txt
```

### Build standalone executable (PyInstaller)
```bash
pyinstaller ShenkController.spec
# Output: standAlone/Shenk-app.exe
```

## Architecture

The entire application is a single file (`shenk-acquisition-v4.py`) with one class `App`.

**Serial communication**: Connects to an Arduino at 9600 baud. Each read cycle (`read_data`, called every 50ms via `root.after`) reads two consecutive `readline()` calls — first line is position (mm), second is force (kg). Auto-reconnect logic retries every 300ms if the port is unavailable.

**Data flow**:
- Raw values from serial → zero-offset correction (`zeropos`, `zeroforce`) → stored in `arrayp` (position) and `arrayf` (force)
- `Zero Pos` / `Zero Forza` buttons capture the current reading as the zero reference
- Test start clears both arrays and resets the plot; test stop freezes the lamp indicator

**Plot**: Matplotlib `FigureCanvasTkAgg` embedded in the tkinter grid. Redrawn on every serial read while the test is running (`self.stop == False`). Resizes dynamically via `<Configure>` event binding.

**Export**: Saves to `.xlsx` using `openpyxl` — writes metadata (datetime, speed) in columns A/C and data in columns E/F, then saves a PNG of the plot, embeds it at anchor `H1`, and deletes the temporary PNG.

**Old versions**: Previous iterations are in `old/` and `original_file/` — they are not used by the current build.
