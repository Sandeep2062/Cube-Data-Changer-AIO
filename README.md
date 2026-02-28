<div align="center">

# ◆ Cube Data Changer AIO

**All-in-One tool for generating and processing concrete & mortar cube test data**

[![Build](https://github.com/Sandeep2062/Cube-Data-Changer-AIO/actions/workflows/build.yml/badge.svg)](https://github.com/Sandeep2062/Cube-Data-Changer-AIO/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.8+](https://img.shields.io/badge/Python-3.8+-3776AB.svg)](https://python.org)

</div>

---

## What is this?

**Cube Data Changer AIO** merges two separate tools into one seamless workflow:

| Before (2 separate tools) | After (AIO) |
|---|---|
| 1. Run **Cube Data Generator** → creates Excel files | 1. Select grades + office template |
| 2. Open **Cube Data Processor** → load those Excel files | 2. Click **Start** → done ✅ |
| 3. Configure, process, save | No intermediate files needed |

### Features

- **Auto-Generate** concrete (M10–M45) and mortar (1:4, 1:6) test data in-memory
- **Auto-Process** — generated data is written directly into office template sheets
- **Calendar Date Processing** — 7-day / 28-day test dates from calendar file
- **Modern Dark UI** built with CustomTkinter
- **Legacy Mode** — still supports loading pre-made grade Excel files
- **Cross-Platform** settings (JSON-based, no Windows Registry dependency)
- **One-Click EXE** build via GitHub Actions

---

## Processing Modes

| Mode | Description |
|---|---|
| ⚡ **Auto Generate + Date** | Generate data + apply calendar dates (recommended) |
| 🔄 **Auto Generate Only** | Generate and apply grade data, skip dates |
| 📅 **Date Only** | Only apply calendar dates to existing sheets |
| 📁 **Files + Date (Legacy)** | Use existing grade Excel files + dates |
| 📁 **Files Only (Legacy)** | Use existing grade Excel files only |

---

## Supported Grades

### Concrete Mixes
| Grade | Weight Range (kg) | 7-Day Strength (kN) | 28-Day Strength (kN) |
|---|---|---|---|
| M10 | 8.100 – 8.300 | 214.00 – 267.40 | 320.10 – 365.50 |
| M15 | 8.100 – 8.300 | 290.10 – 320.50 | 433.10 – 480.10 |
| M20 | 8.100 – 8.300 | 366.10 – 410.10 | 547.10 – 590.10 |
| M25 | 8.180 – 8.350 | 442.10 – 490.10 | 660.10 – 710.10 |
| M30 | 8.100 – 8.350 | 518.10 – 560.10 | 770.10 – 812.10 |
| M35 | 8.100 – 8.350 | 595.10 – 632.80 | 880.90 – 925.10 |
| M40 | 8.100 – 8.350 | 669.10 – 728.10 | 995.10 – 1038.10 |
| M45 | 8.200 – 8.400 | 735.10 – 788.10 | 1105.35 – 1150.10 |

### Mortar Mixes
| Type | Weight Range (kg) | 7-Day Strength (kN) | 28-Day Strength (kN) |
|---|---|---|---|
| 1:4 | 0.800 – 0.835 | 25.20 – 33.90 | 40.60 – 50.10 |
| 1:6 | 0.800 – 0.835 | 15.20 – 25.00 | 25.20 – 33.90 |

---

## Quick Start

### Run from Source
```bash
# Clone
git clone https://github.com/Sandeep2062/Cube-Data-Changer-AIO.git
cd Cube-Data-Changer-AIO

# Install dependencies
pip install -r requirements.txt

# Run
python app.py
```

### Download EXE
Go to [Releases](https://github.com/Sandeep2062/Cube-Data-Changer-AIO/releases) and download the latest `.exe`.

---

## How It Works

1. **Select grades** (M10–M45, Mortar 1:4/1:6) in the sidebar
2. **Browse** your office template Excel file
3. **Browse** calendar file (optional, for date processing)
4. **Select** output folder
5. Click **▶ START PROCESSING**

The app will:
- Generate random but realistic weight and strength values for each selected grade
- Match sheets in your office template by checking cell **B12** for the grade name
- Write weights to **row 25, columns C–H**
- Write 7-day + 28-day strengths to **row 27, columns C–H**
- Optionally write test dates from the calendar file

---

## Project Structure

```
Cube-Data-Changer-AIO/
├── app.py              # Main GUI application
├── generator.py        # Data generation module
├── processor.py        # Data processing module
├── settings.py         # Cross-platform settings (JSON)
├── requirements.txt    # Python dependencies
├── icon.ico            # Application icon
├── logo.png            # Sidebar logo
├── LICENSE             # MIT License
└── .github/
    └── workflows/
        └── build.yml   # GitHub Actions: build EXE + release
```

---

## Building EXE

### Automatic (GitHub Actions)
Push a version tag to trigger the build:
```bash
git tag v1.0.0
git push origin v1.0.0
```

### Manual (Local)
```bash
pip install pyinstaller
pyinstaller --onefile --noconsole \
  --name "Cube-Data-Changer-AIO" \
  --icon="icon.ico" \
  --add-data "logo.png:." \
  --add-data "icon.ico:." \
  --collect-all customtkinter \
  --collect-all openpyxl \
  --hidden-import=PIL \
  --hidden-import=numpy \
  app.py
```

---

## Credits

Merged from:
- [Cube-Data-Generator](https://github.com/Sandeep2062/Cube-Data-Generator) — data generation logic
- [Cube-Data-Processor](https://github.com/Sandeep2062/Cube-Data-Processor) — processing logic & UI base

---

<div align="center">

**Developer:** [Sandeep](https://github.com/Sandeep2062) · © 2026

</div>
