# Smart Library Analytics Hub
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-1.0.0-green.svg)]()
[![Build Status](https://img.shields.io/badge/build-passing-brightgreen.svg)]()
[![Downloads](https://img.shields.io/npm/dm/package-name.svg)]()


## Overview
Smart Library Analytics Hub is a desktop application designed to manage and monitor library attendance using QR code technology. It provides tools for user management, real-time scanning, and analytical insights into library usage patterns.

The system is tailored for educational institutions and supports Khmer script along with grade-level reporting.

## Features

### QR Code Scanning
Real-time check-in and check-out using a webcam.

### Live Dashboard
Displays key metrics including total users, daily check-ins, and gender-based attendance.

### Advanced Analytics
Interactive charts for monthly trends and grade distribution.

### User Management
Add, edit, delete, and search users efficiently.

### Reporting
Export attendance data to PDF and Excel formats.

### QR Card Generation
Generate QR identification cards individually or in batches.

### Data Security
Uses a local SQLite database with backup support.

### Manual Input
Allows manual entry of QR codes when camera hardware is unavailable.

## Tech Stack

- Language: Python 3.x  
- GUI Framework: PyQt5  
- Computer Vision: OpenCV (cv2)  
- Data Processing: Pandas, NumPy  
- Database: SQLite3  
- PDF Generation: ReportLab  
- Excel Export: Openpyxl  
- Imaging: Pillow (PIL), qrcode  

## Setup and Installation

### Prerequisites

- Python 3.8 or higher  
- Webcam (for scanning functionality)  

### Installation

#### Clone the repository
```bash
git clone https://github.com/username/library_attendance.git
cd library_attendance
```

#### Install dependencies
```bash
pip install PyQt5 opencv-python pandas openpyxl reportlab qrcode pillow numpy
```

Note: sqlite3 is included with Python and does not need to be installed separately.

### Asset Requirements

Ensure the following files are present:

```
asset/
├── sound/
│   ├── checkin.mp3
│   ├── checkout.mp3
│   └── invaild.mp3
└── logo/
    └── logo.png
```

### Run the application
```bash
python main.py
```

## Usage

### Scanner Tab
Used for check-in and check-out. Click "Start Camera" to begin scanning QR codes.

### Dashboard Tab
Provides a summary of daily activity. Supports importing users from Excel and generating reports.

### Analytics Tab
Displays long-term trends and statistics by grade and time period.

### Settings
Configure camera index, study years, and file paths for saved reports.

## Configuration

The application generates a settings.json file on first launch.

Example configuration:

```json
{
  "sound_enabled": true,
  "camera_index": 0,
  "qr_save_dir": "path/to/output"
}
```

## Contribution Guidelines

1. Fork the repository  
2. Create a feature branch:
```bash
git checkout -b feature/NewFeature
```
3. Commit your changes:
```bash
git commit -m "Add NewFeature"
```
4. Push to the branch:
```bash
git push origin feature/NewFeature
```
5. Open a Pull Request  

## License

This project is distributed under the MIT License. See the LICENSE file for details.

## Contact

Developer: Sao Moni  
Email: keomony074@gmail.com  
Institution: Samdech Ov Samdech Mae High School  
