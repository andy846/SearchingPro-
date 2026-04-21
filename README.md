# SearchingPro

A powerful desktop file search application built with PyQt5. Supports filename, content, regex, fuzzy, and boolean expression search across multiple file formats.

## Features

- **Multi-mode search**: Filename, content, regex, and fuzzy matching
- **Boolean expression search**: `AND`, `OR`, `NOT`, `NEAR` operators
- **File content indexing**: Index and search inside PDF, DOCX, XLSX, PPTX, TXT, and more
- **Realtime file watch**: Auto-detect filesystem changes and update index
- **Advanced filters**: File type, size range, date range, path keywords, exclusion rules
- **Pagination**: Configurable page size for large result sets
- **Search history & templates**: Save and reuse frequent search patterns
- **Statistics dashboard**: File type distribution, size charts, modification time trends (requires matplotlib)
- **Dark/Light theme**: Toggle between themes
- **Internationalization**: English, Simplified Chinese, Traditional Chinese
- **Cross-platform**: macOS (Intel/Apple Silicon), Windows, Linux
- **Export results**: CSV, JSON, Excel
- **File preview & context menu**: Open file, open folder, copy path, batch operations

## Requirements

- Python 3.8+
- PyQt5 >= 5.15.0
- psutil >= 5.8.0
- watchdog >= 2.0.0
- PyPDF2 >= 3.0.0 (for PDF content indexing)
- python-docx >= 0.8.11 (for DOCX content indexing)
- openpyxl >= 3.0.0 (for XLSX content indexing)
- python-pptx >= 0.6.21 (for PPTX content indexing)
- matplotlib >= 3.3.0 (optional, for statistics charts)

## Installation

```bash
pip install -r requirements.txt
# Optional: for statistics charts
pip install matplotlib
```

## Usage

```bash
python SearchingPro.py
```

## Building Standalone Binaries

### macOS

```bash
pyinstaller SearchingPro_mac.spec
```

Produces `dist/SearchingPro.app`.

### Windows

```bash
pyinstaller SearchingPro.spec
```

Produces `dist/SearchingPro.exe`.

## Project Structure

```
SearchingPro/
├── SearchingPro.py          # Main application (single-file)
├── SearchingPro.spec        # PyInstaller spec (Windows)
├── SearchingPro_mac.spec    # PyInstaller spec (macOS)
├── requirements.txt         # Python dependencies
├── i18n/
│   ├── en.json              # English translations
│   ├── zh-CN.json           # Simplified Chinese translations
│   └── zh-TW.json           # Traditional Chinese translations
├── app_icon.ico             # Windows icon
├── SearchingPro.icns        # macOS app icon
└── README.md
```

## License

MIT License -- see [LICENSE](LICENSE).
