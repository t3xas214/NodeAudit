# NodeAudit Excel Automation Tool

A PyQt5-based desktop application for managing and automating Excel data entry and PRISM data extraction.

## Features

- Excel file loading and manipulation
- Automated PRISM data extraction
- Support for multiple browsers (Edge and Chrome)
- Configurable data entry for PIDs, Nodes, Scopes, and Magellan data
- Progress tracking and error handling
- Read-only Excel file viewing

## Requirements

- Python 3.7+
- Microsoft Edge or Google Chrome browser
- Required Python packages (see requirements.txt)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/NodeAudit.git
cd NodeAudit
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the application:
```bash
python main.py
```

2. Use the interface to:
   - Load Excel files
   - Enter data in the provided fields
   - Extract PRISM data automatically
   - Navigate through rows
   - Save changes

## Configuration

- Browser settings are automatically saved in `settings.json`
- Supports both Edge and Chrome browsers for PRISM data extraction

## Error Handling

The application includes comprehensive error handling and will display appropriate error messages when:
- Excel files fail to load
- Data cannot be saved
- PRISM data extraction encounters issues
- Browser operations fail

## Contributing

Feel free to submit issues and enhancement requests. 