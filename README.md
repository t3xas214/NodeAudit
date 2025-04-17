# NodeAudit Excel Automation Tool

A PySide6-based desktop application for managing and automating Excel data entry with PRISM integration.

## Features

- Configurable data entry for PIDs, Nodes, Scopes, and Magellan data
- Browser selection (Edge/Chrome) for PRISM integration
- Dark/Light mode toggle for better visibility
- Improved UI layout with better organization
- Progress tracking and error handling
- Read-only Excel file viewing
- Automatic PID and Node field population from PRISM
- Smart layout with properly spaced controls

## Requirements

- Python 3.7+
- Microsoft Edge or Google Chrome browser
- Required Python packages (see requirements.txt)
- PySide6 and PySide6-WebEngine for browser integration
- Active network connection for PRISM access

## Installation

1. Clone the repository:
```bash
git clone https://github.com/t3xas214/NodeAudit.git
cd NodeAudit
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Getting Started

1. Run the application:
```bash
python main.py
```

2. Initial Setup:
   - Select your preferred browser (Edge/Chrome) from the dropdown
   - The application will remember your browser preference
   - Toggle dark/light mode based on your preference

3. Basic Workflow:
   - Load an Excel file using the "Load Excel File" button
   - Enter data in the PID, Node, Scope, and Magellan fields
   - Use the browser integration to fetch data from PRISM
   - Save your changes using "Save & Next"
   - Navigate through rows using "Back" or "Go to Row"

## Browser Configuration

- Edge Configuration:
  1. Ensure Microsoft Edge is installed
  2. Select "Edge" from the browser dropdown
  3. The app will automatically launch Edge for PRISM access

- Chrome Configuration:
  1. Ensure Google Chrome is installed
  2. Select "Chrome" from the browser dropdown
  3. The app will automatically launch Chrome for PRISM access

Note: Make sure you're connected to your work network or VPN for PRISM access.

## Troubleshooting

Common issues and solutions:

1. Browser Not Opening:
   - Verify browser installation
   - Check network connectivity
   - Ensure you're connected to work VPN if accessing remotely

2. Excel Issues:
   - Make sure Excel file isn't open elsewhere
   - Check file permissions
   - Verify Excel file format (.xlsx or .xls)

3. PRISM Integration:
   - Confirm network/VPN connection
   - Verify PRISM URL accessibility
   - Check browser compatibility

## Error Handling

The application includes comprehensive error handling and will display appropriate error messages when:
- Excel files fail to load
- Data cannot be saved
- PRISM data extraction encounters issues
- Browser operations fail

## Contributing

Feel free to submit issues and enhancement requests. Pull requests are welcome:

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a new Pull Request
