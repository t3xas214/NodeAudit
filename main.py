import sys
import json
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QFileDialog,
    QLabel, QTextEdit, QComboBox, QGridLayout, QLineEdit, QMessageBox
)
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtCore import Qt, QUrl, pyqtSlot, QObject, pyqtSignal
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtWebChannel import QWebChannel
from openpyxl import load_workbook
import traceback

# ==== PRISM Web Browser Integration ====
class Bridge(QObject):
    def __init__(self, status_label, parent=None):
        super().__init__(parent)
        self.status_label = status_label
        self.design_status = None

    @pyqtSlot(str)
    def receiveStatus(self, status):
        self.status_label.setText(f"Design Status: {status}")
        # Store the received status
        self.design_status = status
        
        # Emit signal when design approved status is detected
        # Check if parent exists and if it has the design_approved_detected signal
        if self.parent() and hasattr(self.parent(), 'design_approved_detected'):
            if status == "✅ Design Approved":
                self.parent().design_approved_detected.emit()

class PrismBrowser(QMainWindow):
    # Add signal for design approval detection
    design_approved_detected = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PRISM Viewer")
        self.setGeometry(200, 200, 1000, 800)
        
        # Create a central widget to hold our layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        self.status_label = QLabel("Status: Not Pulled")
        self.browser = QWebEngineView()
        self.browser.load(QUrl("https://prism.charter.com"))

        self.pull_button = QPushButton("Pull Design Status")
        self.pull_button.clicked.connect(self.pull_status)

        layout.addWidget(self.browser)
        layout.addWidget(self.pull_button)
        layout.addWidget(self.status_label)
        
        # Set window flags to make it a proper standalone window
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)

        self.channel = QWebChannel()
        self.bridge = Bridge(self.status_label, self)
        self.channel.registerObject("bridge", self.bridge)
        self.browser.page().setWebChannel(self.channel)

    def pull_status(self):
        js_code = """
        (function() {
            let cell = [...document.querySelectorAll('td')].find(td => td.innerText.includes('Design Status'));
            if (cell && cell.nextElementSibling) {
                let status = cell.nextElementSibling.innerText.trim();
                if (status.toLowerCase().includes("design approved")) {
                    bridge.receiveStatus("✅ Design Approved");
                } else {
                    bridge.receiveStatus("⚠️ Not Approved – Use dropdown");
                }
            } else {
                bridge.receiveStatus("Status Not Found");
            }
        })();
        """
        self.browser.page().runJavaScript(js_code)

# ==== Your Original App ====
class ExcelAutomationApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Automation App")
        self.setGeometry(100, 100, 800, 500)

        # Layout
        layout = QGridLayout()

        # Load Excel Button
        self.load_button = QPushButton("Load Excel File")
        self.load_button.clicked.connect(self.load_excel)
        layout.addWidget(self.load_button, 0, 0)

        # PID Inputs
        self.pid_inputs = [QLineEdit() for _ in range(4)]
        for i, pid_input in enumerate(self.pid_inputs):
            pid_input.setPlaceholderText(f"PID {i + 1}")
            layout.addWidget(pid_input, 1, i)

        # Node Inputs
        self.node_inputs = [QLineEdit() for _ in range(4)]
        for i, node_input in enumerate(self.node_inputs):
            node_input.setPlaceholderText(f"Node {i + 1}")
            layout.addWidget(node_input, 2, i)

        # Scope Inputs
        self.scope_inputs = [QLineEdit() for _ in range(4)]
        for i, scope_input in enumerate(self.scope_inputs):
            scope_input.setPlaceholderText(f"Scope {i + 1}")
            layout.addWidget(scope_input, 3, i)

        # Magellan Inputs
        self.magellan_inputs = [QLineEdit() for _ in range(4)]
        for i, magellan_input in enumerate(self.magellan_inputs):
            magellan_input.setPlaceholderText(f"Magellan {i + 1}")
            layout.addWidget(magellan_input, 4, i)

        # Config Dropdown
        self.config_dropdown = QComboBox()
        self.config_dropdown.addItems(["1x1", "2x2", "4x4", "N/A"])
        layout.addWidget(self.config_dropdown, 5, 0)

        # Build State Dropdown
        self.build_state_dropdown = QComboBox()
        self.build_state_dropdown.setEditable(True)
        self.build_state_dropdown.addItems(["In Design", "In Progress", "Does Not Exist", "PRO-I", "Design Approved"])
        layout.addWidget(self.build_state_dropdown, 5, 1)

        # Save & Next Button
        self.save_next_button = QPushButton("Save & Next")
        self.save_next_button.clicked.connect(self.save_next_action)
        layout.addWidget(self.save_next_button, 6, 0)

        # Back Button
        self.back_button = QPushButton("Back")
        self.back_button.clicked.connect(self.load_previous_row)
        layout.addWidget(self.back_button, 6, 1)
        
        self.current_row = None

        # Status Label
        self.status_label = QLabel("No file loaded.")
        layout.addWidget(self.status_label, 7, 0, 1, 3)

        self.last_node_label = QLabel("Last Node: N/A")
        layout.addWidget(self.last_node_label, 8, 0, 1, 3)

        # Row Input and Go Button
        self.row_input = QLineEdit()
        self.row_input.setPlaceholderText("Enter row #")
        layout.addWidget(self.row_input, 9, 0)

        self.go_button = QPushButton("Go to Row")
        self.go_button.clicked.connect(self.load_specific_row)
        layout.addWidget(self.go_button, 9, 1)

        self.open_excel_button = QPushButton("Open in Excel (Read-Only)")
        self.open_excel_button.clicked.connect(self.open_excel_readonly)
        layout.addWidget(self.open_excel_button, 10, 0)
        
        self.current_row_label = QLabel("Current Row: N/A")
        layout.addWidget(self.current_row_label, 11, 0, 1, 3)

        # Toggle Dark Mode Button
        self.toggle_dark_mode_button = QPushButton("Toggle Dark Mode")
        self.toggle_dark_mode_button.clicked.connect(self.toggle_dark_mode)
        layout.addWidget(self.toggle_dark_mode_button, 6, 2)

        # Main Widget
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Placeholder for Excel file path
        self.excel_file_path = None

        # Set up theme after all UI elements are created
        self.setup_theme()

    def open_prism_browser(self):
        self.browser_window = PrismBrowser()  # Without parent parameter
        # Connect the signal to a slot that will update the build state dropdown
        self.browser_window.design_approved_detected.connect(self.set_design_approved)
        self.browser_window.show()

    def set_design_approved(self):
        """Set the build state dropdown to 'Design Approved' when detected in PRISM."""
        # First check if "Design Approved" is already in the dropdown list
        index = self.build_state_dropdown.findText("Design Approved")
        if index == -1:  # If not found, add it
            self.build_state_dropdown.addItem("Design Approved")
            index = self.build_state_dropdown.findText("Design Approved")
        
        # Set the dropdown to "Design Approved"
        self.build_state_dropdown.setCurrentIndex(index)
        
        # Show notification to the user
        self.status_label.setText("Status updated: Design Approved detected and set in dropdown")

    def setup_theme(self, dark_mode=False):
        app = QApplication.instance()
        palette = QPalette()
        
        # Define neon green color
        neon_green = "#39FF14"
        neon_green_hover = "#45FF28"
        neon_green_pressed = "#32E011"

        if dark_mode:
            # Dark mode settings
            palette.setColor(QPalette.Window, QColor(53, 53, 53))
            palette.setColor(QPalette.WindowText, Qt.white)
            palette.setColor(QPalette.Base, QColor(35, 35, 35))
            palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
            palette.setColor(QPalette.ToolTipBase, QColor(25, 25, 25))
            palette.setColor(QPalette.ToolTipText, Qt.white)
            palette.setColor(QPalette.Text, Qt.white)
            palette.setColor(QPalette.Button, QColor(53, 53, 53))
            palette.setColor(QPalette.ButtonText, Qt.white)
            palette.setColor(QPalette.BrightText, Qt.red)
            palette.setColor(QPalette.Link, QColor(42, 130, 218))
            palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
            palette.setColor(QPalette.HighlightedText, QColor(35, 35, 35))
            
            # Set the dark style sheet with neon green accents
            self.setStyleSheet(f"""
                QMainWindow {{
                    background-color: #353535;
                }}
                QWidget {{
                    background-color: #353535;
                    color: white;
                }}
                QComboBox {{
                    background-color: #353535;
                    color: white;
                    border: 1px solid {neon_green};
                    padding: 5px;
                    border-radius: 3px;
                }}
                QComboBox:drop-down {{
                    border: 1px solid {neon_green};
                }}
                QComboBox:down-arrow {{
                    width: 15px;
                    height: 15px;
                }}
                QLineEdit {{
                    background-color: #353535;
                    color: white;
                    border: 1px solid {neon_green};
                    padding: 5px;
                    border-radius: 3px;
                }}
                QPushButton {{
                    background-color: #353535;
                    color: {neon_green};
                    border: 2px solid {neon_green};
                    padding: 5px;
                    border-radius: 3px;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    background-color: {neon_green};
                    color: #353535;
                }}
                QPushButton:pressed {{
                    background-color: {neon_green_pressed};
                    color: #353535;
                }}
                QLabel {{
                    color: white;
                }}
                QMessageBox {{
                    background-color: #353535;
                }}
                QMessageBox QLabel {{
                    color: white;
                }}
                QMessageBox QPushButton {{
                    background-color: #353535;
                    color: {neon_green};
                    border: 2px solid {neon_green};
                }}
            """)
            self.status_label.setText("Dark Mode Enabled")
        else:
            # Light mode settings
            palette = app.style().standardPalette()
            # Updated light mode stylesheet with filled neon green buttons and black text
            self.setStyleSheet(f"""
                QMainWindow {{
                    background-color: white;
                }}
                QWidget {{
                    background-color: white;
                    color: black;
                }}
                QComboBox {{
                    background-color: white;
                    color: black;
                    border: 1px solid {neon_green};
                    padding: 5px;
                    border-radius: 3px;
                }}
                QLineEdit {{
                    background-color: white;
                    color: black;
                    border: 1px solid {neon_green};
                    padding: 5px;
                    border-radius: 3px;
                }}
                QPushButton {{
                    background-color: {neon_green};
                    color: black;
                    border: none;
                    padding: 6px;
                    border-radius: 3px;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    background-color: {neon_green_hover};
                    color: black;
                }}
                QPushButton:pressed {{
                    background-color: {neon_green_pressed};
                    color: black;
                }}
                QLabel {{
                    color: black;
                }}
            """)
            self.status_label.setText("Light Mode Enabled")

        app.setPalette(palette)
        self.is_dark_mode = dark_mode  # Store the current theme state

    def load_excel(self):
        """Load the Excel file and read its contents."""
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)"
            )
            if file_path:
                self.excel_file_path = file_path
                self.status_label.setText(f"File loaded: {os.path.basename(file_path)}")
                # Start with the first row of data (assuming row 2 if row 1 is headers)
                self.current_row = 2
                self.load_row_data()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load Excel file: {str(e)}")
            traceback.print_exc()

    def load_row_data(self):
        """Load data from the current row in the Excel file."""
        if not self.excel_file_path or not self.current_row:
            return

        wb = None
        try:
            wb = load_workbook(self.excel_file_path, read_only=True)
            ws = wb.active
            
            # Update current row label
            self.current_row_label.setText(f"Current Row: {self.current_row}")
            
            # Get row data
            row_data = [cell.value for cell in ws[self.current_row]]
            
            # Check if we have valid row data
            if not row_data:
                QMessageBox.information(self, "Info", f"No data found at row {self.current_row}.")
                return
            
            if all(val is None for val in row_data):
                QMessageBox.information(self, "Info", f"Row {self.current_row} is empty.")
                return
            
            # Assuming columns are in order: PID, Node, Scope, Magellan, Config, Build State
            # Excel columns are 1-indexed but list indices are 0-indexed
            pid_col, node_col, scope_col, magellan_col = 0, 1, 2, 3  # These are list indices (0-based)
            config_col, build_state_col = 4, 5
            
            # Clear all input fields first
            for input_field in self.pid_inputs + self.node_inputs + self.scope_inputs + self.magellan_inputs:
                input_field.clear()
            
            # PID inputs
            if pid_col < len(row_data) and row_data[pid_col]:
                pid_values = str(row_data[pid_col]).split(',')
                for i, pid_input in enumerate(self.pid_inputs):
                    pid_input.setText(pid_values[i].strip() if i < len(pid_values) else "")
            
            # Node inputs
            if node_col < len(row_data) and row_data[node_col]:
                node_values = str(row_data[node_col]).split(',')
                for i, node_input in enumerate(self.node_inputs):
                    node_input.setText(node_values[i].strip() if i < len(node_values) else "")
                
                # Update last node label
                if node_values:
                    self.last_node_label.setText(f"Last Node: {node_values[0].strip() if node_values else 'N/A'}")
            
            # Scope inputs
            if scope_col < len(row_data) and row_data[scope_col]:
                scope_values = str(row_data[scope_col]).split(',')
                for i, scope_input in enumerate(self.scope_inputs):
                    scope_input.setText(scope_values[i].strip() if i < len(scope_values) else "")
            
            # Magellan inputs
            if magellan_col < len(row_data) and row_data[magellan_col]:
                magellan_values = str(row_data[magellan_col]).split(',')
                for i, magellan_input in enumerate(self.magellan_inputs):
                    magellan_input.setText(magellan_values[i].strip() if i < len(magellan_values) else "")
            
            # Config dropdown
            if config_col < len(row_data) and row_data[config_col]:
                config_value = str(row_data[config_col])
                index = self.config_dropdown.findText(config_value)
                if index >= 0:
                    self.config_dropdown.setCurrentIndex(index)
            
            # Build state dropdown
            if build_state_col < len(row_data) and row_data[build_state_col]:
                build_state = str(row_data[build_state_col])
                index = self.build_state_dropdown.findText(build_state)
                if index >= 0:
                    self.build_state_dropdown.setCurrentIndex(index)
                else:
                    self.build_state_dropdown.setCurrentText(build_state)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load row data: {str(e)}")
            traceback.print_exc()
        finally:
            if wb:
                wb.close()

    def save_next_action(self):
        """Save current row data and move to the next row."""
        if not self.excel_file_path:
            QMessageBox.warning(self, "Warning", "No file loaded.")
            return
            
        try:
            self.save_current_row()
            self.current_row += 1
            self.load_row_data()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save and move to next row: {str(e)}")
            traceback.print_exc()

    def save_current_row(self):
        """Save the current row data to the Excel file."""
        if not self.excel_file_path or not self.current_row:
            return
            
        wb = None
        try:
            # Load the workbook (not read-only this time)
            wb = load_workbook(self.excel_file_path)
            ws = wb.active
            
            # Collect data from input fields
            # PIDs
            pid_values = [pid_input.text() for pid_input in self.pid_inputs if pid_input.text()]
            pid_str = ", ".join(pid_values)
            
            # Nodes
            node_values = [node_input.text() for node_input in self.node_inputs if node_input.text()]
            node_str = ", ".join(node_values)
            
            # Scopes
            scope_values = [scope_input.text() for scope_input in self.scope_inputs if scope_input.text()]
            scope_str = ", ".join(scope_values)
            
            # Magellan
            magellan_values = [magellan_input.text() for magellan_input in self.magellan_inputs if magellan_input.text()]
            magellan_str = ", ".join(magellan_values)
            
            # Config
            config_value = self.config_dropdown.currentText()
            
            # Build state
            build_state = self.build_state_dropdown.currentText()
            
            # Update the cells in the row
            # Using 1-indexed column numbers for Excel's API
            # These should match the column indices in load_row_data (accounting for 0 vs 1 indexing)
            pid_col, node_col, scope_col, magellan_col = 1, 2, 3, 4  # Excel columns are 1-indexed
            config_col, build_state_col = 5, 6
            
            ws.cell(row=self.current_row, column=pid_col, value=pid_str)
            ws.cell(row=self.current_row, column=node_col, value=node_str)
            ws.cell(row=self.current_row, column=scope_col, value=scope_str)
            ws.cell(row=self.current_row, column=magellan_col, value=magellan_str)
            ws.cell(row=self.current_row, column=config_col, value=config_value)
            ws.cell(row=self.current_row, column=build_state_col, value=build_state)
            
            # Save the workbook
            wb.save(self.excel_file_path)
            self.status_label.setText(f"Saved row {self.current_row}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save row data: {str(e)}")
            traceback.print_exc()
        finally:
            if wb:
                wb.close()

    def load_previous_row(self):
        """Load the previous row data."""
        if not self.excel_file_path or not self.current_row or self.current_row <= 2:
            QMessageBox.information(self, "Info", "Already at the first row.")
            return
            
        try:
            self.current_row -= 1
            self.load_row_data()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load previous row: {str(e)}")
            traceback.print_exc()

    def load_specific_row(self):
        """Load a specific row based on user input."""
        if not self.excel_file_path:
            QMessageBox.warning(self, "Warning", "No file loaded.")
            return
            
        try:
            row_text = self.row_input.text()
            if not row_text:
                return
                
            row_num = int(row_text)
            if row_num < 2:
                QMessageBox.warning(self, "Warning", "Row number must be at least 2.")
                return
                
            self.current_row = row_num
            self.load_row_data()
            self.row_input.clear()
        except ValueError:
            QMessageBox.warning(self, "Warning", "Please enter a valid row number.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load specified row: {str(e)}")
            traceback.print_exc()

    def open_excel_readonly(self):
        """Open the current Excel file in the default application (read-only)."""
        if not self.excel_file_path:
            QMessageBox.warning(self, "Warning", "No file loaded.")
            return
            
        try:
            if sys.platform == 'win32':
                os.startfile(self.excel_file_path)
            elif sys.platform == 'darwin':  # macOS
                os.system(f'open "{self.excel_file_path}"')
            else:  # Linux
                os.system(f'xdg-open "{self.excel_file_path}"')
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open Excel file: {str(e)}")
            traceback.print_exc()

    def toggle_dark_mode(self):
        """Toggle between dark and light mode."""
        current_palette = QApplication.instance().palette()
        is_dark_mode = current_palette.color(QPalette.Window).lightness() < 128
        self.setup_theme(not is_dark_mode)  # Toggle the theme

# Run the app
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelAutomationApp()
    window.show()
    sys.exit(app.exec_())
