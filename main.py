import sys
import json
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QFileDialog, QLabel, QTextEdit, QComboBox, QGridLayout, QLineEdit
from PyQt5.QtWidgets import QMessageBox
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
import time

class ExcelAutomationApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Automation App")
        self.setGeometry(100, 100, 600, 400)
        self.settings_file = "settings.json"
        selected_browser = self.load_browser_setting()

        # Layout
        layout = QGridLayout()

        # Load Excel Button
        self.load_button = QPushButton("Load Excel File")
        self.load_button.clicked.connect(self.load_excel)
        layout.addWidget(self.load_button, 0, 0)
        self.browser_dropdown = QComboBox()
        self.browser_dropdown.addItems(["Edge", "Chrome"])
        self.browser_dropdown.setCurrentText(self.load_browser_setting())
        layout.addWidget(self.browser_dropdown, 0, 1)

        # PID Inputs
        self.pid_inputs = [QLineEdit() for _ in range(4)]
        for i, pid_input in enumerate(self.pid_inputs):
            pid_input.setPlaceholderText(f"PID {i + 1}")
            layout.addWidget(pid_input, 1, i)

        # Scope Inputs
        self.scope_inputs = [QLineEdit() for _ in range(4)]
        for i, scope_input in enumerate(self.scope_inputs):
            scope_input.setPlaceholderText(f"Scope {i + 1}")
            layout.addWidget(scope_input, 3, i)

        self.node_inputs = [QLineEdit() for _ in range(4)]
        for i, node_input in enumerate(self.node_inputs):
            node_input.setPlaceholderText(f"Node {i + 1}")
            layout.addWidget(node_input, 2, i)

        # Magellan Inputs
        self.magellan_inputs = [QLineEdit() for _ in range(4)]
        for i, magellan_input in enumerate(self.magellan_inputs):
            magellan_input.setPlaceholderText(f"Magellan {i + 1}")
            layout.addWidget(magellan_input, 4, i)

        # Config Dropdown
        self.config_dropdown = QComboBox()
        self.config_dropdown.addItems(["1x1", "2x2", "4x4"])
        layout.addWidget(self.config_dropdown, 5, 0)

        # Build State Dropdown
        self.build_state_dropdown = QComboBox()
        self.build_state_dropdown.setEditable(True)
        self.build_state_dropdown.addItems(["In Design", "In Progress", "Does Not Exist", "PRO I"])
        layout.addWidget(self.build_state_dropdown, 5, 1)


        # Save & Next Button
        self.save_next_button = QPushButton("Save & Next")
        self.save_next_button.clicked.connect(self.save_next_action)
        layout.addWidget(self.save_next_button, 6, 1)

        # Back Button
        self.back_button = QPushButton("Back")
        self.back_button.clicked.connect(self.load_previous_row)
        layout.addWidget(self.back_button, 6, 2)
        self.current_row = None

        # Status Label
        self.status_label = QLabel("No file loaded.")
        layout.addWidget(self.status_label, 7, 0, 1, 2)

        self.last_node_label = QLabel("Last Node: N/A")
        layout.addWidget(self.last_node_label, 8, 0, 1, 2)

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
        layout.addWidget(self.current_row_label, 11, 0, 1, 2)

        # Main Widget
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Placeholder for Excel file path
        self.excel_file_path = None

    def load_excel(self):
        # Open file dialog to select Excel file
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.excel_file_path = file_path
            self.status_label.setText(f"Loaded: {file_path}")
            QMessageBox.information(self, "Excel Loaded", f"Loaded: {os.path.basename(file_path)}")


    def save_next_action(self):
        if not self.excel_file_path:
            self.status_label.setText("No Excel file loaded.")
            return

        workbook = load_workbook(self.excel_file_path)
        sheet = workbook.active

        headers = [cell.value for cell in sheet[1]]

        # Decide which row to write to
        row = self.current_row if self.current_row else sheet.max_row + 1

        # Write PIDs
        for i, pid_input in enumerate(self.pid_inputs):
            col_name = f"PID {i+1}"
            if col_name in headers:
                col_idx = headers.index(col_name) + 1
                sheet.cell(row=row, column=col_idx).value = pid_input.text().strip().replace('\u00A0', '').replace('\u200B', '')

        # Write Scope
        for i, scope_input in enumerate(self.scope_inputs):
            col_name = f"SCOPE  {i+1}"
            if col_name in headers:
                col_idx = headers.index(col_name) + 1
                sheet.cell(row=row, column=col_idx).value = scope_input.text().strip().replace('\u00A0', '').replace('\u200B', '')

        # Write Magellan
        for i, mag_input in enumerate(self.magellan_inputs):
            col_name = f"MAGELLAN  {i+1}"
            if col_name in headers:
                col_idx = headers.index(col_name) + 1
                sheet.cell(row=row, column=col_idx).value = mag_input.text().strip().replace('\u00A0', '').replace('\u200B', '')

        # Write AOI NODE (Config)
        if "AOI NODE" in headers:
            col_idx = headers.index("AOI NODE") + 1
            sheet.cell(row=row, column=col_idx).value = self.config_dropdown.currentText().strip().replace('\u00A0', '').replace('\u200B', '')

        # Write NOTES (Build State)
        if "NOTES" in headers:
            col_idx = headers.index("NOTES") + 1
            sheet.cell(row=row, column=col_idx).value = self.build_state_dropdown.currentText().strip().replace('\u00A0', '').replace('\u200B', '')

        for i, node_input in enumerate(self.node_inputs):
            col_name = f"NODE {i+1}"
            if col_name in headers:
                col_idx = headers.index(col_name) + 1
                sheet.cell(row=row, column=col_idx).value = node_input.text().strip().replace('\u00A0', '').replace('\u200B', '')

        # Update labels
        self.last_node_label.setText(f"Last Node: {self.magellan_inputs[0].text()}")
        self.current_row_label.setText(f"Current Row: {row}")
        self.status_label.setText(f"Saved row {row}")

        workbook.save(self.excel_file_path)
        QMessageBox.information(self, "Saved", f"Row {row} saved successfully!")

        # Clear inputs
        for input_list in [self.pid_inputs, self.scope_inputs, self.magellan_inputs, self.node_inputs]:
            for field in input_list:
                field.clear()

        self.config_dropdown.setCurrentIndex(0)
        self.build_state_dropdown.setCurrentIndex(0)

        # Move to next row
        self.current_row = row + 1

    def load_previous_row(self):
        if not self.excel_file_path:
            self.status_label.setText("No Excel file loaded.")
            return

        workbook = load_workbook(self.excel_file_path)
        sheet = workbook.active

        if self.current_row is None:
            self.current_row = sheet.max_row
        else:
            self.current_row = max(2, self.current_row - 1)

        headers = [cell.value for cell in sheet[1]]

        # Load values back into input fields
        for i in range(4):
            pid_col = headers.index(f"PID {i+1}") + 1 if f"PID {i+1}" in headers else None
            if pid_col:
                self.pid_inputs[i].setText(str(sheet.cell(row=self.current_row, column=pid_col).value or ""))

            scope_col = headers.index(f"SCOPE {i+1}") + 1 if f"SCOPE {i+1}" in headers else None
            if scope_col:
                self.scope_inputs[i].setText(str(sheet.cell(row=self.current_row, column=scope_col).value or ""))

            mag_col = headers.index(f"MAGELLAN {i+1}") + 1 if f"MAGELLAN {i+1}" in headers else None
            if mag_col:
                self.magellan_inputs[i].setText(str(sheet.cell(row=self.current_row, column=mag_col).value or ""))

            node_col = headers.index(f"NODE {i+1}") + 1 if f"NODE {i+1}" in headers else None
            if node_col:
                self.node_inputs[i].setText(str(sheet.cell(row=self.current_row, column=node_col).value or ""))

        config_col = headers.index("CONFIG") + 1 if "CONFIG" in headers else None
        if config_col:
            config_value = sheet.cell(row=self.current_row, column=config_col).value
            if config_value in [self.config_dropdown.itemText(i) for i in range(self.config_dropdown.count())]:
                self.config_dropdown.setCurrentText(config_value)

        state_col = headers.index("BUILD STATE") + 1 if "BUILD STATE" in headers else None
        if state_col:
            state_value = sheet.cell(row=self.current_row, column=state_col).value
            if state_value in [self.build_state_dropdown.itemText(i) for i in range(self.build_state_dropdown.count())]:
                self.build_state_dropdown.setCurrentText(state_value)

        self.status_label.setText(f"Loaded previous row: {self.current_row}")
        self.last_node_label.setText(f"Last Node: {self.magellan_inputs[0].text()}")
        self.current_row_label.setText(f"Current Row: {self.current_row}")

    def load_specific_row(self):
        if not self.excel_file_path:
            self.status_label.setText("No Excel file loaded.")
            return

        try:
            target_row = int(self.row_input.text().strip())
        except ValueError:
            self.status_label.setText("Invalid row number.")
            return

        workbook = load_workbook(self.excel_file_path)
        sheet = workbook.active

        if target_row == 1:
            QMessageBox.information(self, "Invalid Row", "Row 1 contains headers and cannot be edited.")
            self.status_label.setText("Row 1 contains headers and cannot be edited.")
            return
        elif target_row < 2 or target_row > sheet.max_row:
            self.status_label.setText("Row number out of range.")
            return

        self.current_row = target_row
        headers = [cell.value for cell in sheet[1]]

        for i in range(4):
            pid_col = headers.index(f"PID {i+1}") + 1 if f"PID {i+1}" in headers else None
            if pid_col:
                self.pid_inputs[i].setText(str(sheet.cell(row=self.current_row, column=pid_col).value or ""))

            scope_col = headers.index(f"SCOPE {i+1}") + 1 if f"SCOPE {i+1}" in headers else None
            if scope_col:
                self.scope_inputs[i].setText(str(sheet.cell(row=self.current_row, column=scope_col).value or ""))

            mag_col = headers.index(f"MAGELLAN {i+1}") + 1 if f"MAGELLAN {i+1}" in headers else None
            if mag_col:
                self.magellan_inputs[i].setText(str(sheet.cell(row=self.current_row, column=mag_col).value or ""))

            node_col = headers.index(f"NODE {i+1}") + 1 if f"NODE {i+1}" in headers else None
            if node_col:
                self.node_inputs[i].setText(str(sheet.cell(row=self.current_row, column=node_col).value or ""))

        config_col = headers.index("CONFIG") + 1 if "CONFIG" in headers else None
        if config_col:
            config_value = sheet.cell(row=self.current_row, column=config_col).value
            if config_value in [self.config_dropdown.itemText(i) for i in range(self.config_dropdown.count())]:
                self.config_dropdown.setCurrentText(config_value)

        state_col = headers.index("BUILD STATE") + 1 if "BUILD STATE" in headers else None
        if state_col:
            state_value = sheet.cell(row=self.current_row, column=state_col).value
            if state_value in [self.build_state_dropdown.itemText(i) for i in range(self.build_state_dropdown.count())]:
                self.build_state_dropdown.setCurrentText(state_value)

        self.status_label.setText(f"Loaded row: {self.current_row}")
        self.last_node_label.setText(f"Last Node: {self.magellan_inputs[0].text()}")
        self.current_row_label.setText(f"Current Row: {self.current_row}")

    def open_excel_readonly(self):
        if not self.excel_file_path:
            self.status_label.setText("No Excel file loaded.")
            return

        try:
            if sys.platform == "win32":
                # Use the /r switch with start to explicitly open Excel in read-only mode
                os.system(f'start "" "EXCEL.EXE" /r "{self.excel_file_path}"')
            elif sys.platform == "darwin":
                os.system(f'open -a "Microsoft Excel" "{self.excel_file_path}"')
            else:
                os.system(f'libreoffice --view "{self.excel_file_path}"')

            self.status_label.setText("Opened Excel in read-only mode.")
        except Exception as e:
            self.status_label.setText(f"Failed to open Excel: {e}")

    def extract_prism_data(self):
        if not self.excel_file_path:
            self.status_label.setText("Please load an Excel file first.")
            return

        try:
            selected_browser = self.browser_dropdown.currentText()
            if selected_browser == "Edge":
                edge_options = EdgeOptions()
                edge_options.add_argument("--start-maximized")
                service = EdgeService(EdgeChromiumDriverManager().install())
                driver = webdriver.Edge(service=service, options=edge_options)
            else:  # Chrome
                from selenium.webdriver.chrome.service import Service as ChromeService
                from selenium.webdriver.chrome.options import Options as ChromeOptions
                from webdriver_manager.chrome import ChromeDriverManager

                chrome_options = ChromeOptions()
                chrome_options.add_argument("--start-maximized")
                service = ChromeService(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=chrome_options)

            driver.get("https://prism.charter.com/")

            workbook = load_workbook(self.excel_file_path)
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]

                # Ensure headers for new columns
                headers = [cell.value for cell in sheet[1]]
                for i in range(1, 5):
                    status_col = f"DESIGN STATUS {i}"
                    if status_col not in headers:
                        sheet.cell(row=1, column=len(headers) + 1, value=status_col)
                        headers.append(status_col)

                for row in range(2, sheet.max_row + 1):
                    for i in range(1, 5):
                        pid_cell = sheet.cell(row=row, column=headers.index(f"PID {i}") + 1)
                        status_cell = sheet.cell(row=row, column=headers.index(f"DESIGN STATUS {i}") + 1)

                        pid = pid_cell.value
                        if not pid:
                            continue

                        try:
                            search_box = driver.find_element(By.ID, "searchForm:searchPrismId")
                            search_box.clear()
                            search_box.send_keys(str(int(pid)))
                            driver.find_element(By.ID, "searchForm:searchSubmit").click()
                            time.sleep(4)

                            driver.find_element(By.LINK_TEXT, "Design").click()
                            time.sleep(2)

                            design_status = driver.find_element(By.XPATH, "//td[contains(text(), 'Design Status')]/following-sibling::td").text.strip()
                            status_cell.value = design_status
                        except Exception as inner_e:
                            status_cell.value = f"Error: {str(inner_e)}"

            workbook.save(self.excel_file_path)
            driver.quit()
            self.status_label.setText(f"PRISM data written to: {self.excel_file_path}")

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)}")

    def write_to_excel(self, data_rows):
        if not self.excel_file_path:
            self.status_label.setText("No Excel file loaded.")
            return

        workbook = load_workbook(self.excel_file_path)
        sheet = workbook.active

        sheet["A1"] = "PRISM ID"
        sheet["B1"] = "Node"
        sheet["C1"] = "Design Status"

        for i, row in enumerate(data_rows, start=2):
            sheet[f"A{i}"] = row[0]
            sheet[f"B{i}"] = row[1]
            sheet[f"C{i}"] = row[2]

        workbook.save(self.excel_file_path)
        self.status_label.setText(f"Data written to: {self.excel_file_path}")

    def save_browser_setting(self):
        settings = {"browser": self.browser_dropdown.currentText()}
        with open(self.settings_file, "w") as f:
            json.dump(settings, f)

    def load_browser_setting(self):
        if os.path.exists(self.settings_file):
            with open(self.settings_file, "r") as f:
                settings = json.load(f)
                return settings.get("browser", "Edge")
        return "Edge"

# Run the app
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelAutomationApp()
    window.show()
    sys.exit(app.exec_())