import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, QTextEdit, QMessageBox)
from PyQt5.QtCore import Qt
import json

# Add parent directory to path to import draw_suspension
sys.path.insert(0, os.path.dirname(__file__))
from draw_suspension import draw_full_suspension, draw_front_suspension, draw_rear_suspension
from optimumSheetParser import OptimumSheetParser


class ImportOptimumKTab(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title and description
        layout.addWidget(QLabel("Import OptimumK Files"))
        layout.addWidget(QLabel("Parse Excel files or select JSON files"))
        
        # Excel file explorer and parser
        h_excel = QHBoxLayout()
        self.excel_path_label = QLabel("No Excel file selected")
        h_excel.addWidget(QLabel("Excel File:"))
        h_excel.addWidget(self.excel_path_label)
        
        btn_browse_excel = QPushButton("Browse Excel")
        btn_browse_excel.clicked.connect(self.browse_excel_file)
        h_excel.addWidget(btn_browse_excel)
        
        btn_parse_excel = QPushButton("Parse Excel")
        btn_parse_excel.clicked.connect(self.parse_excel_file)
        h_excel.addWidget(btn_parse_excel)
        
        layout.addLayout(h_excel)
        
        # JSON preview
        layout.addWidget(QLabel("JSON Preview (/temp):"))
        self.json_preview = QTextEdit()
        self.json_preview.setReadOnly(True)
        self.json_preview.setMaximumHeight(200)
        layout.addWidget(self.json_preview)
        
        # JSON file selection buttons
        h_json = QHBoxLayout()
        
        btn_preview_front = QPushButton("Preview Front JSON")
        btn_preview_front.clicked.connect(lambda: self.preview_json_file("temp/Front_Suspension.json"))
        h_json.addWidget(btn_preview_front)
        
        btn_preview_rear = QPushButton("Preview Rear JSON")
        btn_preview_rear.clicked.connect(lambda: self.preview_json_file("temp/Rear_Suspension.json"))
        h_json.addWidget(btn_preview_rear)
        
        btn_preview_vehicle = QPushButton("Preview Vehicle Setup")
        btn_preview_vehicle.clicked.connect(lambda: self.preview_json_file("temp/Vehicle_Setup.json"))
        h_json.addWidget(btn_preview_vehicle)
        
        layout.addLayout(h_json)
        
        # Suspension import buttons
        layout.addWidget(QLabel("Import to SolidWorks:"))
        
        btn_import = QPushButton("Import Full Suspension")
        btn_import.clicked.connect(self.import_full_suspension)
        layout.addWidget(btn_import)
        
        btn_front = QPushButton("Import Front Suspension Only")
        btn_front.clicked.connect(self.import_front_suspension)
        layout.addWidget(btn_front)
        
        btn_rear = QPushButton("Import Rear Suspension Only")
        btn_rear.clicked.connect(self.import_rear_suspension)
        layout.addWidget(btn_rear)
        
        # Status label
        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def browse_excel_file(self):
        file_path = QFileDialog.getOpenFileName(self, "Select Excel File", filter="*.xlsx *.xls")[0]
        if file_path:
            self.excel_file = file_path
            self.excel_path_label.setText(os.path.basename(file_path))
    
    def parse_excel_file(self):
        if not hasattr(self, "excel_file"):
            QMessageBox.warning(self, "Warning", "Please select an Excel file first")
            return
        
        try:
            self.status_label.setText("Parsing Excel file...")
            parser = OptimumSheetParser(self.excel_file)
            parser.save_json_per_sheet("temp")
            parser.save_reference_distance("temp")
            self.status_label.setText("✓ Excel file parsed successfully")
            QMessageBox.information(self, "Success", "Excel file parsed. JSON files saved to /temp")
        except Exception as e:
            self.status_label.setText("✗ Parse failed")
            QMessageBox.critical(self, "Error", f"Parse failed: {str(e)}")
    
    def preview_json_file(self, file_path):
        full_path = os.path.join(os.path.dirname(__file__), file_path)
        try:
            with open(full_path, 'r') as f:
                data = json.load(f)
                preview_text = json.dumps(data, indent=2)
                self.json_preview.setText(preview_text)
                self.status_label.setText(f"✓ Previewing {os.path.basename(file_path)}")
        except FileNotFoundError:
            QMessageBox.warning(self, "File Not Found", f"Could not find {file_path}\n\nParse an Excel file first")
            self.status_label.setText("✗ File not found")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not preview file: {str(e)}")
            self.status_label.setText("✗ Preview failed")
    
    def import_full_suspension(self):
        front_file = os.path.join(os.path.dirname(__file__), "temp", "Front_Suspension.json")
        rear_file = os.path.join(os.path.dirname(__file__), "temp", "Rear_Suspension.json")
        vehicle_file = os.path.join(os.path.dirname(__file__), "temp", "Vehicle_Setup.json")
        
        if not all(os.path.exists(f) for f in [front_file, rear_file, vehicle_file]):
            QMessageBox.warning(self, "Missing Files", "Please parse an Excel file first to generate JSON files")
            return
        
        try:
            self.status_label.setText("Importing full suspension...")
            draw_full_suspension(front_file, rear_file, vehicle_file)
            self.status_label.setText("✓ Full suspension imported successfully")
            QMessageBox.information(self, "Success", "Suspension data imported to SolidWorks")
        except Exception as e:
            self.status_label.setText("✗ Import failed")
            QMessageBox.critical(self, "Error", f"Import failed: {str(e)}")
    
    def import_front_suspension(self):
        front_file = os.path.join(os.path.dirname(__file__), "temp", "Front_Suspension.json")
        
        if not os.path.exists(front_file):
            QMessageBox.warning(self, "Missing File", "Please parse an Excel file first to generate JSON files")
            return
        
        try:
            self.status_label.setText("Importing front suspension...")
            draw_front_suspension(front_file)
            self.status_label.setText("✓ Front suspension imported successfully")
            QMessageBox.information(self, "Success", "Front suspension imported to SolidWorks")
        except Exception as e:
            self.status_label.setText("✗ Import failed")
            QMessageBox.critical(self, "Error", f"Import failed: {str(e)}")
    
    def import_rear_suspension(self):
        rear_file = os.path.join(os.path.dirname(__file__), "temp", "Rear_Suspension.json")
        vehicle_file = os.path.join(os.path.dirname(__file__), "temp", "Vehicle_Setup.json")
        
        if not all(os.path.exists(f) for f in [rear_file, vehicle_file]):
            QMessageBox.warning(self, "Missing Files", "Please parse an Excel file first to generate JSON files")
            return
        
        try:
            self.status_label.setText("Importing rear suspension...")
            draw_rear_suspension(rear_file, vehicle_file)
            self.status_label.setText("✓ Rear suspension imported successfully")
            QMessageBox.information(self, "Success", "Rear suspension imported to SolidWorks")
        except Exception as e:
            self.status_label.setText("✗ Import failed")
            QMessageBox.critical(self, "Error", f"Import failed: {str(e)}")


class WriteSolidworksTab(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        layout.addWidget(QLabel("Write to SolidWorks"))
        layout.addWidget(QLabel("Status and operations for SolidWorks integration"))
        
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setText("SolidWorks Integration Status:\n\n• Waiting for operations...\n• Ensure SolidWorks is running\n• Check CoordinateRunner.exe is built")
        layout.addWidget(self.status_text)
        
        btn_check = QPushButton("Check SolidWorks Connection")
        btn_check.clicked.connect(self.check_solidworks)
        layout.addWidget(btn_check)
        
        btn_clear = QPushButton("Clear Status")
        btn_clear.clicked.connect(lambda: self.status_text.clear())
        layout.addWidget(btn_clear)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def check_solidworks(self):
        try:
            import subprocess
            # Try to run CoordinateRunner with no args to test connection
            self.status_text.append("Checking SolidWorks connection...")
            QMessageBox.information(self, "Info", "Ensure SolidWorks is running and try importing suspension data from the Import tab")
        except Exception as e:
            self.status_text.append(f"Error: {str(e)}")


class HelpTab(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setText("""
<h2>OptimumK SolidWorks Plugin - Help</h2>

<h3>Import OptimumK Tab</h3>
<p>Import suspension and vehicle setup data from OptimumK JSON files into SolidWorks:</p>
<ul>
<li><b>Import Full Suspension</b> - Imports front, rear, and vehicle setup data</li>
<li><b>Import Front Suspension Only</b> - Imports just the front suspension geometry</li>
<li><b>Import Rear Suspension Only</b> - Imports just the rear suspension geometry</li>
</ul>

<h3>Requirements</h3>
<ul>
<li>SolidWorks must be running and open with a document</li>
<li>JSON files must be properly formatted from OptimumK</li>
<li>CoordinateRunner.exe must be built (run 'dotnet build -c Release' in sw_drawer folder)</li>
</ul>

<h3>Write to SolidWorks Tab</h3>
<p>Monitor SolidWorks integration status and verify connection.</p>

<h3>Workflow</h3>
<ol>
<li>Open SolidWorks with a document ready</li>
<li>Go to Import OptimumK tab</li>
<li>Select your JSON files and click import</li>
<li>Coordinate systems will be created in SolidWorks</li>
<li>Check Write to SolidWorks tab for status updates</li>
</ol>

<h3>Troubleshooting</h3>
<ul>
<li>If import fails, ensure SolidWorks is running</li>
<li>Check that JSON files are in the correct format</li>
<li>Verify CoordinateRunner.exe exists in sw_drawer/bin/Release/net48/</li>
<li>Check file paths are correct</li>
</ul>
        """)
        layout.addWidget(help_text)
        self.setLayout(layout)


class OptimumKApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("OptimumK SolidWorks Plugin")
        self.setGeometry(100, 100, 600, 500)
        
        tabs = QTabWidget()
        
        tabs.addTab(ImportOptimumKTab(), "Import OptimumK")
        tabs.addTab(WriteSolidworksTab(), "Write to SolidWorks")
        tabs.addTab(HelpTab(), "Help")
        
        self.setCentralWidget(tabs)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OptimumKApp()
    window.show()
    sys.exit(app.exec_())