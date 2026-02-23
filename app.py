import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, QTextEdit, QMessageBox,
                             QProgressBar, QGroupBox, QGridLayout, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject, QUrl
from PyQt5.QtWebEngineWidgets import QWebEngineView
import json

sys.path.insert(0, os.path.dirname(__file__))
from draw_suspension import (draw_full_suspension, draw_front_suspension, draw_rear_suspension,
                              count_hardpoints, count_wheels, load_json,
                              set_front_suspension_visibility, set_rear_suspension_visibility,
                              set_all_wheels_visibility, set_front_wheels_visibility,
                              set_rear_wheels_visibility, set_chassis_points_visibility,
                              set_non_chassis_visibility, set_all_suspension_visibility,
                              set_visibility_by_substring)
from optimumSheetParser import OptimumSheetParser
from test_solidworks_connection import get_active_document_name


class QtStream(QObject):
    """Redirect stdout to a PyQt signal."""
    text_written = pyqtSignal(str)

    def write(self, text):
        if text.strip():
            self.text_written.emit(text.strip())

    def flush(self):
        pass


class SolidWorksWorker(QThread):
    """Worker thread for SolidWorks operations."""
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(int)
    log = pyqtSignal(str)

    def __init__(self, operation, total_count, *args):
        super().__init__()
        self.operation = operation
        self.total_count = total_count
        self.args = args

    def progress_callback(self, count):
        self.progress.emit(count)

    def run(self):
        # Redirect stdout to log signal for this thread
        stream = QtStream()
        stream.text_written.connect(self.log.emit)
        old_stdout = sys.stdout
        sys.stdout = stream
        try:
            self.operation(*self.args, progress_callback=self.progress_callback)
            self.finished.emit(True, "Suspension imported successfully")
        except Exception as e:
            self.finished.emit(False, str(e))
        finally:
            sys.stdout = old_stdout


class ImportOptimumKTab(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title and description
        layout.addWidget(QLabel("Import OptimumK Files"))
        layout.addWidget(QLabel("Parse Excel files to generate JSON data"))
        
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


class WriteSolidworksTab(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()

        # Top bar with title on left, connection button on right
        h_top = QHBoxLayout()
        h_top.addWidget(QLabel("Write to SolidWorks"))
        h_top.addStretch()
        self.connection_label = QLabel("Not tested")
        h_top.addWidget(self.connection_label)
        btn_test_connection = QPushButton("Test Connection")
        btn_test_connection.clicked.connect(self.test_solidworks_connection)
        h_top.addWidget(btn_test_connection)
        layout.addLayout(h_top)

        layout.addWidget(QLabel("Import suspension geometry into SolidWorks"))

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Suspension import buttons
        layout.addWidget(QLabel("Import Suspension to SolidWorks:"))

        self.btn_import_full = QPushButton("Import Full Suspension")
        self.btn_import_full.clicked.connect(self.import_full_suspension)
        layout.addWidget(self.btn_import_full)

        self.btn_import_front = QPushButton("Import Front Suspension Only")
        self.btn_import_front.clicked.connect(self.import_front_suspension)
        layout.addWidget(self.btn_import_front)

        self.btn_import_rear = QPushButton("Import Rear Suspension Only")
        self.btn_import_rear.clicked.connect(self.import_rear_suspension)
        layout.addWidget(self.btn_import_rear)

        # Status / console output text box
        layout.addWidget(QLabel("Console Output:"))
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setFontFamily("Courier")
        self.status_text.setText("• Click 'Test Connection' to verify SolidWorks is running\n• Parse Excel file in Import tab first")
        layout.addWidget(self.status_text)

        btn_clear = QPushButton("Clear")
        btn_clear.clicked.connect(lambda: self.status_text.clear())
        layout.addWidget(btn_clear)

        layout.addStretch()
        self.setLayout(layout)

    def append_log(self, text):
        self.status_text.append(text)

    def test_solidworks_connection(self):
        try:
            doc_name = get_active_document_name()
            self.connection_label.setText(f"✓ {doc_name}")
            self.status_text.append(f"✓ Connected — Active document: {doc_name}")
        except Exception as e:
            self.connection_label.setText("✗ Not connected")
            self.status_text.append(f"✗ Connection failed: {str(e)}")
            QMessageBox.critical(self, "SolidWorks Connection Error",
                                 f"Could not connect to SolidWorks:\n\n{str(e)}\n\n"
                                 "Make sure:\n• SolidWorks is running\n• A document is open\n• pywin32 is installed")

    def set_buttons_enabled(self, enabled):
        self.btn_import_full.setEnabled(enabled)
        self.btn_import_front.setEnabled(enabled)
        self.btn_import_rear.setEnabled(enabled)

    def start_loading(self, message, total_count):
        self.progress_bar.setMaximum(total_count)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.set_buttons_enabled(False)
        self.status_text.append(f"\n{message}")

    def on_progress(self, count):
        self.progress_bar.setValue(self.progress_bar.value() + count)

    def stop_loading(self, success, message):
        self.progress_bar.setVisible(False)
        self.set_buttons_enabled(True)
        if success:
            self.status_text.append(f"✓ {message}")
            QMessageBox.information(self, "Success", message)
        else:
            self.status_text.append(f"✗ Error: {message}")
            QMessageBox.critical(self, "Error", f"Import failed: {message}")

    def _start_worker(self, operation, total, *args):
        self.start_loading(f"Running {operation.__name__}...", total)
        self.worker = SolidWorksWorker(operation, total, *args)
        self.worker.log.connect(self.append_log)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.stop_loading)
        self.worker.start()

    def import_full_suspension(self):
        front_file = os.path.join(os.path.dirname(__file__), "temp", "Front_Suspension.json")
        rear_file = os.path.join(os.path.dirname(__file__), "temp", "Rear_Suspension.json")
        vehicle_file = os.path.join(os.path.dirname(__file__), "temp", "Vehicle_Setup.json")

        if not all(os.path.exists(f) for f in [front_file, rear_file, vehicle_file]):
            QMessageBox.warning(self, "Missing Files", "Please parse an Excel file first")
            return
        try:
            front_data, rear_data = load_json(front_file), load_json(rear_file)
            total = (count_hardpoints(front_data) + count_wheels(front_data.get("Wheels", {})) +
                     count_hardpoints(rear_data) + count_wheels(rear_data.get("Wheels", {})))
            self._start_worker(draw_full_suspension, total, front_file, rear_file, vehicle_file)
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def import_front_suspension(self):
        front_file = os.path.join(os.path.dirname(__file__), "temp", "Front_Suspension.json")
        if not os.path.exists(front_file):
            QMessageBox.warning(self, "Missing File", "Please parse an Excel file first")
            return
        try:
            front_data = load_json(front_file)
            total = count_hardpoints(front_data) + count_wheels(front_data.get("Wheels", {}))
            self._start_worker(draw_front_suspension, total, front_file)
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def import_rear_suspension(self):
        rear_file = os.path.join(os.path.dirname(__file__), "temp", "Rear_Suspension.json")
        vehicle_file = os.path.join(os.path.dirname(__file__), "temp", "Vehicle_Setup.json")
        if not all(os.path.exists(f) for f in [rear_file, vehicle_file]):
            QMessageBox.warning(self, "Missing Files", "Please parse an Excel file first")
            return
        try:
            rear_data = load_json(rear_file)
            total = count_hardpoints(rear_data) + count_wheels(rear_data.get("Wheels", {}))
            self._start_worker(draw_rear_suspension, total, rear_file, vehicle_file)
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


class ViewTab(QWidget):
    """Tab for controlling visibility of suspension features in SolidWorks."""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        layout.addWidget(QLabel("Control Visibility of Suspension Features"))
        
        # Status/console output
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setFontFamily("Courier")
        self.status_text.setMaximumHeight(100)
        self.status_text.setText("Toggle visibility of suspension components in SolidWorks")
        
        # Main suspension groups
        group_main = QGroupBox("Suspension Groups")
        grid_main = QGridLayout()
        
        self.btn_show_all = QPushButton("Show All")
        self.btn_show_all.clicked.connect(lambda: self.set_visibility(set_all_suspension_visibility, True, "all"))
        grid_main.addWidget(self.btn_show_all, 0, 0)
        
        self.btn_hide_all = QPushButton("Hide All")
        self.btn_hide_all.clicked.connect(lambda: self.set_visibility(set_all_suspension_visibility, False, "all"))
        grid_main.addWidget(self.btn_hide_all, 0, 1)
        
        self.btn_show_front = QPushButton("Show Front")
        self.btn_show_front.clicked.connect(lambda: self.set_visibility(set_front_suspension_visibility, True, "front"))
        grid_main.addWidget(self.btn_show_front, 1, 0)
        
        self.btn_hide_front = QPushButton("Hide Front")
        self.btn_hide_front.clicked.connect(lambda: self.set_visibility(set_front_suspension_visibility, False, "front"))
        grid_main.addWidget(self.btn_hide_front, 1, 1)
        
        self.btn_show_rear = QPushButton("Show Rear")
        self.btn_show_rear.clicked.connect(lambda: self.set_visibility(set_rear_suspension_visibility, True, "rear"))
        grid_main.addWidget(self.btn_show_rear, 2, 0)
        
        self.btn_hide_rear = QPushButton("Hide Rear")
        self.btn_hide_rear.clicked.connect(lambda: self.set_visibility(set_rear_suspension_visibility, False, "rear"))
        grid_main.addWidget(self.btn_hide_rear, 2, 1)
        
        group_main.setLayout(grid_main)
        layout.addWidget(group_main)
        
        # Wheels group
        group_wheels = QGroupBox("Wheels")
        grid_wheels = QGridLayout()
        
        self.btn_show_wheels = QPushButton("Show All Wheels")
        self.btn_show_wheels.clicked.connect(lambda: self.set_visibility(set_all_wheels_visibility, True, "wheels"))
        grid_wheels.addWidget(self.btn_show_wheels, 0, 0)
        
        self.btn_hide_wheels = QPushButton("Hide All Wheels")
        self.btn_hide_wheels.clicked.connect(lambda: self.set_visibility(set_all_wheels_visibility, False, "wheels"))
        grid_wheels.addWidget(self.btn_hide_wheels, 0, 1)
        
        self.btn_show_front_wheels = QPushButton("Show Front Wheels")
        self.btn_show_front_wheels.clicked.connect(lambda: self.set_visibility(set_front_wheels_visibility, True, "front wheels"))
        grid_wheels.addWidget(self.btn_show_front_wheels, 1, 0)
        
        self.btn_hide_front_wheels = QPushButton("Hide Front Wheels")
        self.btn_hide_front_wheels.clicked.connect(lambda: self.set_visibility(set_front_wheels_visibility, False, "front wheels"))
        grid_wheels.addWidget(self.btn_hide_front_wheels, 1, 1)
        
        self.btn_show_rear_wheels = QPushButton("Show Rear Wheels")
        self.btn_show_rear_wheels.clicked.connect(lambda: self.set_visibility(set_rear_wheels_visibility, True, "rear wheels"))
        grid_wheels.addWidget(self.btn_show_rear_wheels, 2, 0)
        
        self.btn_hide_rear_wheels = QPushButton("Hide Rear Wheels")
        self.btn_hide_rear_wheels.clicked.connect(lambda: self.set_visibility(set_rear_wheels_visibility, False, "rear wheels"))
        grid_wheels.addWidget(self.btn_hide_rear_wheels, 2, 1)
        
        group_wheels.setLayout(grid_wheels)
        layout.addWidget(group_wheels)
        
        # Chassis/Non-Chassis group
        group_chassis = QGroupBox("Chassis vs Non-Chassis Points")
        grid_chassis = QGridLayout()
        
        self.btn_show_chassis = QPushButton("Show Chassis Points")
        self.btn_show_chassis.clicked.connect(lambda: self.set_visibility(set_chassis_points_visibility, True, "chassis points"))
        grid_chassis.addWidget(self.btn_show_chassis, 0, 0)
        
        self.btn_hide_chassis = QPushButton("Hide Chassis Points")
        self.btn_hide_chassis.clicked.connect(lambda: self.set_visibility(set_chassis_points_visibility, False, "chassis points"))
        grid_chassis.addWidget(self.btn_hide_chassis, 0, 1)
        
        self.btn_show_nonchassis = QPushButton("Show Non-Chassis")
        self.btn_show_nonchassis.clicked.connect(lambda: self.set_visibility(set_non_chassis_visibility, True, "non-chassis points"))
        grid_chassis.addWidget(self.btn_show_nonchassis, 1, 0)
        
        self.btn_hide_nonchassis = QPushButton("Hide Non-Chassis")
        self.btn_hide_nonchassis.clicked.connect(lambda: self.set_visibility(set_non_chassis_visibility, False, "non-chassis points"))
        grid_chassis.addWidget(self.btn_hide_nonchassis, 1, 1)
        
        group_chassis.setLayout(grid_chassis)
        layout.addWidget(group_chassis)
        
        # Custom substring search
        group_custom = QGroupBox("Custom Filter (by name substring)")
        h_custom = QHBoxLayout()
        
        from PyQt5.QtWidgets import QLineEdit
        self.custom_filter = QLineEdit()
        self.custom_filter.setPlaceholderText("Enter text to match (e.g., 'Upright', 'Pushrod')")
        h_custom.addWidget(self.custom_filter)
        
        btn_show_custom = QPushButton("Show")
        btn_show_custom.clicked.connect(self.show_custom)
        h_custom.addWidget(btn_show_custom)
        
        btn_hide_custom = QPushButton("Hide")
        btn_hide_custom.clicked.connect(self.hide_custom)
        h_custom.addWidget(btn_hide_custom)
        
        group_custom.setLayout(h_custom)
        layout.addWidget(group_custom)
        
        # Console output
        layout.addWidget(QLabel("Output:"))
        layout.addWidget(self.status_text)
        
        btn_clear = QPushButton("Clear Log")
        btn_clear.clicked.connect(lambda: self.status_text.clear())
        layout.addWidget(btn_clear)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def set_visibility(self, func, visible, description):
        """Execute a visibility function and update status."""
        action = "Showing" if visible else "Hiding"
        self.status_text.append(f"{action} {description}...")
        try:
            success = func(visible)
            if success:
                self.status_text.append(f"✓ {description} {'shown' if visible else 'hidden'}")
            else:
                self.status_text.append(f"✗ Failed to modify {description}")
        except FileNotFoundError as e:
            self.status_text.append(f"✗ Error: {str(e)}")
            QMessageBox.critical(self, "Error", 
                                 f"SuspensionTools.exe not found.\n\n"
                                 "Run 'dotnet build -c Release' in the sw_drawer folder first.")
        except Exception as e:
            self.status_text.append(f"✗ Error: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed: {str(e)}")
    
    def show_custom(self):
        """Show features matching custom filter."""
        text = self.custom_filter.text().strip()
        if not text:
            QMessageBox.warning(self, "Warning", "Please enter a filter text")
            return
        self.status_text.append(f"Showing features containing '{text}'...")
        try:
            set_visibility_by_substring(text, True)
            self.status_text.append(f"✓ Done")
        except Exception as e:
            self.status_text.append(f"✗ Error: {str(e)}")
    
    def hide_custom(self):
        """Hide features matching custom filter."""
        text = self.custom_filter.text().strip()
        if not text:
            QMessageBox.warning(self, "Warning", "Please enter a filter text")
            return
        self.status_text.append(f"Hiding features containing '{text}'...")
        try:
            set_visibility_by_substring(text, False)
            self.status_text.append(f"✓ Done")
        except Exception as e:
            self.status_text.append(f"✗ Error: {str(e)}")


class HelpTab(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        
        web = QWebEngineView()
        help_path = os.path.join(os.path.dirname(__file__), "help.htm")
        web.load(QUrl.fromLocalFile(help_path))
        layout.addWidget(web)
        
        self.setLayout(layout)


class OptimumKApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("OptimumK SolidWorks Plugin")
        self.setGeometry(100, 100, 600, 600)
        
        tabs = QTabWidget()
        
        tabs.addTab(ImportOptimumKTab(), "Import OptimumK")
        tabs.addTab(WriteSolidworksTab(), "Write to SolidWorks")
        tabs.addTab(ViewTab(), "View")
        tabs.addTab(HelpTab(), "Help")
        
        self.setCentralWidget(tabs)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OptimumKApp()
    window.show()
    sys.exit(app.exec_())