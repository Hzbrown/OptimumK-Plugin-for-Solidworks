import sys
import os
import re
import json
import subprocess
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, QTextEdit, QMessageBox,
                             QProgressBar, QGroupBox, QGridLayout, QCheckBox, QLineEdit, QSpinBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject, QUrl
from PyQt5.QtWebEngineWidgets import QWebEngineView

sys.path.insert(0, os.path.dirname(__file__))
from coordinate_insertion import CoordinateInsertionWorker, insert_coordinates, validate_files, get_marker_path, create_coordinates_folder
from pose_creation import PoseCreationWorker, create_pose, validate_pose_name, get_existing_poses
from visualization_control import VisualizationWorker, set_suspension_visibility, set_marker_visibility, get_visualization_controls, get_color_coding_info
from optimumSheetParser import OptimumSheetParser
from test_solidworks_connection import get_active_document_name
from draw_suspension import (
    load_json, count_hardpoints, count_wheels,
    draw_full_suspension, draw_front_suspension, draw_rear_suspension,
    set_all_suspension_visibility, set_front_suspension_visibility, set_rear_suspension_visibility,
    set_all_wheels_visibility, set_front_wheels_visibility, set_rear_wheels_visibility,
    set_chassis_points_visibility, set_non_chassis_visibility,
    set_visibility_by_substring,
    set_all_markers_visibility, set_front_markers_visibility, set_rear_markers_visibility,
    set_marker_visibility_by_name,
    create_all_markers_with_worker, delete_all_markers_with_worker
)


class QtStream(QObject):
    """Redirect stdout to a PyQt signal."""
    text_written = pyqtSignal(str)

    def write(self, text):
        if text.strip():
            self.text_written.emit(text.strip())

    def flush(self):
        pass


class SolidWorksWorker(QThread):
    """Worke.r thread for SolidWorks operations."""
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


class HardpointWorker(QThread):
    """Worker thread for hardpoint operations."""
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(int, int)  # current, total
    state_changed = pyqtSignal(str)  # state description
    log = pyqtSignal(str)

    def __init__(self, operation, *args):
        super().__init__()
        self.operation = operation
        self.args = args
        self._abort = False
        self._process = None
        self._total_tasks = 0
        self._current_progress = 0

    def abort(self):
        """Request abort and terminate any running subprocess."""
        self._abort = True
        if self._process and self._process.poll() is None:
            try:
                self._process.terminate()
                self._process.wait(timeout=2)
            except:
                try:
                    self._process.kill()
                except:
                    pass

    def parse_output_line(self, line):
        """Parse special output lines for progress and state."""
        if line.startswith("TOTAL:"):
            try:
                new_total = int(line.split(":")[1])
                # Always update to the latest total (may be refined during execution)
                self._total_tasks = new_total
                # Emit progress with updated total, keeping current progress
                self.progress.emit(self._current_progress, self._total_tasks)
            except:
                pass
            return None  # Don't show in log
        elif line.startswith("PROGRESS:"):
            try:
                self._current_progress = int(line.split(":")[1])
                self.progress.emit(self._current_progress, self._total_tasks)
            except:
                pass
            return None  # Don't show in log
        elif line.startswith("STATE:"):
            state = line.split(":")[1]
            state_descriptions = {
                "Initializing": "Starting...",
                "LoadingJson": "Loading JSON data...",
                "LoadingMarkerPart": "Loading marker part...",
                "InsertingBodies": "Inserting marker bodies...",
                "RenamingBodies": "Renaming bodies...",
                "ApplyingColors": "Applying colors...",
                "CreatingCoordinateSystems": "Creating coordinate systems...",
                "CreatingHardpointsFolder": "Creating Hardpoints folder...",
                "CreatingTransformsFolder": "Creating Transforms folder...",
                "CreatingTransforms": "Creating transform features...",
                "UpdatingSuppression": "Updating suppression...",
                "Rebuilding": "Rebuilding model...",
                "Complete": "Done"
            }
            description = state_descriptions.get(state, state)
            self.state_changed.emit(description)
            return None  # Don't show in log
        return line  # Normal log line

    def run(self):
        stream = QtStream()
        stream.text_written.connect(self._handle_log)
        old_stdout = sys.stdout
        sys.stdout = stream
        try:
            result = self.operation(*self.args, worker=self)
            if self._abort:
                self.finished.emit(False, "Operation aborted")
            else:
                self.finished.emit(True, "Operation completed successfully")
        except Exception as e:
            if self._abort:
                self.finished.emit(False, "Operation aborted")
            else:
                self.finished.emit(False, str(e))
        finally:
            sys.stdout = old_stdout

    def _handle_log(self, text):
        """Handle log output, parsing special lines."""
        result = self.parse_output_line(text)
        if result is not None:
            self.log.emit(result)


class MarkerWorker(QThread):
    """Worker thread for marker operations."""
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(int, int)  # current, total
    state_changed = pyqtSignal(str)  # state description
    log = pyqtSignal(str)

    def __init__(self, operation, *args):
        super().__init__()
        self.operation = operation
        self.args = args
        self._abort = False
        self._process = None
        self._total_tasks = 0
        self._current_progress = 0

    def abort(self):
        """Request abort and terminate any running subprocess."""
        self._abort = True
        if self._process and self._process.poll() is None:
            try:
                self._process.terminate()
                self._process.wait(timeout=2)
            except:
                try:
                    self._process.kill()
                except:
                    pass

    def parse_output_line(self, line):
        """Parse special output lines for progress and state."""
        if line.startswith("TOTAL:"):
            try:
                new_total = int(line.split(":")[1])
                # Always update to the latest total (may be refined during execution)
                self._total_tasks = new_total
                # Emit progress with updated total, keeping current progress
                self.progress.emit(self._current_progress, self._total_tasks)
            except:
                pass
            return None  # Don't show in log
        elif line.startswith("PROGRESS:"):
            try:
                self._current_progress = int(line.split(":")[1])
                self.progress.emit(self._current_progress, self._total_tasks)
            except:
                pass
            return None  # Don't show in log
        elif line.startswith("STATE:"):
            state = line.split(":")[1]
            state_descriptions = {
                "Initializing": "Starting...",
                "LoadingMarker": "Reading marker file...",
                "ScanningCoordSystems": "Scanning coordinate systems...",
                "InsertingComponents": "Inserting markers...",
                "MatingMarkers": "Mating markers...",
                "PostProcessing": "Post-processing markers...",
                "Cleanup": "Cleaning up...",
                "Complete": "Done"
            }
            description = state_descriptions.get(state, state)
            self.state_changed.emit(description)
            return None  # Don't show in log
        return line  # Normal log line

    def run(self):
        stream = QtStream()
        stream.text_written.connect(self._handle_log)
        old_stdout = sys.stdout
        sys.stdout = stream
        try:
            result = self.operation(*self.args, worker=self)
            if self._abort:
                self.finished.emit(False, "Operation aborted")
            else:
                self.finished.emit(True, "Operation completed successfully")
        except Exception as e:
            if self._abort:
                self.finished.emit(False, "Operation aborted")
            else:
                self.finished.emit(False, str(e))
        finally:
            sys.stdout = old_stdout

    def _handle_log(self, text):
        """Handle log output, parsing special lines."""
        result = self.parse_output_line(text)
        if result is not None:
            self.log.emit(result)


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


class MarkersTab(QWidget):
    """Tab for creating and controlling visibility of marker spheres."""
    
    def __init__(self):
        super().__init__()
        self.worker = None
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        layout.addWidget(QLabel("Marker Spheres - Visual indicators at hardpoints"))
        
        # State label
        self.state_label = QLabel("")
        self.state_label.setStyleSheet("font-weight: bold; color: #0066cc;")
        self.state_label.setVisible(False)
        layout.addWidget(self.state_label)
        
        # Progress bar and abort button (hidden by default)
        h_progress = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFormat("%v / %m (%p%)")
        h_progress.addWidget(self.progress_bar)
        
        self.btn_abort = QPushButton("Abort")
        self.btn_abort.setVisible(False)
        self.btn_abort.setStyleSheet("background-color: #ff6666;")
        self.btn_abort.clicked.connect(self.abort_operation)
        h_progress.addWidget(self.btn_abort)
        
        layout.addLayout(h_progress)
        
        # Create/Delete section
        group_create = QGroupBox("Create / Delete Markers")
        h_create = QHBoxLayout()
        
        h_create.addWidget(QLabel("Size (mm):"))
        
        self.radius_spin = QSpinBox()
        self.radius_spin.setRange(1, 50)
        self.radius_spin.setValue(5)
        h_create.addWidget(self.radius_spin)
        
        self.btn_create_all = QPushButton("Create All Markers")
        self.btn_create_all.clicked.connect(self.create_all_markers)
        h_create.addWidget(self.btn_create_all)
        
        self.btn_delete_all = QPushButton("Delete All Markers")
        self.btn_delete_all.clicked.connect(self.delete_all_markers)
        self.btn_delete_all.setStyleSheet("background-color: #ffcccc;")
        h_create.addWidget(self.btn_delete_all)
        
        group_create.setLayout(h_create)
        layout.addWidget(group_create)
        
        # Show/Hide All and Front/Rear
        group_main = QGroupBox("All / Front / Rear")
        grid_main = QGridLayout()
        
        btn_show_all = QPushButton("Show All")
        btn_show_all.clicked.connect(lambda: self.set_visibility(set_all_markers_visibility, True, "all markers"))
        grid_main.addWidget(btn_show_all, 0, 0)
        
        btn_hide_all = QPushButton("Hide All")
        btn_hide_all.clicked.connect(lambda: self.set_visibility(set_all_markers_visibility, False, "all markers"))
        grid_main.addWidget(btn_hide_all, 0, 1)
        
        btn_show_front = QPushButton("Show FRONT")
        btn_show_front.clicked.connect(lambda: self.set_visibility(set_front_markers_visibility, True, "front markers"))
        grid_main.addWidget(btn_show_front, 1, 0)
        
        btn_hide_front = QPushButton("Hide FRONT")
        btn_hide_front.clicked.connect(lambda: self.set_visibility(set_front_markers_visibility, False, "front markers"))
        grid_main.addWidget(btn_hide_front, 1, 1)
        
        btn_show_rear = QPushButton("Show REAR")
        btn_show_rear.clicked.connect(lambda: self.set_visibility(set_rear_markers_visibility, True, "rear markers"))
        grid_main.addWidget(btn_show_rear, 2, 0)
        
        btn_hide_rear = QPushButton("Hide REAR")
        btn_hide_rear.clicked.connect(lambda: self.set_visibility(set_rear_markers_visibility, False, "rear markers"))
        grid_main.addWidget(btn_hide_rear, 2, 1)
        
        group_main.setLayout(grid_main)
        layout.addWidget(group_main)
        
        # Component Types - matching actual JSON naming patterns
        group_types = QGroupBox("By Component Type (JSON prefixes)")
        grid_types = QGridLayout()
        
        # Row 0: CHAS_ (Chassis) and UPRI_ (Upright)
        btn_show_chas = QPushButton("Show CHAS_")
        btn_show_chas.setStyleSheet("background-color: #FF0000; color: white;")
        btn_show_chas.clicked.connect(lambda: self.set_name_visibility("CHAS_", True))
        grid_types.addWidget(btn_show_chas, 0, 0)
        
        btn_hide_chas = QPushButton("Hide")
        btn_hide_chas.clicked.connect(lambda: self.set_name_visibility("CHAS_", False))
        grid_types.addWidget(btn_hide_chas, 0, 1)
        
        btn_show_upri = QPushButton("Show UPRI_")
        btn_show_upri.setStyleSheet("background-color: #0000FF; color: white;")
        btn_show_upri.clicked.connect(lambda: self.set_name_visibility("UPRI_", True))
        grid_types.addWidget(btn_show_upri, 0, 2)
        
        btn_hide_upri = QPushButton("Hide")
        btn_hide_upri.clicked.connect(lambda: self.set_name_visibility("UPRI_", False))
        grid_types.addWidget(btn_hide_upri, 0, 3)
        
        # Row 1: ROCK_ (Rocker) and NSMA_ (Non-Sprung Mass)
        btn_show_rock = QPushButton("Show ROCK_")
        btn_show_rock.setStyleSheet("background-color: #0080FF; color: white;")
        btn_show_rock.clicked.connect(lambda: self.set_name_visibility("ROCK_", True))
        grid_types.addWidget(btn_show_rock, 1, 0)
        
        btn_hide_rock = QPushButton("Hide")
        btn_hide_rock.clicked.connect(lambda: self.set_name_visibility("ROCK_", False))
        grid_types.addWidget(btn_hide_rock, 1, 1)
        
        btn_show_nsma = QPushButton("Show NSMA_")
        btn_show_nsma.setStyleSheet("background-color: #FFC0CB; color: black;")
        btn_show_nsma.clicked.connect(lambda: self.set_name_visibility("NSMA_", True))
        grid_types.addWidget(btn_show_nsma, 1, 2)
        
        btn_hide_nsma = QPushButton("Hide")
        btn_hide_nsma.clicked.connect(lambda: self.set_name_visibility("NSMA_", False))
        grid_types.addWidget(btn_hide_nsma, 1, 3)
        
        group_types.setLayout(grid_types)
        layout.addWidget(group_types)
        
        # Custom filter
        group_custom = QGroupBox("Custom Filter")
        h_custom = QHBoxLayout()
        
        self.custom_filter = QLineEdit()
        self.custom_filter.setPlaceholderText("Filter by name (e.g., 'Low', 'Upp', 'Piv')")
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
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setFontFamily("Courier")
        self.status_text.setMaximumHeight(120)
        self.status_text.setText("Ready")
        layout.addWidget(self.status_text)
        
        btn_clear = QPushButton("Clear Log")
        btn_clear.clicked.connect(lambda: self.status_text.clear())
        layout.addWidget(btn_clear)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def set_buttons_enabled(self, enabled):
        """Enable or disable all action buttons."""
        self.btn_create_all.setEnabled(enabled)
        self.btn_delete_all.setEnabled(enabled)
    
    def on_progress(self, current, total):
        """Update progress bar with current/total values like progressBar.setValue(i)."""
        if total > 0:
            # Set up determinate mode if not already set
            if self.progress_bar.maximum() != total:
                self.progress_bar.setRange(0, total)  # Like progressBar.setRange(0, steps)
                self.progress_bar.setFormat("%v / %m (%p%)")  # Show "current / max (percentage%)"
            
            # Set current progress directly like progressBar.setValue(i)
            self.progress_bar.setValue(current)
        else:
            # Keep indeterminate mode if total is 0 or unknown
            if self.progress_bar.maximum() != 0:
                self.progress_bar.setRange(0, 0)  # Indeterminate mode
                self.progress_bar.setFormat("Working...")
    
    def on_state_changed(self, state_description):
        """Update state label with current operation state."""
        self.state_label.setText(state_description)
        self.state_label.setVisible(True)
    
    def start_loading(self, message):
        """Show progress bar and abort button in indeterminate mode initially."""
        # Start in indeterminate mode (animated) like progressBar.setRange(0, 0)
        self.progress_bar.setRange(0, 0)  # Indeterminate mode until we get TOTAL
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Starting...")  # Show status text initially
        self.progress_bar.setVisible(True)
        self.btn_abort.setVisible(True)
        self.state_label.setText("Starting...")
        self.state_label.setVisible(True)
        self.set_buttons_enabled(False)
        self.status_text.append(f"\n{message}")
    
    def stop_loading(self, success, message):
        """Hide progress bar and abort button."""
        self.progress_bar.setVisible(False)
        self.btn_abort.setVisible(False)
        self.state_label.setVisible(False)
        self.set_buttons_enabled(True)
        self.worker = None
        if success:
            self.status_text.append(f"✓ {message}")
        else:
            self.status_text.append(f"✗ {message}")
    
    def append_log(self, text):
        """Append text to log."""
        self.status_text.append(text)
    
    def abort_operation(self):
        """Abort the current operation."""
        if self.worker:
            self.status_text.append("Aborting operation...")
            self.worker.abort()
    
    def create_all_markers(self):
        """Create markers at all coordinate systems using worker thread."""
        radius = self.radius_spin.value()
        self.start_loading(f"Creating markers (size: {radius}mm)...")
        
        self.worker = MarkerWorker(create_all_markers_with_worker, radius)
        self.worker.log.connect(self.append_log)
        self.worker.progress.connect(self.on_progress)
        self.worker.state_changed.connect(self.on_state_changed)
        self.worker.finished.connect(self.stop_loading)
        self.worker.start()
    
    def delete_all_markers(self):
        """Delete all markers after confirmation."""
        reply = QMessageBox.question(self, "Confirm", "Delete all markers?", 
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.start_loading("Deleting markers...")
            
            self.worker = MarkerWorker(delete_all_markers_with_worker)
            self.worker.log.connect(self.append_log)
            self.worker.progress.connect(self.on_progress)
            self.worker.state_changed.connect(self.on_state_changed)
            self.worker.finished.connect(self.stop_loading)
            self.worker.start()

    def set_visibility(self, func, visible, description):
        """Execute a visibility function and update status."""
        action = "Showing" if visible else "Hiding"
        self.status_text.append(f"{action} {description}...")
        try:
            success = func(visible)
            self.status_text.append("✓ Done" if success else "✗ Failed")
        except FileNotFoundError:
            self.status_text.append("✗ InsertMarker.exe not found. Build the project first.")
            QMessageBox.critical(self, "Error", 
                                 "InsertMarker.exe not found.\n\n"
                                 "Run 'dotnet build -c Release' in the InsertMarker folder first.")
        except Exception as e:
            self.status_text.append(f"✗ {str(e)}")
    
    def set_name_visibility(self, substring, visible):
        """Show/hide markers by name substring."""
        action = "Showing" if visible else "Hiding"
        self.status_text.append(f"{action} {substring} markers...")
        try:
            success = set_marker_visibility_by_name(substring, visible)
            self.status_text.append("✓ Done" if success else "✗ Failed")
        except FileNotFoundError:
            self.status_text.append("✗ InsertMarker.exe not found. Build the project first.")
        except Exception as e:
            self.status_text.append(f"✗ {str(e)}")
    
    def show_custom(self):
        """Show markers matching custom filter."""
        text = self.custom_filter.text().strip()
        if not text:
            QMessageBox.warning(self, "Warning", "Enter filter text")
            return
        self.set_name_visibility(text, True)
    
    def hide_custom(self):
        """Hide markers matching custom filter."""
        text = self.custom_filter.text().strip()
        if not text:
            QMessageBox.warning(self, "Warning", "Enter filter text")
            return
        self.set_name_visibility(text, False)


class CoordinateInsertionTab(QWidget):
    """Tab for inserting coordinate systems with color coding and naming."""
    
    def __init__(self):
        super().__init__()
        self.worker = None
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        layout.addWidget(QLabel("Insert Coordinates"))
        layout.addWidget(QLabel("Creates virtual hardpoint markers from JSON with automatic folder organization"))
        
        # State label
        self.state_label = QLabel("")
        self.state_label.setStyleSheet("font-weight: bold; color: #0066cc;")
        self.state_label.setVisible(False)
        layout.addWidget(self.state_label)
        
        # Progress bar and abort button (hidden by default)
        h_progress = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFormat("%v / %m (%p%)")
        h_progress.addWidget(self.progress_bar)
        
        self.btn_abort = QPushButton("Abort")
        self.btn_abort.setVisible(False)
        self.btn_abort.setStyleSheet("background-color: #ff6666;")
        self.btn_abort.clicked.connect(self.abort_operation)
        h_progress.addWidget(self.btn_abort)
        
        layout.addLayout(h_progress)
        
        # Automatic paths
        group_files = QGroupBox("Automatic Path Configuration")
        v_files = QVBoxLayout()
        v_files.addWidget(QLabel("JSON File: Always uses latest from /temp"))
        v_files.addWidget(QLabel("Marker Part: Always uses Marker.SLDPRT in plugin folder"))
        group_files.setLayout(v_files)
        layout.addWidget(group_files)
        
        # What this does
        group_workflow = QGroupBox("What This Does")
        v_workflow = QVBoxLayout()
        v_workflow.addWidget(QLabel("✓ Reads all hardpoints from Front_Suspension.json (including wheels)"))
        v_workflow.addWidget(QLabel("✓ Inserts virtual Marker.SLDPRT at each location"))
        v_workflow.addWidget(QLabel("✓ Colors components by type (CHAS=Red, UPRI=Blue, wheels=Green, etc.)"))
        v_workflow.addWidget(QLabel("✓ Creates 'Hardpoints' folder in feature tree"))
        v_workflow.addWidget(QLabel("✓ Adds coordinate systems for mating"))
        v_workflow.addWidget(QLabel("✓ Organizes all hardpoints into the Hardpoints folder"))
        group_workflow.setLayout(v_workflow)
        layout.addWidget(group_workflow)
        
        # Operations
        self.btn_insert = QPushButton("Insert Hardpoints")
        self.btn_insert.setMinimumHeight(40)
        self.btn_insert.setStyleSheet("font-size: 12pt; font-weight: bold;")
        self.btn_insert.clicked.connect(self.insert_hardpoints)
        layout.addWidget(self.btn_insert)
        
        # Add a separator and note about folder organization
        layout.addWidget(QLabel(""))
        note_label = QLabel("Note: All hardpoints are automatically organized into folders in the SolidWorks feature tree")
        note_label.setStyleSheet("color: #666666; font-style: italic;")
        layout.addWidget(note_label)
        
        # Color coding information
        group_colors = QGroupBox("Color Coding")
        layout_colors = QVBoxLayout()
        
        color_info = get_color_coding_info()
        for prefix, info in color_info.items():
            color_label = QLabel(f"{info['name']}: RGB({info['rgb'][0]}, {info['rgb'][1]}, {info['rgb'][2]})")
            color_label.setStyleSheet(f"background-color: rgb({info['rgb'][0]}, {info['rgb'][1]}, {info['rgb'][2]}); color: white; padding: 2px;")
            layout_colors.addWidget(color_label)
        
        group_colors.setLayout(layout_colors)
        layout.addWidget(group_colors)
        
        # Console output
        layout.addWidget(QLabel("Output:"))
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setFontFamily("Courier")
        self.status_text.setMaximumHeight(150)
        self.status_text.setText("Ready")
        layout.addWidget(self.status_text)
        
        btn_clear = QPushButton("Clear Log")
        btn_clear.clicked.connect(lambda: self.status_text.clear())
        layout.addWidget(btn_clear)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def set_buttons_enabled(self, enabled):
        """Enable or disable all action buttons."""
        self.btn_insert.setEnabled(enabled)
    
    def on_progress(self, current, total):
        """Update progress bar with current/total values."""
        if total > 0:
            if self.progress_bar.maximum() != total:
                self.progress_bar.setRange(0, total)
                self.progress_bar.setFormat("%v / %m (%p%)")
            self.progress_bar.setValue(current)
        else:
            if self.progress_bar.maximum() != 0:
                self.progress_bar.setRange(0, 0)
                self.progress_bar.setFormat("Working...")
    
    def on_state_changed(self, state_description):
        """Update state label with current operation state."""
        self.state_label.setText(state_description)
        self.state_label.setVisible(True)
    
    def start_loading(self, message):
        """Show progress bar and abort button."""
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Starting...")
        self.progress_bar.setVisible(True)
        self.btn_abort.setVisible(True)
        self.state_label.setText("Starting...")
        self.state_label.setVisible(True)
        self.set_buttons_enabled(False)
        self.status_text.append(f"\n{message}")
    
    def stop_loading(self, success, message):
        """Hide progress bar and abort button."""
        self.progress_bar.setVisible(False)
        self.btn_abort.setVisible(False)
        self.state_label.setVisible(False)
        self.set_buttons_enabled(True)
        self.worker = None
        if success:
            self.status_text.append(f"✓ {message}")
        else:
            self.status_text.append(f"✗ {message}")
    
    def append_log(self, text):
        """Append text to log."""
        self.status_text.append(text)
    
    def abort_operation(self):
        """Abort the current operation."""
        if self.worker:
            self.status_text.append("Aborting operation...")
            self.worker.abort()
    
    def browse_json_file(self):
        """Browse for JSON file."""
        file_path = QFileDialog.getOpenFileName(self, "Select JSON File", filter="*.json")[0]
        if file_path:
            self.json_path.setText(file_path)
    
    def browse_marker_file(self):
        """Browse for Marker.sldprt file."""
        file_path = QFileDialog.getOpenFileName(self, "Select Marker Part", filter="*.sldprt")[0]
        if file_path:
            self.marker_path.setText(file_path)
    
    def insert_hardpoints(self):
        """Insert all hardpoints including wheels using automatic paths."""
        json_path = os.path.join(os.path.dirname(__file__), "temp", "Front_Suspension.json")
        marker_path = os.path.join(os.path.dirname(__file__), "Marker.SLDPRT")
        
        if not os.path.exists(json_path):
            QMessageBox.warning(self, "Missing File", "Front_Suspension.json not found in /temp. Parse an Excel file first.")
            return
            
        if not os.path.exists(marker_path):
            QMessageBox.critical(self, "Missing Marker", "Marker.SLDPRT not found in plugin root folder")
            return
        
        self.start_loading("Inserting all hardpoints (including wheels) from latest Front_Suspension.json...")
        
        # Use the hardpoint runner for comprehensive hardpoint insertion
        try:
            exe_path = self._get_latest_suspension_tools_exe()
            
            # Run the hardpoint runner with add command for comprehensive insertion
            cmd = [exe_path, "hardpoints", "add", json_path, marker_path]
            
            self.worker = HardpointWorker(self.run_hardpoint_command, cmd)
            self.worker.log.connect(self.append_log)
            self.worker.progress.connect(self.on_progress)
            self.worker.state_changed.connect(self.on_state_changed)
            self.worker.finished.connect(self.stop_loading)
            self.worker.start()
            
        except Exception as e:
            self.status_text.append(f"✗ Error: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to run hardpoint runner: {str(e)}")

    def _get_latest_suspension_tools_exe(self):
        """Return the newest available SuspensionTools executable (Release/Debug)."""
        sw_drawer_path = os.path.join(os.path.dirname(__file__), "sw_drawer")
        candidates = [
            os.path.join(sw_drawer_path, "bin", "Release", "net48", "SuspensionTools.exe"),
            os.path.join(sw_drawer_path, "bin", "Debug", "net48", "SuspensionTools.exe"),
            os.path.join(sw_drawer_path, "bin", "Release", "net48", "sw_drawer.exe"),
            os.path.join(sw_drawer_path, "bin", "Debug", "net48", "sw_drawer.exe"),
        ]

        existing = [p for p in candidates if os.path.exists(p)]
        if not existing:
            raise FileNotFoundError("SuspensionTools.exe not found. Build the project first.")

        latest = max(existing, key=lambda p: os.path.getmtime(p))
        self.status_text.append(f"Using executable: {latest}")
        return latest
    
    def run_hardpoint_command(self, cmd, worker=None):
        """Run a hardpoint command using subprocess."""
        try:
            # Run the command
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=True,
                bufsize=1
            )
            
            # Store process reference for abort functionality
            if worker:
                worker._process = process
            
            # Read output line by line
            for line in iter(process.stdout.readline, ''):
                line = line.strip()
                if line:
                    # Parse special output lines for progress and state
                    if line.startswith("TOTAL:"):
                        try:
                            new_total = int(line.split(":")[1])
                            if worker:
                                worker._total_tasks = new_total
                                worker.progress.emit(worker._current_progress, worker._total_tasks)
                        except:
                            pass
                    elif line.startswith("PROGRESS:"):
                        try:
                            current = int(line.split(":")[1])
                            if worker:
                                worker._current_progress = current
                                worker.progress.emit(worker._current_progress, worker._total_tasks)
                        except:
                            pass
                    elif line.startswith("STATE:"):
                        state = line.split(":")[1]
                        if worker:
                            worker.state_changed.emit(state)
                    else:
                        # Normal log line
                        if worker:
                            worker.log.emit(line)
            
            # Wait for completion
            process.wait()
            
            if process.returncode == 0:
                return True
            else:
                raise Exception(f"Command failed with exit code {process.returncode}")
                
        except Exception as e:
            if worker:
                worker.log.emit(f"Error: {str(e)}")
            raise e

    def insert_wheel_coordinates(self):
        """Insert wheel coordinates using automatic paths."""
        json_path = os.path.join(os.path.dirname(__file__), "temp", "Front_Suspension.json")
        marker_path = os.path.join(os.path.dirname(__file__), "Marker.SLDPRT")
        
        if not os.path.exists(json_path):
            QMessageBox.warning(self, "Missing File", "Front_Suspension.json not found in /temp. Parse an Excel file first.")
            return
            
        if not os.path.exists(marker_path):
            QMessageBox.critical(self, "Missing Marker", "Marker.SLDPRT not found in plugin root folder")
            return
        
        self.start_loading("Inserting wheel coordinates from latest Front_Suspension.json...")
        
        # Use the hardpoint runner for wheel insertion
        try:
            exe_path = self._get_latest_suspension_tools_exe()
            
            # Run the hardpoint runner with addwheels command
            cmd = [exe_path, "hardpoints", "addwheels", json_path, marker_path]
            
            self.worker = HardpointWorker(self.run_hardpoint_command, cmd)
            self.worker.log.connect(self.append_log)
            self.worker.progress.connect(self.on_progress)
            self.worker.state_changed.connect(self.on_state_changed)
            self.worker.finished.connect(self.stop_loading)
            self.worker.start()
            
        except Exception as e:
            self.status_text.append(f"✗ Error: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to run hardpoint runner: {str(e)}")
    
    def create_coordinates_folder(self):
        """Create the coordinates folder."""
        try:
            folder_path = create_coordinates_folder()
            self.status_text.append(f"✓ Created coordinates folder: {folder_path}")
            QMessageBox.information(self, "Success", f"Created coordinates folder: {folder_path}")
        except Exception as e:
            self.status_text.append(f"✗ Error creating folder: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to create folder: {str(e)}")


class PoseCreationTab(QWidget):
    """Tab for creating poses with coordinate systems and mates."""
    
    def __init__(self):
        super().__init__()
        self.worker = None
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        layout.addWidget(QLabel("Write Pose"))
        layout.addWidget(QLabel("Create coordinate systems and mates for pose configurations"))
        
        # State label
        self.state_label = QLabel("")
        self.state_label.setStyleSheet("font-weight: bold; color: #0066cc;")
        self.state_label.setVisible(False)
        layout.addWidget(self.state_label)
        
        # Progress bar and abort button (hidden by default)
        h_progress = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFormat("%v / %m (%p%)")
        h_progress.addWidget(self.progress_bar)
        
        self.btn_abort = QPushButton("Abort")
        self.btn_abort.setVisible(False)
        self.btn_abort.setStyleSheet("background-color: #ff6666;")
        self.btn_abort.clicked.connect(self.abort_operation)
        h_progress.addWidget(self.btn_abort)
        
        layout.addLayout(h_progress)
        
        # Pose name input
        group_config = QGroupBox("Configuration")
        h_config = QHBoxLayout()

        # Pose name input with validation
        h_config.addWidget(QLabel("Pose Name:"))
        self.pose_name = QLineEdit()
        self.pose_name.setPlaceholderText("Front, Rear, Test...")
        self.pose_name.setToolTip("Enter pose name without special characters")
        h_config.addWidget(self.pose_name)

        # Transform folder suffix
        h_config.addWidget(QLabel("Transform Folder:"))
        self.transform_suffix = QLineEdit()
        self.transform_suffix.setText(" Transforms")
        self.transform_suffix.setEnabled(False)
        h_config.addWidget(self.transform_suffix)

        group_config.setLayout(h_config)
        layout.addWidget(group_config)
        
        # Operations
        group_ops = QGroupBox("Operations")
        h_ops = QHBoxLayout()
        
        self.btn_create_pose = QPushButton("Create Pose")
        self.btn_create_pose.clicked.connect(self.create_pose)
        h_ops.addWidget(self.btn_create_pose)
        
        self.btn_create_transforms_folder = QPushButton("Create Transforms Folder")
        self.btn_create_transforms_folder.clicked.connect(self.create_transforms_folder)
        h_ops.addWidget(self.btn_create_transforms_folder)
        
        group_ops.setLayout(h_ops)
        layout.addWidget(group_ops)
        
        # Existing poses
        group_poses = QGroupBox("Existing Poses")
        layout_poses = QVBoxLayout()
        
        self.pose_list = QTextEdit()
        self.pose_list.setReadOnly(True)
        self.pose_list.setMaximumHeight(100)
        layout_poses.addWidget(self.pose_list)
        
        btn_refresh_poses = QPushButton("Refresh Poses")
        btn_refresh_poses.clicked.connect(self.refresh_poses)
        layout_poses.addWidget(btn_refresh_poses)
        
        group_poses.setLayout(layout_poses)
        layout.addWidget(group_poses)
        
        # Console output
        layout.addWidget(QLabel("Output:"))
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setFontFamily("Courier")
        self.status_text.setMaximumHeight(150)
        self.status_text.setText("Ready")
        layout.addWidget(self.status_text)
        
        btn_clear = QPushButton("Clear Log")
        btn_clear.clicked.connect(lambda: self.status_text.clear())
        layout.addWidget(btn_clear)
        
        layout.addStretch()
        self.setLayout(layout)
        self.refresh_poses()
    
    def set_buttons_enabled(self, enabled):
        """Enable or disable all action buttons."""
        self.btn_create_pose.setEnabled(enabled)
        self.btn_create_transforms_folder.setEnabled(enabled)
    
    def on_progress(self, current, total):
        """Update progress bar with current/total values."""
        if total > 0:
            if self.progress_bar.maximum() != total:
                self.progress_bar.setRange(0, total)
                self.progress_bar.setFormat("%v / %m (%p%)")
            self.progress_bar.setValue(current)
        else:
            if self.progress_bar.maximum() != 0:
                self.progress_bar.setRange(0, 0)
                self.progress_bar.setFormat("Working...")
    
    def on_state_changed(self, state_description):
        """Update state label with current operation state."""
        self.state_label.setText(state_description)
        self.state_label.setVisible(True)
    
    def start_loading(self, message):
        """Show progress bar and abort button."""
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Starting...")
        self.progress_bar.setVisible(True)
        self.btn_abort.setVisible(True)
        self.state_label.setText("Starting...")
        self.state_label.setVisible(True)
        self.set_buttons_enabled(False)
        self.status_text.append(f"\n{message}")
    
    def stop_loading(self, success, message):
        """Hide progress bar and abort button."""
        self.progress_bar.setVisible(False)
        self.btn_abort.setVisible(False)
        self.state_label.setVisible(False)
        self.set_buttons_enabled(True)
        self.worker = None
        if success:
            self.status_text.append(f"✓ {message}")
            self.refresh_poses()
        else:
            self.status_text.append(f"✗ {message}")
    
    def append_log(self, text):
        """Append text to log."""
        self.status_text.append(text)
    
    def abort_operation(self):
        """Abort the current operation."""
        if self.worker:
            self.status_text.append("Aborting operation...")
            self.worker.abort()
    
    def browse_json_file(self):
        """Browse for JSON file."""
        file_path = QFileDialog.getOpenFileName(self, "Select JSON File", filter="*.json")[0]
        if file_path:
            self.json_path.setText(file_path)
    
    def create_pose(self):
        """Create pose using pose creation module."""
        json_path = self.json_path.text().strip()
        pose_name = self.pose_name.text().strip()
        transform_folder = f"{pose_name}{self.transform_suffix.text().strip()}"

        if not pose_name:
            QMessageBox.warning(self, "Warning", "Pose name cannot be empty")
            return
            
        if not re.match(r"^[a-zA-Z0-9_\- ]+$", pose_name):
            QMessageBox.warning(self, "Invalid Name", 
                "Pose name can only contain letters, numbers, spaces, hyphens and underscores")
            return
        
        if not json_path or not pose_name:
            QMessageBox.warning(self, "Warning", "Please select JSON file and enter pose name")
            return
        
        # Validate pose name
        valid, message = validate_pose_name(pose_name)
        if not valid:
            QMessageBox.warning(self, "Invalid Pose Name", message)
            return
        
        # Validate file
        if not os.path.exists(json_path):
            QMessageBox.warning(self, "Warning", f"JSON file not found: {json_path}")
            return
        
        self.start_loading(f"Creating pose '{pose_name}' from {os.path.basename(json_path)}...")
        
        self.worker = PoseCreationWorker(create_pose, json_path, pose_name)
        self.worker.log.connect(self.append_log)
        self.worker.progress.connect(self.on_progress)
        self.worker.state_changed.connect(self.on_state_changed)
        self.worker.finished.connect(self.stop_loading)
        self.worker.start()
    
    def create_transforms_folder(self):
        """Create the transforms folder for a pose."""
        pose_name = self.pose_name.text().strip()
        if not pose_name:
            QMessageBox.warning(self, "Warning", "Please enter a pose name")
            return
        
        try:
            from pose_creation import create_pose_folder
            folder_path = create_pose_folder(pose_name)
            self.status_text.append(f"✓ Created transforms folder: {folder_path}")
            QMessageBox.information(self, "Success", f"Created transforms folder: {folder_path}")
        except Exception as e:
            self.status_text.append(f"✗ Error creating folder: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to create folder: {str(e)}")
    
    def refresh_poses(self):
        """Refresh the list of existing poses."""
        try:
            poses = get_existing_poses()
            if poses:
                self.pose_list.setText("\n".join(poses))
            else:
                self.pose_list.setText("No poses found")
        except Exception as e:
            self.pose_list.setText(f"Error loading poses: {str(e)}")


class VisualizationControlTab(QWidget):
    """Tab for controlling visualization with color-coded buttons."""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        layout.addWidget(QLabel("Visualization Controls"))
        layout.addWidget(QLabel("Control visibility of suspension components with color coding"))
        
        # Toggle all controls
        group_all = QGroupBox("Toggle All")
        h_all = QHBoxLayout()
        
        btn_show_all = QPushButton("Show All")
        btn_show_all.clicked.connect(lambda: self.set_suspension_visibility('all', True))
        h_all.addWidget(btn_show_all)
        
        btn_hide_all = QPushButton("Hide All")
        btn_hide_all.clicked.connect(lambda: self.set_suspension_visibility('all', False))
        h_all.addWidget(btn_hide_all)
        
        group_all.setLayout(h_all)
        layout.addWidget(group_all)
        
        # Front/Rear controls
        group_front_rear = QGroupBox("Front / Rear")
        h_front_rear = QHBoxLayout()
        
        btn_show_front = QPushButton("Show Front")
        btn_show_front.clicked.connect(lambda: self.set_suspension_visibility('front', True))
        h_front_rear.addWidget(btn_show_front)
        
        btn_hide_front = QPushButton("Hide Front")
        btn_hide_front.clicked.connect(lambda: self.set_suspension_visibility('front', False))
        h_front_rear.addWidget(btn_hide_front)
        
        btn_show_rear = QPushButton("Show Rear")
        btn_show_rear.clicked.connect(lambda: self.set_suspension_visibility('rear', True))
        h_front_rear.addWidget(btn_show_rear)
        
        btn_hide_rear = QPushButton("Hide Rear")
        btn_hide_rear.clicked.connect(lambda: self.set_suspension_visibility('rear', False))
        h_front_rear.addWidget(btn_hide_rear)
        
        group_front_rear.setLayout(h_front_rear)
        layout.addWidget(group_front_rear)
        
        # Color-coded category controls
        group_categories = QGroupBox("By Category (Color Coded)")
        layout_categories = QVBoxLayout()
        
        color_info = get_color_coding_info()
        for prefix, info in color_info.items():
            h_category = QHBoxLayout()
            
            btn_show = QPushButton(f"Show {info['name']}")
            btn_show.setStyleSheet(f"background-color: rgb({info['rgb'][0]}, {info['rgb'][1]}, {info['rgb'][2]}); color: white;")
            btn_show.clicked.connect(lambda p=prefix: self.set_suspension_visibility('substring', True, p))
            h_category.addWidget(btn_show)
            
            btn_hide = QPushButton(f"Hide {info['name']}")
            btn_hide.setStyleSheet(f"background-color: rgb({info['rgb'][0]}, {info['rgb'][1]}, {info['rgb'][2]}); color: white;")
            btn_hide.clicked.connect(lambda p=prefix: self.set_suspension_visibility('substring', False, p))
            h_category.addWidget(btn_hide)
            
            layout_categories.addLayout(h_category)
        
        group_categories.setLayout(layout_categories)
        layout.addWidget(group_categories)
        
        # Marker controls
        group_markers = QGroupBox("Marker Controls")
        layout_markers = QVBoxLayout()
        
        h_markers_all = QHBoxLayout()
        btn_show_markers_all = QPushButton("Show All Markers")
        btn_show_markers_all.clicked.connect(lambda: self.set_marker_visibility('all', True))
        h_markers_all.addWidget(btn_show_markers_all)
        
        btn_hide_markers_all = QPushButton("Hide All Markers")
        btn_hide_markers_all.clicked.connect(lambda: self.set_marker_visibility('all', False))
        h_markers_all.addWidget(btn_hide_markers_all)
        layout_markers.addLayout(h_markers_all)
        
        h_markers_front_rear = QHBoxLayout()
        btn_show_markers_front = QPushButton("Show Front Markers")
        btn_show_markers_front.clicked.connect(lambda: self.set_marker_visibility('front', True))
        h_markers_front_rear.addWidget(btn_show_markers_front)
        
        btn_hide_markers_front = QPushButton("Hide Front Markers")
        btn_hide_markers_front.clicked.connect(lambda: self.set_marker_visibility('front', False))
        h_markers_front_rear.addWidget(btn_hide_markers_front)
        
        btn_show_markers_rear = QPushButton("Show Rear Markers")
        btn_show_markers_rear.clicked.connect(lambda: self.set_marker_visibility('rear', True))
        h_markers_front_rear.addWidget(btn_show_markers_rear)
        
        btn_hide_markers_rear = QPushButton("Hide Rear Markers")
        btn_hide_markers_rear.clicked.connect(lambda: self.set_marker_visibility('rear', False))
        h_markers_front_rear.addWidget(btn_hide_markers_rear)
        layout_markers.addLayout(h_markers_front_rear)
        
        group_markers.setLayout(layout_markers)
        layout.addWidget(group_markers)
        
        # Console output
        layout.addWidget(QLabel("Output:"))
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setFontFamily("Courier")
        self.status_text.setMaximumHeight(100)
        self.status_text.setText("Ready")
        layout.addWidget(self.status_text)
        
        btn_clear = QPushButton("Clear Log")
        btn_clear.clicked.connect(lambda: self.status_text.clear())
        layout.addWidget(btn_clear)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def set_suspension_visibility(self, target, visible):
        """Set suspension visibility."""
        try:
            success = set_suspension_visibility(target, visible)
            action = "Showing" if visible else "Hiding"
            self.status_text.append(f"{action} {target}... {'✓ Done' if success else '✗ Failed'}")
        except Exception as e:
            self.status_text.append(f"✗ Error: {str(e)}")
    
    def set_marker_visibility(self, target, visible):
        """Set marker visibility."""
        try:
            success = set_marker_visibility(target, visible)
            action = "Showing" if visible else "Hiding"
            self.status_text.append(f"{action} {target}... {'✓ Done' if success else '✗ Failed'}")
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
        self.setGeometry(100, 100, 700, 700)
        
        tabs = QTabWidget()
        
        tabs.addTab(ImportOptimumKTab(), "Parse")
        tabs.addTab(CoordinateInsertionTab(), "Insert Coordinates") 
        tabs.addTab(PoseCreationTab(), "Write Pose")
        tabs.addTab(VisualizationControlTab(), "Visualization")
        tabs.addTab(HelpTab(), "Help")
        
        self.setCentralWidget(tabs)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OptimumKApp()
    window.show()
    sys.exit(app.exec_())