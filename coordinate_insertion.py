import os
import sys
import json
import subprocess
from PyQt5.QtCore import QThread, pyqtSignal, QObject
from PyQt5.QtWidgets import QMessageBox
from solidworks_release import release_solidworks_command_state

class CoordinateInsertionWorker(QThread):
    """Worker thread for coordinate insertion operations."""
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

        if self._process is not None:
            released, release_message = release_solidworks_command_state()
            if not released:
                print(f"Warning: Failed to release SolidWorks state after abort: {release_message}")

    def parse_output_line(self, line):
        """Parse special output lines for progress and state."""
        if line.startswith("TOTAL:"):
            try:
                new_total = int(line.split(":")[1])
                self._total_tasks = new_total
                self.progress.emit(self._current_progress, self._total_tasks)
            except:
                pass
            return None
        elif line.startswith("PROGRESS:"):
            try:
                self._current_progress = int(line.split(":")[1])
                self.progress.emit(self._current_progress, self._total_tasks)
            except:
                pass
            return None
        elif line.startswith("STATE:"):
            state = line.split(":")[1]
            state_descriptions = {
                "Initializing": "Starting coordinate insertion...",
                "LoadingJson": "Loading JSON data...",
                "LoadingMarkerPart": "Loading marker part...",
                "InsertingBodies": "Inserting marker bodies...",
                "RenamingBodies": "Renaming bodies...",
                "ApplyingColors": "Applying colors...",
                "CreatingCoordinateSystems": "Creating coordinate systems...",
                "CreatingHardpointsFolder": "Creating Hardpoints folder...",
                "Rebuilding": "Rebuilding model...",
                "Complete": "Coordinate insertion complete"
            }
            description = state_descriptions.get(state, state)
            self.state_changed.emit(description)
            return None
        return line

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
                self.finished.emit(True, "Coordinate insertion completed successfully")
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


class QtStream(QObject):
    """Redirect stdout to a PyQt signal."""
    text_written = pyqtSignal(str)

    def write(self, text):
        if text.strip():
            self.text_written.emit(text.strip())

    def flush(self):
        pass


def get_suspension_tools_exe():
    """Get the path to SuspensionTools.exe"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Check multiple possible locations - PRIORITIZE RELEASE BUILD
    paths_to_check = [
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "net48", "SuspensionTools.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "net48", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug", "net48", "SuspensionTools.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug", "net48", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "SuspensionTools.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "sw_drawer.exe"),
    ]
    
    for path in paths_to_check:
        if os.path.exists(path):
            return path
    
    raise FileNotFoundError(
        "sw_drawer.exe not found. Run 'dotnet build -c Release' in the sw_drawer folder first.\n"
        f"Searched in: {paths_to_check[0]}"
    )


def insert_coordinates(json_path, marker_path, worker=None, progress_callback=None):
    """Insert coordinates from JSON file with color coding and naming."""
    if not os.path.exists(json_path):
        raise FileNotFoundError(f"JSON file not found at {json_path}")
    
    if not os.path.exists(marker_path):
        raise FileNotFoundError(f"Marker part not found at {marker_path}")
    
    exe_path = get_suspension_tools_exe()
    args = [exe_path, "hardpoints", "add", json_path, marker_path]
    
    print(f"Running: {' '.join(args)}")
    
    process = subprocess.Popen(
        args,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1
    )
    
    # Store process in worker for abort capability
    if worker:
        worker._process = process
    elif progress_callback and hasattr(progress_callback, '_process'):
        progress_callback._process = process
    
    # Stream output line by line
    for line in process.stdout:
        if worker and worker._abort:
            process.terminate()
            released, release_message = release_solidworks_command_state()
            if not released:
                print(f"Warning: Failed to release SolidWorks state after abort: {release_message}")
            print("Operation aborted by user")
            return False
        elif progress_callback and hasattr(progress_callback, '_abort') and progress_callback._abort:
            process.terminate()
            released, release_message = release_solidworks_command_state()
            if not released:
                print(f"Warning: Failed to release SolidWorks state after abort: {release_message}")
            print("Operation aborted by user")
            return False
        print(line.rstrip())
    
    process.wait()
    return process.returncode == 0


def load_json(file_path):
    """Load JSON file."""
    with open(file_path, 'r') as f:
        return json.load(f)


def extract_hardpoints(json_data):
    """Extract hardpoints from JSON data."""
    hardpoints = []
    
    # Extract from Front Suspension
    if json_data.get("Front Suspension"):
        extract_hardpoints_from_section(json_data["Front Suspension"], "_FRONT", hardpoints)
    
    # Extract from Rear Suspension  
    if json_data.get("Rear Suspension"):
        extract_hardpoints_from_section(json_data["Rear Suspension"], "_REAR", hardpoints)
    
    # Extract wheels
    if json_data.get("Wheels"):
        extract_wheels(json_data["Wheels"], hardpoints)
    
    return hardpoints


def extract_hardpoints_from_section(section, suffix, hardpoints):
    """Extract hardpoints from a suspension section."""
    if not section:
        return
    
    for section_name, section_obj in section.items():
        if section_name == "Wheels":
            continue
        
        if isinstance(section_obj, dict):
            for point_name, coords in section_obj.items():
                if isinstance(coords, list) and len(coords) >= 3:
                    hardpoints.append({
                        'name': f"{point_name}{suffix}",
                        'x': float(coords[0]),
                        'y': float(coords[1]),
                        'z': float(coords[2]),
                        'base_name': point_name,
                        'suffix': suffix
                    })


def extract_wheels(wheels_data, hardpoints):
    """Extract wheel hardpoints."""
    if not wheels_data:
        return
    
    try:
        half_track = float(wheels_data.get("Half Track", {}).get("left", 0))
        tire_radius = float(wheels_data.get("Tire Diameter", {}).get("left", 0)) / 2.0
        
        # Front wheels
        hardpoints.append({
            'name': 'FL_wheel_FRONT',
            'x': 0,
            'y': half_track,
            'z': tire_radius,
            'base_name': 'FL_wheel',
            'suffix': '_FRONT'
        })
        
        hardpoints.append({
            'name': 'FR_wheel_FRONT',
            'x': 0,
            'y': -half_track,
            'z': tire_radius,
            'base_name': 'FR_wheel',
            'suffix': '_FRONT'
        })
        
        # Rear wheels
        hardpoints.append({
            'name': 'RL_wheel_REAR',
            'x': 0,
            'y': half_track,
            'z': tire_radius,
            'base_name': 'RL_wheel',
            'suffix': '_REAR'
        })
        
        hardpoints.append({
            'name': 'RR_wheel_REAR',
            'x': 0,
            'y': -half_track,
            'z': tire_radius,
            'base_name': 'RR_wheel',
            'suffix': '_REAR'
        })
        
    except (KeyError, ValueError) as e:
        print(f"Warning: Could not extract wheel data: {e}")


def get_color_for_name(name):
    """Get color RGB values based on name prefix."""
    color_map = {
        "CHAS_": [255, 0, 0],       # Red - Chassis
        "UPRI_": [0, 0, 255],       # Blue - Upright
        "ROCK_": [0, 128, 255],     # Light Blue - Rocker
        "NSMA_": [255, 192, 203],   # Pink - Non-Sprung Mass
        "PUSH_": [0, 255, 0],       # Green - Pushrod
        "TIER_": [255, 165, 0],     # Orange - Tie Rod
        "DAMP_": [128, 0, 128],     # Purple - Damper
        "ARBA_": [255, 255, 0],     # Yellow - ARB
        "_FRONT": [0, 200, 200],    # Cyan - Front
        "_REAR": [200, 100, 0],     # Brown - Rear
        "wheel": [64, 64, 64],      # Dark Gray - Wheels
    }
    
    default_color = [128, 128, 128]  # Gray default
    
    upper = name.upper()
    for prefix, color in color_map.items():
        if prefix.upper() in upper:
            return color
    return default_color


def validate_files(json_path, marker_path):
    """Validate that required files exist."""
    errors = []
    
    if not os.path.exists(json_path):
        errors.append(f"JSON file not found: {json_path}")
    
    if not os.path.exists(marker_path):
        errors.append(f"Marker part not found: {marker_path}")
    
    # Check if marker is a .sldprt file
    if marker_path and not marker_path.lower().endswith('.sldprt'):
        errors.append("Marker file must be a .sldprt file")
    
    return errors


def create_coordinates_folder():
    """Create coordinates folder if it doesn't exist."""
    coords_dir = os.path.join(os.path.dirname(__file__), "coordinates")
    os.makedirs(coords_dir, exist_ok=True)
    return coords_dir


def get_marker_path():
    """Get the default marker path."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    marker_paths = [
        os.path.join(script_dir, "Marker.SLDPRT"),
        os.path.join(script_dir, "Marker.sldprt"),
        os.path.join(script_dir, "coordinates", "Marker.SLDPRT"),
        os.path.join(script_dir, "coordinates", "Marker.sldprt"),
    ]
    
    for path in marker_paths:
        if os.path.exists(path):
            return path
    
    # If no marker found, return default path
    return os.path.join(script_dir, "Marker.SLDPRT")