import os
import sys
import json
import subprocess
from PyQt5.QtCore import QThread, pyqtSignal, QObject

class PoseCreationWorker(QThread):
    """Worker thread for pose creation operations."""
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
                "Initializing": "Starting pose creation...",
                "LoadingJson": "Loading JSON data...",
                "LoadingMarkerPart": "Loading marker part...",
                "CreatingCoordinateSystems": "Creating coordinate systems...",
                "CreatingTransformsFolder": "Creating Transforms folder...",
                "CreatingTransforms": "Creating transform features...",
                "Rebuilding": "Rebuilding model...",
                "Complete": "Pose creation complete"
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
                self.finished.emit(True, "Pose creation completed successfully")
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
    
    # Check multiple possible locations
    paths_to_check = [
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "net48", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug", "net48", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "net6.0", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug", "net6.0", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "net8.0", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug", "net8.0", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug", "sw_drawer.exe"),
    ]
    
    for path in paths_to_check:
        if os.path.exists(path):
            return path
    
    raise FileNotFoundError(
        "sw_drawer.exe not found. Run 'dotnet build -c Release' in the sw_drawer folder first.\n"
        f"Searched in: {paths_to_check[0]}"
    )


def create_pose(json_path, pose_name, worker=None, progress_callback=None):
    """Create pose from JSON file with coordinate systems and mates."""
    if not os.path.exists(json_path):
        raise FileNotFoundError(f"JSON file not found at {json_path}")
    
    exe_path = get_suspension_tools_exe()
    args = [exe_path, "hardpoints", "pose", json_path, pose_name]
    
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
            print("Operation aborted by user")
            return False
        elif progress_callback and hasattr(progress_callback, '_abort') and progress_callback._abort:
            process.terminate()
            print("Operation aborted by user")
            return False
        print(line.rstrip())
    
    process.wait()
    return process.returncode == 0


def load_json(file_path):
    """Load JSON file."""
    with open(file_path, 'r') as f:
        return json.load(f)


def extract_hardpoints_for_pose(json_data):
    """Extract hardpoints for pose creation."""
    hardpoints = []
    
    # Extract from Front Suspension
    if json_data.get("Front Suspension"):
        extract_hardpoints_from_section(json_data["Front Suspension"], "_FRONT", hardpoints)
    
    # Extract from Rear Suspension  
    if json_data.get("Rear Suspension"):
        extract_hardpoints_from_section(json_data["Rear Suspension"], "_REAR", hardpoints)
    
    # Extract wheels
    if json_data.get("Wheels"):
        extract_wheels_for_pose(json_data["Wheels"], hardpoints)
    
    return hardpoints


def extract_hardpoints_from_section(section, suffix, hardpoints):
    """Extract hardpoints from a suspension section for pose."""
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


def extract_wheels_for_pose(wheels_data, hardpoints):
    """Extract wheel hardpoints for pose."""
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


def validate_pose_files(json_path):
    """Validate that required files exist for pose creation."""
    errors = []
    
    if not os.path.exists(json_path):
        errors.append(f"JSON file not found: {json_path}")
    
    return errors


def create_pose_folder(pose_name):
    """Create pose folder if it doesn't exist."""
    pose_dir = os.path.join(os.path.dirname(__file__), "poses", pose_name)
    os.makedirs(pose_dir, exist_ok=True)
    return pose_dir


def get_coordinate_system_name(pose_name, coordinate_name):
    """Get the coordinate system name for a pose."""
    return f"{pose_name} {coordinate_name}"


def get_transforms_folder_name(pose_name):
    """Get the transforms folder name for a pose."""
    return f"{pose_name} Transforms"


def validate_pose_name(pose_name):
    """Validate pose name format."""
    if not pose_name or not pose_name.strip():
        return False, "Pose name cannot be empty"
    
    # Check for invalid characters
    invalid_chars = ['<', '>', ':', '"', '|', '?', '*']
    for char in invalid_chars:
        if char in pose_name:
            return False, f"Pose name cannot contain '{char}'"
    
    return True, "Valid pose name"


def get_existing_poses():
    """Get list of existing poses from the poses directory."""
    poses_dir = os.path.join(os.path.dirname(__file__), "poses")
    if not os.path.exists(poses_dir):
        return []
    
    try:
        return [d for d in os.listdir(poses_dir) if os.path.isdir(os.path.join(poses_dir, d))]
    except:
        return []


def delete_pose_folder(pose_name):
    """Delete a pose folder and its contents."""
    import shutil
    pose_dir = os.path.join(os.path.dirname(__file__), "poses", pose_name)
    if os.path.exists(pose_dir):
        try:
            shutil.rmtree(pose_dir)
            return True
        except Exception as e:
            print(f"Error deleting pose folder: {e}")
            return False
    return True