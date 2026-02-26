import os
import sys
import subprocess
from PyQt5.QtCore import QThread, pyqtSignal, QObject

class VisualizationWorker(QThread):
    """Worker thread for visualization operations."""
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
                "Initializing": "Starting visualization control...",
                "LoadingJson": "Loading JSON data...",
                "LoadingMarkerPart": "Loading marker part...",
                "UpdatingVisibility": "Updating visibility...",
                "Rebuilding": "Rebuilding model...",
                "Complete": "Visualization control complete"
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
                self.finished.emit(True, "Visualization control completed successfully")
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


def set_suspension_visibility(target, visible, filter_text=None):
    """Set suspension visibility using SuspensionTools.exe."""
    exe_path = get_suspension_tools_exe()
    
    # Map target to command
    command_map = {
        'all': 'all',
        'front': 'front',
        'rear': 'rear',
        'front_wheels': 'frontwheels',
        'rear_wheels': 'rearwheels',
        'wheels': 'wheels',
        'chassis': 'chassis',
        'non_chassis': 'nonchassis',
        'substring': 'substring'
    }
    
    command = command_map.get(target, 'all')
    vis_str = 'show' if visible else 'hide'
    
    args = [exe_path, 'vis', command, vis_str]
    if filter_text:
        args.append(filter_text)
    
    print(f"Running: {' '.join(args)}")
    
    result = subprocess.run(args, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout.strip())
    if result.returncode != 0 and result.stderr:
        print(result.stderr.strip())
    
    return result.returncode == 0


def set_marker_visibility(target, visible, filter_text=None):
    """Set marker visibility using SuspensionTools.exe."""
    exe_path = get_suspension_tools_exe()
    
    # Map target to command
    command_map = {
        'all': 'all',
        'front': 'front',
        'rear': 'rear',
        'name': 'name'
    }
    
    command = command_map.get(target, 'all')
    vis_str = 'show' if visible else 'hide'
    
    args = [exe_path, 'marker', 'vis', command, vis_str]
    if filter_text:
        args.append(filter_text)
    
    print(f"Running: {' '.join(args)}")
    
    result = subprocess.run(args, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout.strip())
    if result.returncode != 0 and result.stderr:
        print(result.stderr.strip())
    
    return result.returncode == 0


def set_feature_visibility(feature_name, visible):
    """Set specific feature visibility."""
    exe_path = get_suspension_tools_exe()
    vis_str = 'show' if visible else 'hide'
    
    args = [exe_path, 'vis', 'feature', vis_str, feature_name]
    
    print(f"Running: {' '.join(args)}")
    
    result = subprocess.run(args, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout.strip())
    if result.returncode != 0 and result.stderr:
        print(result.stderr.strip())
    
    return result.returncode == 0


def get_color_coding_info():
    """Get information about color coding for visualization."""
    color_info = {
        'CHAS_': {'name': 'Chassis', 'rgb': [255, 0, 0], 'description': 'Chassis pickup points'},
        'wheel': {'name': 'Wheels', 'rgb': [0, 255, 0], 'description': 'Wheel components'},
        'UPRI_': {'name': 'Upright', 'rgb': [0, 0, 255], 'description': 'Upright components'},
        'Other': {'name': 'Other', 'rgb': [128, 0, 128], 'description': 'Other components'}
    }
    return color_info


def get_visualization_controls():
    """Get list of available visualization controls."""
    controls = {
        'suspension': [
            {'name': 'All Suspension', 'target': 'all', 'description': 'Show/hide all suspension components'},
            {'name': 'Front Suspension', 'target': 'front', 'description': 'Show/hide front suspension only'},
            {'name': 'Rear Suspension', 'target': 'rear', 'description': 'Show/hide rear suspension only'},
            {'name': 'All Wheels', 'target': 'wheels', 'description': 'Show/hide all wheel components'},
            {'name': 'Front Wheels', 'target': 'front_wheels', 'description': 'Show/hide front wheels only'},
            {'name': 'Rear Wheels', 'target': 'rear_wheels', 'description': 'Show/hide rear wheels only'},
            {'name': 'Chassis Points', 'target': 'chassis', 'description': 'Show/hide chassis pickup points'},
            {'name': 'Non-Chassis Points', 'target': 'non_chassis', 'description': 'Show/hide non-chassis suspension points'}
        ],
        'markers': [
            {'name': 'All Markers', 'target': 'all', 'description': 'Show/hide all marker components'},
            {'name': 'Front Markers', 'target': 'front', 'description': 'Show/hide front suspension markers'},
            {'name': 'Rear Markers', 'target': 'rear', 'description': 'Show/hide rear suspension markers'},
            {'name': 'By Name', 'target': 'name', 'description': 'Show/hide markers by name pattern'}
        ],
        'categories': [
            {'name': 'Chassis', 'target': 'substring', 'filter': 'CHAS_', 'description': 'Show/hide chassis components'},
            {'name': 'Upright', 'target': 'substring', 'filter': 'UPRI_', 'description': 'Show/hide upright components'},
            {'name': 'Rocker', 'target': 'substring', 'filter': 'ROCK_', 'description': 'Show/hide rocker components'},
            {'name': 'Non-Sprung Mass', 'target': 'substring', 'filter': 'NSMA_', 'description': 'Show/hide non-sprung mass components'}
        ]
    }
    return controls


def validate_solidworks_connection():
    """Validate that SolidWorks is running and accessible."""
    try:
        import win32com.client
        swApp = win32com.client.GetActiveObject("SldWorks.Application")
        if swApp:
            return True, "SolidWorks is running"
        else:
            return False, "SolidWorks is not running"
    except:
        return False, "Could not connect to SolidWorks"


def get_active_document_info():
    """Get information about the active SolidWorks document."""
    try:
        import win32com.client
        swApp = win32com.client.GetActiveObject("SldWorks.Application")
        swModel = swApp.ActiveDoc
        
        if swModel:
            doc_type = swModel.GetType()
            doc_types = {
                1: "Part",
                2: "Assembly", 
                3: "Drawing"
            }
            
            return {
                'title': swModel.GetTitle(),
                'path': swModel.GetPathName(),
                'type': doc_types.get(doc_type, f"Unknown ({doc_type})"),
                'is_assembly': doc_type == 2
            }
        else:
            return None
    except:
        return None


def create_visualization_profile(profile_name, settings):
    """Create a visualization profile for quick access."""
    profiles_dir = os.path.join(os.path.dirname(__file__), "visualization_profiles")
    os.makedirs(profiles_dir, exist_ok=True)
    
    profile_path = os.path.join(profiles_dir, f"{profile_name}.json")
    
    profile_data = {
        'name': profile_name,
        'settings': settings,
        'created': __import__('datetime').datetime.now().isoformat()
    }
    
    import json
    with open(profile_path, 'w') as f:
        json.dump(profile_data, f, indent=2)
    
    return profile_path


def load_visualization_profile(profile_name):
    """Load a visualization profile."""
    profiles_dir = os.path.join(os.path.dirname(__file__), "visualization_profiles")
    profile_path = os.path.join(profiles_dir, f"{profile_name}.json")
    
    if os.path.exists(profile_path):
        import json
        with open(profile_path, 'r') as f:
            return json.load(f)
    return None


def get_available_profiles():
    """Get list of available visualization profiles."""
    profiles_dir = os.path.join(os.path.dirname(__file__), "visualization_profiles")
    if not os.path.exists(profiles_dir):
        return []
    
    try:
        profiles = []
        for filename in os.listdir(profiles_dir):
            if filename.endswith('.json'):
                profile_name = filename[:-5]  # Remove .json extension
                profiles.append(profile_name)
        return profiles
    except:
        return []


def apply_visualization_profile(profile_name):
    """Apply a visualization profile."""
    profile_data = load_visualization_profile(profile_name)
    if not profile_data:
        return False, f"Profile '{profile_name}' not found"
    
    settings = profile_data.get('settings', {})
    
    # Apply suspension settings
    suspension_settings = settings.get('suspension', {})
    for target, visible in suspension_settings.items():
        set_suspension_visibility(target, visible)
    
    # Apply marker settings
    marker_settings = settings.get('markers', {})
    for target, visible in marker_settings.items():
        set_marker_visibility(target, visible)
    
    # Apply category settings
    category_settings = settings.get('categories', {})
    for category, visible in category_settings.items():
        set_suspension_visibility('substring', visible, category)
    
    return True, f"Applied profile '{profile_name}'"