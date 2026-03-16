import os
import json
import subprocess
from PyQt5.QtWidgets import QMessageBox
from workers import WorkerBase
from utils import find_suspension_tools_exe
from solidworks_release import release_solidworks_command_state


class CoordinateInsertionWorker(WorkerBase):
    """Worker thread for coordinate insertion operations."""
    SUCCESS_MESSAGE = "Coordinate insertion completed successfully"
    STATE_DESCRIPTIONS = {
        "Initializing": "Starting coordinate insertion...",
        "LoadingJson": "Loading JSON data...",
        "LoadingMarkerPart": "Loading marker part...",
        "InsertingBodies": "Inserting marker bodies...",
        "RenamingBodies": "Renaming bodies...",
        "ApplyingColors": "Applying colors...",
        "CreatingCoordinateSystems": "Creating coordinate systems...",
        "CreatingHardpointsFolder": "Creating Hardpoints folder...",
        "Rebuilding": "Rebuilding model...",
        "Complete": "Coordinate insertion complete",
    }


def insert_coordinates(json_path, marker_path, worker=None, progress_callback=None):
    """Insert coordinates from JSON file with color coding and naming."""
    if not os.path.exists(json_path):
        raise FileNotFoundError(f"JSON file not found at {json_path}")
    
    if not os.path.exists(marker_path):
        raise FileNotFoundError(f"Marker part not found at {marker_path}")
    
    exe_path = find_suspension_tools_exe()
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