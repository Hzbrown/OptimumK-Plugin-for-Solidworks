import json
import os
import subprocess

SCRIPT_DIR = os.path.dirname(__file__)
CS_EXE = os.path.join(SCRIPT_DIR, "sw_drawer", "bin", "Release", "net48", "CoordinateRunner.exe")


def load_json(file_path):
    """Load JSON file."""
    with open(file_path, 'r') as f:
        return json.load(f)


def insert_coordinate_system(name, x, y, z, angle_x=0.0, angle_y=0.0, angle_z=0.0):
    """Call C# CoordinateRunner.exe via subprocess."""
    if not os.path.exists(CS_EXE):
        raise FileNotFoundError(f"CoordinateRunner.exe not found at {CS_EXE}. Run 'dotnet build -c Release' in sw_drawer folder.")
    args = [CS_EXE, name, str(x), str(y), str(z), str(angle_x), str(angle_y), str(angle_z)]
    result = subprocess.run(args, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout.strip())
    if result.returncode != 0 and result.stderr:
        print(result.stderr.strip())
    return result.returncode == 0


def count_hardpoints(suspension_data: dict) -> int:
    """Count total hardpoints in suspension data."""
    count = 0
    for section_name, section_data in suspension_data.items():
        if section_name == "Wheels":
            continue
        if not isinstance(section_data, dict):
            continue
        for point_name, coords in section_data.items():
            if isinstance(coords, list) and len(coords) >= 3:
                count += 1
    return count


def count_wheels(wheels_data: dict) -> int:
    """Count total wheel coordinate systems (2 per axle)."""
    return 2 if wheels_data else 0


def InsertHardpoint(suspension_data: dict, suffix: str, x_offset: float = 0.0, progress_callback=None):
    """Insert all hardpoints from suspension data as coordinate systems."""
    for section_name, section_data in suspension_data.items():
        if section_name == "Wheels":
            continue
        if not isinstance(section_data, dict):
            continue
        for point_name, coords in section_data.items():
            if not isinstance(coords, list) or len(coords) < 3:
                continue
            try:
                x = float(coords[0]) + float(x_offset)
                y = float(coords[1])
                z = float(coords[2])
                cs_name = f"{point_name}{suffix}"
                insert_coordinate_system(cs_name, x, y, z)
                if progress_callback:
                    progress_callback(1)
            except Exception as e:
                print(f"Error inserting {point_name}{suffix}: {e}")


def InsertWheel(wheels_data: dict, reference_distance: float, is_rear: bool, progress_callback=None):
    """Insert wheel coordinate systems with toe and camber angles."""
    try:
        half_track = float(wheels_data["Half Track"]["left"])
        tire_diameter = float(wheels_data["Tire Diameter"]["left"])
        lateral_offset = float(wheels_data["Lateral Offset"]["left"])
        vertical_offset = float(wheels_data["Vertical Offset"]["left"])
        longitudinal_offset = float(wheels_data["Longitudinal Offset"]["left"])
        camber = float(wheels_data["Static Camber"]["left"])
        toe = float(wheels_data["Static Toe"]["left"])
        
        x_base = float(reference_distance) + longitudinal_offset
        y_base = half_track + lateral_offset 
        z = tire_diameter / 2.0 + vertical_offset
        
        prefix = "R" if is_rear else "F"
        
        # Left wheel (positive Y)
        insert_coordinate_system(f"{prefix}L_wheel", x_base, y_base, z, camber, 0.0, toe)
        if progress_callback:
            progress_callback(1)
        
        # Right wheel (negative Y, mirrored angles)
        insert_coordinate_system(f"{prefix}R_wheel", x_base, -y_base, z, -camber, 0.0, -toe)
        if progress_callback:
            progress_callback(1)
        
    except KeyError as e:
        print(f"Error: Missing wheel parameter {e}")
    except Exception as e:
        print(f"Error inserting wheels: {e}")


def draw_front_suspension(front_suspension_path: str, progress_callback=None):
    """Draw front suspension from JSON file. Returns total coordinate systems inserted."""
    front_data = load_json(front_suspension_path)
    total = count_hardpoints(front_data) + count_wheels(front_data.get("Wheels", {}))
    
    print("=== Inserting Front Hardpoints ===")
    InsertHardpoint(front_data, "_FRONT", x_offset=0.0, progress_callback=progress_callback)
    
    print("=== Inserting Front Wheels ===")
    InsertWheel(front_data.get("Wheels", {}), reference_distance=0.0, is_rear=False, progress_callback=progress_callback)
    
    return total


def draw_rear_suspension(rear_suspension_path: str, vehicle_setup_path: str, progress_callback=None):
    """Draw rear suspension from JSON file with reference distance offset. Returns total coordinate systems inserted."""
    rear_data = load_json(rear_suspension_path)
    vehicle_data = load_json(vehicle_setup_path)
    
    reference_distance = float(vehicle_data.get("Reference distance", 0.0))
    total = count_hardpoints(rear_data) + count_wheels(rear_data.get("Wheels", {}))
    
    print("=== Inserting Rear Hardpoints ===")
    InsertHardpoint(rear_data, "_REAR", x_offset=reference_distance, progress_callback=progress_callback)
    
    print("=== Inserting Rear Wheels ===")
    InsertWheel(rear_data.get("Wheels", {}), reference_distance=reference_distance, is_rear=True, progress_callback=progress_callback)
    
    return total


def draw_full_suspension(front_path: str, rear_path: str, vehicle_setup_path: str, progress_callback=None):
    """Draw complete suspension from all JSON files. Returns total coordinate systems inserted."""
    front_total = draw_front_suspension(front_path, progress_callback=progress_callback)
    rear_total = draw_rear_suspension(rear_path, vehicle_setup_path, progress_callback=progress_callback)
    return front_total + rear_total


if __name__ == "__main__":
    results_dir = os.path.join(SCRIPT_DIR, "results", "Final EV2024")
    
    front_path = os.path.join(results_dir, "Front_Suspension.json")
    rear_path = os.path.join(results_dir, "Rear_Suspension.json")
    vehicle_setup_path = os.path.join(results_dir, "Vehicle_Setup.json")
    
    draw_full_suspension(front_path, rear_path, vehicle_setup_path)
