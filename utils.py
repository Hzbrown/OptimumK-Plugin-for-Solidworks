import os
import sys


def get_data_dir():
    """Return the writable base directory for user data (profiles, poses, etc.).

    Packaged builds: %LOCALAPPDATA%/OptimumK  (avoids writing into Program Files)
    Dev builds:      the project root (same as script location)
    """
    if getattr(sys, 'frozen', False):
        base = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'OptimumK')
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(base, exist_ok=True)
    return base


def get_temp_dir():
    """Return the writable temp directory for parsed JSON data."""
    temp = os.path.join(get_data_dir(), 'temp')
    os.makedirs(temp, exist_ok=True)
    return temp


def find_suspension_tools_exe():
    """Locate SuspensionTools.exe.

    Search order:
    1. Next to sys.executable (packaged PyInstaller build)
    2. Development build outputs, Release before Debug, newest mtime wins
    """
    # Packaged build: SuspensionTools.exe ships alongside the Python exe.
    if getattr(sys, 'frozen', False):
        packaged = os.path.join(os.path.dirname(sys.executable), "SuspensionTools.exe")
        if os.path.exists(packaged):
            return packaged

    script_dir = os.path.dirname(os.path.abspath(__file__))
    paths_to_check = [
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "net48", "SuspensionTools.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug",   "net48", "SuspensionTools.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "net48", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug",   "net48", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "SuspensionTools.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug",   "SuspensionTools.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug",   "sw_drawer.exe"),
    ]

    existing = [p for p in paths_to_check if os.path.exists(p)]
    if existing:
        preferred = [p for p in existing if os.path.basename(p).lower() == "suspensiontools.exe"]
        candidates = preferred if preferred else existing
        return max(candidates, key=os.path.getmtime)

    raise FileNotFoundError(
        "SuspensionTools.exe not found. Run 'dotnet build -c Release' in the sw_drawer folder.\n"
        f"Searched: {paths_to_check[0]}"
    )
