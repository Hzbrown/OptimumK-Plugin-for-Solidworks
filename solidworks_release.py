import os
import subprocess


def _suspension_tools_candidates():
    """Get candidate paths to SuspensionTools executable (preferred first)."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return [
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "net48", "SuspensionTools.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug", "net48", "SuspensionTools.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "net48", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug", "net48", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "SuspensionTools.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug", "SuspensionTools.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Release", "sw_drawer.exe"),
        os.path.join(script_dir, "sw_drawer", "bin", "Debug", "sw_drawer.exe"),
    ]


def release_solidworks_command_state():
    """
    Best-effort release of SolidWorks command/edit state after an aborted subprocess.

    Returns:
        (success: bool, message: str)
    """
    exe_path = None
    for candidate in _suspension_tools_candidates():
        if os.path.exists(candidate):
            exe_path = candidate
            break

    if not exe_path:
        return False, "SuspensionTools executable not found for release"

    try:
        result = subprocess.run(
            [exe_path, "release"],
            capture_output=True,
            text=True,
            timeout=10,
        )

        output = (result.stdout or "").strip()
        error = (result.stderr or "").strip()
        message = output or error or f"release exited with code {result.returncode}"
        return result.returncode == 0, message
    except Exception as ex:
        return False, str(ex)
