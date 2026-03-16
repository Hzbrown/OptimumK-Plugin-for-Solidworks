import subprocess
from utils import find_suspension_tools_exe


def release_solidworks_command_state():
    """
    Best-effort release of SolidWorks command/edit state after an aborted subprocess.

    Returns:
        (success: bool, message: str)
    """
    try:
        exe_path = find_suspension_tools_exe()
    except FileNotFoundError:
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
