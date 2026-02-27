import sys
import os


def _read_com_member(obj, *member_names):
    """Read a COM member that may be exposed as a method or property."""
    for name in member_names:
        attr = getattr(obj, name, None)
        if attr is None:
            continue

        if callable(attr):
            try:
                return attr()
            except TypeError:
                # Some wrappers can expose values unexpectedly; keep trying fallbacks.
                continue

        return attr

    return None


def _get_doc_title(active_doc):
    """Return active document title across COM wrappers/property styles."""
    title = _read_com_member(active_doc, "GetTitle", "Title")
    if title is not None:
        return title

    path_name = _read_com_member(active_doc, "GetPathName", "PathName")
    return path_name if path_name is not None else "<Unknown Document>"


def _get_doc_type(active_doc):
    """Return SolidWorks document type across COM wrappers/property styles."""
    return _read_com_member(active_doc, "GetType", "Type")


def get_active_document_name():
    """Get the name of the active document in SolidWorks."""
    import win32com.client
    sw_app = win32com.client.Dispatch("SldWorks.Application")
    active_doc = sw_app.ActiveDoc
    if active_doc is None:
        raise RuntimeError("No active SolidWorks document found.")
    return _get_doc_title(active_doc)


def get_active_assembly_and_configuration():
    """Get active assembly name and active configuration name."""
    import win32com.client

    sw_app = win32com.client.Dispatch("SldWorks.Application")
    active_doc = sw_app.ActiveDoc
    if active_doc is None:
        raise RuntimeError("No active SolidWorks document found.")

    # swDocumentTypes_e.swDocASSEMBLY == 2
    doc_type = _get_doc_type(active_doc)
    if doc_type != 2:
        doc_name = _get_doc_title(active_doc)
        raise RuntimeError(f"Active document is not an assembly: {doc_name}")

    assembly_name = _get_doc_title(active_doc)
    config_mgr = active_doc.ConfigurationManager
    if config_mgr is None:
        raise RuntimeError("Could not access SolidWorks ConfigurationManager.")

    active_cfg = config_mgr.ActiveConfiguration
    if active_cfg is None:
        raise RuntimeError("Could not determine active SolidWorks configuration.")

    config_name = _read_com_member(active_cfg, "Name")
    if config_name is None:
        raise RuntimeError("Could not read active SolidWorks configuration name.")

    return assembly_name, config_name


def test_solidworks_connection():
    """Test SolidWorks connection and print active assembly/configuration."""
    try:
        assembly_name, config_name = get_active_assembly_and_configuration()
        print(f"✓ SolidWorks connection successful")
        print(f"✓ Active assembly: {assembly_name}")
        print(f"✓ Active configuration: {config_name}")
        return True
    except ImportError:
        print("✗ win32com not installed. Run: pip install pywin32")
        return False
    except Exception as e:
        print(f"✗ Connection failed: {e}")
        return False


if __name__ == "__main__":
    test_solidworks_connection()