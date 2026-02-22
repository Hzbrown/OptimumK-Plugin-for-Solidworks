import sys
import os


def get_active_document_name():
    """Get the name of the active document in SolidWorks."""
    import win32com.client
    sw_app = win32com.client.Dispatch("SldWorks.Application")
    active_doc = sw_app.ActiveDoc
    if active_doc is None:
        raise RuntimeError("No active SolidWorks document found.")
    return active_doc.GetTitle  # property, not a method


def test_solidworks_connection():
    """Test SolidWorks connection and print active document name."""
    try:
        doc_name = get_active_document_name()
        print(f"✓ SolidWorks connection successful")
        print(f"✓ Active document: {doc_name}")
        return True
    except ImportError:
        print("✗ win32com not installed. Run: pip install pywin32")
        return False
    except Exception as e:
        print(f"✗ Connection failed: {e}")
        return False


if __name__ == "__main__":
    test_solidworks_connection()