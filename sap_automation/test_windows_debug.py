"""
Quick debug script to list all visible windows in the VM.
Run this while the SAP Save dialog is open to see its exact title.
"""

from pywinauto import Desktop
import time

def list_all_windows():
    """List all visible windows with details."""
    try:
        desktop = Desktop(backend="uia")
        print("\n═══════════════════════════════════════════")
        print("  ALL VISIBLE WINDOWS (UIA Backend)")
        print("═══════════════════════════════════════════")
        
        windows = desktop.windows()
        print(f"Total windows found: {len(windows)}\n")
        
        for i, win in enumerate(windows, 1):
            try:
                title = win.window_text()
                if title:
                    try:
                        visible = win.is_visible()
                        enabled = win.is_enabled()
                        class_name = win.class_name()
                        print(f"{i}. Title: '{title}'")
                        print(f"   Class: {class_name}")
                        print(f"   Visible: {visible}, Enabled: {enabled}")
                        print()
                    except:
                        print(f"{i}. Title: '{title}' (properties unavailable)")
                        print()
            except Exception as e:
                print(f"{i}. <Unable to read window: {e}>")
                print()
        
        print("═══════════════════════════════════════════\n")
    except Exception as e:
        print(f"ERROR: Could not list windows: {e}")


def list_all_windows_win32():
    """List all visible windows with win32 backend."""
    try:
        desktop = Desktop(backend="win32")
        print("\n═══════════════════════════════════════════")
        print("  ALL VISIBLE WINDOWS (Win32 Backend)")
        print("═══════════════════════════════════════════")
        
        windows = desktop.windows()
        print(f"Total windows found: {len(windows)}\n")
        
        for i, win in enumerate(windows, 1):
            try:
                title = win.window_text()
                if title:
                    try:
                        visible = win.is_visible()
                        class_name = win.class_name()
                        print(f"{i}. Title: '{title}'")
                        print(f"   Class: {class_name}")
                        print(f"   Visible: {visible}")
                        print()
                    except:
                        print(f"{i}. Title: '{title}' (properties unavailable)")
                        print()
            except Exception as e:
                print(f"{i}. <Unable to read window: {e}>")
                print()
        
        print("═══════════════════════════════════════════\n")
    except Exception as e:
        print(f"ERROR: Could not list windows with win32: {e}")


def search_for_dialog(search_text="Save"):
    """Search for windows containing specific text."""
    try:
        desktop = Desktop(backend="uia")
        print(f"\n═══════════════════════════════════════════")
        print(f"  SEARCHING FOR: '{search_text}'")
        print("═══════════════════════════════════════════")
        
        found_count = 0
        for win in desktop.windows():
            try:
                title = win.window_text()
                if title and search_text.lower() in title.lower():
                    found_count += 1
                    print(f"✓ FOUND: '{title}'")
                    try:
                        print(f"  Class: {win.class_name()}")
                        print(f"  Visible: {win.is_visible()}")
                        print(f"  Enabled: {win.is_enabled()}")
                    except:
                        pass
                    print()
            except:
                pass
        
        if found_count == 0:
            print(f"No windows found containing '{search_text}'")
        else:
            print(f"\nTotal matches: {found_count}")
        
        print("═══════════════════════════════════════════\n")
    except Exception as e:
        print(f"ERROR: {e}")


if __name__ == "__main__":
    print("\n" + "="*50)
    print("  PYWINAUTO WINDOW DEBUG TOOL")
    print("="*50)
    print("\nInstructions:")
    print("1. Open the SAP Save dialog manually")
    print("2. Run this script")
    print("3. Find the exact window title in the output")
    print("\nStarting in 3 seconds...\n")
    
    time.sleep(3)
    
    # List all windows with both backends
    list_all_windows()
    list_all_windows_win32()
    
    # Search for common SAP/Save dialog keywords
    search_for_dialog("Save")
    search_for_dialog("Print")
    search_for_dialog("Output")
    search_for_dialog("SAP")
    
    print("\n" + "="*50)
    print("  DEBUG COMPLETE")
    print("="*50)
    print("\nLook for the 'Save Print Output As' dialog in the output above.")
    print("If you see it, check its exact title and update the code accordingly.")
    input("\nPress Enter to exit...")
