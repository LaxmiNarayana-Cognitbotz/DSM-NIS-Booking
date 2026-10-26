"""
Inspect SAP Print Dialog Structure
Run this while the Print dialog is open to see all available controls.
"""

from pywinauto import Application
import time

print("\n" + "="*60)
print("  SAP PRINT DIALOG INSPECTOR")
print("="*60)
print("\nInstructions:")
print("1. Run your SAP automation until the Print dialog appears")
print("2. Keep the Print dialog OPEN")
print("3. Run this script within 10 seconds")
print("\nWaiting 10 seconds for you to prepare...")

for i in range(10, 0, -1):
    print(f"  {i}...", end="\r")
    time.sleep(1)

print("\n\nInspecting Print dialog...\n")

try:
    # Try UIA backend first
    print("─" * 60)
    print("  ATTEMPT 1: UIA Backend")
    print("─" * 60)
    try:
        app = Application(backend="uia").connect(title="Print", timeout=5)
        print_win = app.window(title="Print")
        
        print(f"\n✓ Found Print dialog (UIA)")
        print(f"  Title: {print_win.window_text()}")
        print(f"  Class: {print_win.class_name()}")
        print(f"  Visible: {print_win.is_visible()}")
        print(f"  Enabled: {print_win.is_enabled()}")
        
        print("\n" + "="*60)
        print("  ALL CONTROLS IN PRINT DIALOG (UIA)")
        print("="*60 + "\n")
        
        print_win.print_control_identifiers(depth=5, filename=None)
        
    except Exception as e:
        print(f"✗ UIA backend failed: {e}")
    
    # Try Win32 backend
    print("\n\n" + "─" * 60)
    print("  ATTEMPT 2: Win32 Backend")
    print("─" * 60)
    try:
        app = Application(backend="win32").connect(title="Print", timeout=5)
        print_win = app.window(title="Print")
        
        print(f"\n✓ Found Print dialog (Win32)")
        print(f"  Title: {print_win.window_text()}")
        print(f"  Class: {print_win.class_name()}")
        
        print("\n" + "="*60)
        print("  ALL CONTROLS IN PRINT DIALOG (Win32)")
        print("="*60 + "\n")
        
        print_win.print_control_identifiers(depth=5, filename=None)
        
    except Exception as e:
        print(f"✗ Win32 backend failed: {e}")
    
    # Additional: Try to find specific controls
    print("\n\n" + "="*60)
    print("  SEARCHING FOR KEY CONTROLS")
    print("="*60)
    
    try:
        app = Application(backend="uia").connect(title="Print", timeout=5)
        print_win = app.window(title="Print")
        
        # Look for "Print to file" checkbox
        print("\n1. Searching for 'Print to file' checkbox...")
        try:
            checkboxes = print_win.children(control_type="CheckBox")
            if checkboxes:
                print(f"   Found {len(checkboxes)} checkbox(es):")
                for i, cb in enumerate(checkboxes, 1):
                    try:
                        print(f"   {i}. '{cb.window_text()}' [Checked: {cb.get_toggle_state()}]")
                    except:
                        print(f"   {i}. <unable to read>")
            else:
                print("   No checkboxes found")
        except Exception as e:
            print(f"   Error: {e}")
        
        # Look for combo boxes (printer selection)
        print("\n2. Searching for ComboBox (printer selection)...")
        try:
            combos = print_win.children(control_type="ComboBox")
            if combos:
                print(f"   Found {len(combos)} ComboBox(es):")
                for i, combo in enumerate(combos, 1):
                    try:
                        print(f"   {i}. Current value: '{combo.texts()}'")
                    except:
                        print(f"   {i}. <unable to read>")
            else:
                print("   No combo boxes found")
        except Exception as e:
            print(f"   Error: {e}")
        
        # Look for edit boxes (file path)
        print("\n3. Searching for Edit fields...")
        try:
            edits = print_win.children(control_type="Edit")
            if edits:
                print(f"   Found {len(edits)} Edit field(s):")
                for i, edit in enumerate(edits, 1):
                    try:
                        print(f"   {i}. Value: '{edit.get_value()}'")
                    except:
                        print(f"   {i}. <unable to read>")
            else:
                print("   No edit fields found")
        except Exception as e:
            print(f"   Error: {e}")
        
        # Look for buttons
        print("\n4. Searching for Buttons...")
        try:
            buttons = print_win.children(control_type="Button")
            if buttons:
                print(f"   Found {len(buttons)} button(s):")
                for i, btn in enumerate(buttons, 1):
                    try:
                        print(f"   {i}. '{btn.window_text()}'")
                    except:
                        print(f"   {i}. <unable to read>")
            else:
                print("   No buttons found")
        except Exception as e:
            print(f"   Error: {e}")
            
    except Exception as e:
        print(f"Search failed: {e}")
    
except Exception as e:
    print(f"\nERROR: {e}")
    print("\nMake sure:")
    print("  1. SAP Print dialog is open")
    print("  2. The dialog has focus (click on it)")
    print("  3. Run this script within 10 seconds of opening dialog")

print("\n\n" + "="*60)
print("  INSPECTION COMPLETE")
print("="*60)
print("\nAnalysis:")
print("• Look for 'Print to file' checkbox - if found, we can enable it")
print("• Look for ComboBox with PDF/file printer options")
print("• Look for file path Edit field")
print("\nIf NO 'Print to file' option exists, you need to:")
print("  → Install a PDF printer in the VM (Microsoft Print to PDF)")
print("  → Or use SAP GUI Scripting to configure print output")

input("\n\nPress Enter to exit...")
