# Alternative Approach: SAP Print to File Not Working

## Root Cause Identified

The debug output shows **NO "Save Print Output As" dialog appears** after clicking OK on the Print dialog. This means:

1. ✅ Print dialog opens
2. ✅ OK button is clicked successfully  
3. ❌ Windows file save dialog NEVER appears
4. ❌ Print dialog just closes without action

## Why This Happens in VMs

SAP Print dialog in VMs often requires **explicit printer configuration** or the "Print to file" option to be selected. When you click OK without proper setup, it just closes.

## Solution 1: Configure Printer BEFORE Opening Print Dialog (Recommended)

Instead of dealing with the Windows Print dialog, we can configure SAP to save directly to file using SAP GUI Scripting.

### Add this function to sap_automation.py:

```python
def configure_sap_print_to_file(session, pdf_path):
    """
    Configure SAP print settings to save directly to file.
    Call this BEFORE opening the print dialog.
    """
    try:
        # Access print parameters
        print("  Configuring SAP print settings...")
        
        # Set output device to PDF or LOCAL
        try:
            session.findById("wnd[1]/usr/ctxtPRINT_OPTIONS-PDEST").text = "LP01"  # or "PDF", "LOCAL"
        except:
            pass
        
        # Enable "Print to file" if checkbox exists
        try:
            session.findById("wnd[1]/usr/chkPRINT_OPTIONS-PRI_PARAMS-PRIMM").selected = True
        except:
            pass
        
        # Set output file path
        try:
            session.findById("wnd[1]/usr/ctxtPRINT_OPTIONS-PRI_PARAMS-PRTXT").text = pdf_path
        except:
            pass
            
        print("  ✔ SAP print settings configured")
        return True
        
    except Exception as e:
        print(f"  ⚠ Could not configure SAP print settings: {e}")
        return False
```

### Modify the print workflow (around line 456):

```python
# BEFORE pressing btn[86] (print button), configure SAP print settings
pdf_path = os.path.abspath(os.path.join(pdf_output_dir or config.SAP_PDF_TARGET_DIR, 
                                         f"checklist_{time.strftime('%Y%m%d_%H%M%S')}.pdf"))

# Press print button to open print dialog
session.findById("wnd[1]/tbar[0]/btn[86]").press()
time.sleep(2)

# Configure SAP print settings directly (bypasses Windows dialog)
if configure_sap_print_to_file(session, pdf_path):
    # Press OK to execute print-to-file
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()  # OK button in SAP dialog
        print(f"  ✔ Print initiated to: {pdf_path}")
        time.sleep(3)
        return pdf_path, _checklist_number
    except Exception as e:
        print(f"  ⚠ Could not press SAP OK button: {e}")
```

## Solution 2: Manual Printer Setup in VM

If Solution 1 doesn't work, you need to configure the VM's default printer:

### Steps:
1. In the VM, open Windows Settings → Devices → Printers & scanners
2. Add "Microsoft Print to PDF" as default printer
3. Or install a PDF printer driver (e.g., PDFCreator, CutePDF)
4. Make sure it's set as the default printer

### Then modify the code to use the default printer:

The Print dialog should then automatically show the "Save as PDF" dialog.

## Solution 3: Use SAP Spool System (Most Reliable)

Instead of printing to PDF immediately, let SAP create a spool file, then export it:

```python
def export_spool_to_pdf(session, spool_number, pdf_path):
    """
    Export SAP spool file to PDF.
    More reliable than direct print-to-file.
    """
    try:
        # Go to spool management
        session.findById("wnd[0]/tbar[0]/okcd").text = "SP02"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        
        # Find the spool number
        session.findById("wnd[0]/usr/ctxtRNAME-LOW").text = ""  # All users
        session.findById("wnd[0]/usr/ctxtRQ2TT-LOW").text = spool_number
        session.findById("wnd[0]").sendVKey(8)  # Execute
        time.sleep(2)
        
        # Export to file
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = 0
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/mbar/menu[2]/menu[1]").select()  # Export
        time.sleep(1)
        
        # Specify PDF path
        # ... handle file save dialog ...
        
        return True
    except Exception as e:
        print(f"  ⚠ Spool export failed: {e}")
        return False
```

## Solution 4: Screenshot Fallback

If all else fails, take a screenshot of the SAP window:

```python
def screenshot_checklist(pdf_path):
    """Fallback: screenshot the SAP window as PDF/image."""
    try:
        import pyautogui
        from PIL import Image
        
        # Take screenshot
        screenshot = pyautogui.screenshot()
        
        # Save as PNG (then convert to PDF if needed)
        png_path = pdf_path.replace('.pdf', '.png')
        screenshot.save(png_path)
        
        # Convert to PDF using PIL
        image = Image.open(png_path)
        image.convert('RGB').save(pdf_path)
        
        print(f"  ✔ Screenshot saved as PDF: {pdf_path}")
        return pdf_path
    except Exception as e:
        print(f"  ⚠ Screenshot fallback failed: {e}")
        return None
```

## Immediate Testing Steps

### Test 1: Check SAP Print Dialog Structure

Run this in your VM while the Print dialog is open:

```python
from pywinauto import Application
import time

# Wait for Print dialog to be open
time.sleep(5)

app = Application(backend="uia").connect(title="Print")
print_win = app.window(title="Print")

# List all controls in the Print dialog
print_win.print_control_identifiers()
```

This will show ALL controls in the dialog (checkboxes, dropdowns, etc.).

### Test 2: Check Default Printer

In VM command prompt:
```cmd
wmic printer get name,default
```

If no PDF printer is default, that's your issue.

### Test 3: Manual Click Test

1. Run automation until Print dialog appears
2. **Manually** click OK
3. Does the "Save Print Output As" dialog appear?
   - **YES**: Timing issue - increase waits
   - **NO**: Printer not configured - use Solution 1 or 2

## Recommended Action

Based on your output, I recommend **Solution 1** (configure SAP directly) because:
- No dependency on Windows printer setup
- Works in any VM environment
- More reliable and faster
- SAP GUI Scripting has full control

Would you like me to implement Solution 1 with the SAP GUI Scripting approach?
