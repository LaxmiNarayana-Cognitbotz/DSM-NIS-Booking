# VM SAP Automation Troubleshooting Guide

## Changes Made for VM Compatibility

### 1. Print Dialog Handling Improvements
- **Multiple Backend Support**: Now tries `win32` backend first (more reliable in VMs), then `uia`
- **Increased Timeouts**: Extended connection timeout from 10s to 15s
- **Window Focus Management**: Added `.restore()` to ensure windows aren't minimized
- **Better Error Handling**: Each backend attempt has its own try-catch with detailed logging
- **Keyboard Fallback**: If button click fails, automatically tries `{ENTER}` key

### 2. Save Dialog Handling Improvements
- **Longer Wait Time**: Increased from 3s to 5s for dialog to appear
- **Backend Priority**: Tries `win32` first (more stable in VMs), then `uia`
- **Enhanced Window Focus**: Ensures window is brought to front and restored
- **Backend-Specific Field Finding**: win32 uses `class_name`, uia uses `auto_id`
- **Increased Timing Delays**: All `time.sleep()` values increased for VM slowness

### 3. File Name Field Handling
- **Backend-Aware Search**: Different strategies for win32 vs uia
- **Longer Wait Times**: Field wait timeout increased from 3s to 5s
- **Focus Management**: Extra delays after focus operations (0.5s instead of 0.2s)
- **Clipboard Timing**: Increased delays for clipboard operations

### 4. Save Button Click
- **Backend-Specific Clicking**: win32 uses `class_name`, uia uses `control_type`
- **Button Wait**: Added `.wait("visible enabled")` before clicking
- **Enhanced Fallback**: Better keyboard fallback with focus management

## Common VM Issues and Solutions

### Issue 1: Dialog Not Found
**Symptoms**: "Could not interact with Print dialog" or "Cannot find Save Print Output As dialog"

**Solutions**:
1. **Check VM Window Focus**: Ensure the VM window has focus before running
2. **Disable VM Tools**: Some VM tools intercept window events
3. **Run Script from Within VM**: Don't run from host machine via RDP/SSH
4. **Check Window Title**: SAP dialog titles might differ in different versions

### Issue 2: Button Click Does Nothing
**Symptoms**: Dialog found but clicking OK doesn't work

**Solutions**:
1. The script now tries both backends automatically
2. Falls back to keyboard input if click fails
3. Check if UAC is blocking interactions
4. Run PowerShell/Command Prompt as Administrator

### Issue 3: Timing Issues
**Symptoms**: Works manually but automation fails

**Solutions**:
- Already addressed with increased delays throughout
- If still too fast, modify these values in code:
  ```python
  time.sleep(5)   # After print button (line ~456)
  time.sleep(5)   # Before save dialog (line ~532)
  time.sleep(0.8) # After save dialog focus (line ~582)
  ```

### Issue 4: Clipboard Not Working
**Symptoms**: File path not pasted

**Solutions**:
1. Install `pyperclip` properly: `pip install pyperclip`
2. Check if VM allows clipboard sharing
3. The script has fallback: `.type_keys()` character-by-character

## Testing Steps

1. **Test Print Dialog**:
   ```python
   # Add after line ~465 to verify dialog is found:
   print(f"  DEBUG: Print window found: {_print_win}")
   print(f"  DEBUG: Print window visible: {_print_win.is_visible()}")
   ```

2. **Test Save Dialog**:
   ```python
   # Add after line ~557 to verify:
   print(f"  DEBUG: Save window found: {_save_win}")
   print(f"  DEBUG: Backend used: {_found_backend}")
   ```

3. **Test Button State**:
   ```python
   # Add before clicking OK button:
   ok_btn = _print_win.child_window(title="OK", control_type="Button")
   print(f"  DEBUG: OK button exists: {ok_btn.exists()}")
   print(f"  DEBUG: OK button enabled: {ok_btn.is_enabled()}")
   print(f"  DEBUG: OK button visible: {ok_btn.is_visible()}")
   ```

## VM-Specific Settings to Check

### VMware
- **Enhanced Keyboard**: Disable in VM settings
- **3D Acceleration**: Try disabling if issues persist
- **Display Settings**: Set to "Autodetect" or "Single monitor"

### VirtualBox
- **Guest Additions**: Ensure installed and up-to-date
- **Mouse Integration**: Can cause focus issues - try disabling
- **Video Memory**: Increase to 128MB+

### Hyper-V
- **Enhanced Session Mode**: Try both enabled/disabled
- **Integration Services**: Ensure updated

## Advanced Debugging

### Enable Verbose Logging
Add this at the start of the function:
```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

### Inspect All Windows
Add this to see what windows are available:
```python
from pywinauto import Desktop
desktop = Desktop(backend="uia")
for win in desktop.windows():
    try:
        print(f"Window: {win.window_text()}")
    except:
        pass
```

### Screenshot on Failure
```python
try:
    # ... dialog interaction code ...
except Exception as e:
    import pyautogui
    pyautogui.screenshot(f"error_{time.time()}.png")
    raise
```

## Performance Tuning

If automation is too slow, you can reduce these (but may break in slow VMs):
- Line ~456: `time.sleep(5)` → `time.sleep(3)`
- Line ~532: `time.sleep(5)` → `time.sleep(3)`
- Line ~582: `time.sleep(0.8)` → `time.sleep(0.5)`

If automation is too fast and missing dialogs, increase these values.

## Contact & Support

If issues persist after trying all solutions:
1. Check SAP version and VM configuration
2. Verify pywinauto version: `pip show pywinauto` (should be 0.6.8+)
3. Try running the old code with the complex retry logic from git history
4. Consider using SAP GUI Scripting API directly (no pywinauto) for more reliability
