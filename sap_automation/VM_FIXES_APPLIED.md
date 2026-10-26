# VM Fixes Applied - Summary

## Problem Identified
Your log shows:
```
✔ Clicked OK on Print dialog (backend=uia)  ← THIS WORKED!
Saving PDF to: C:\Users\agel_rpabot_pm\Desktop\DSM-NIS-\sap_automation\temp-book\AWEK01L-50 MW\checklist_20260711_115230.pdf
Connecting to 'Save Print Output As' dialog...
Trying backend=win32...
⚠ backend=win32 failed:                      ← STUCK HERE
Trying backend=uia...
```

**Root Cause**: The "Save Print Output As" dialog is either:
1. Taking too long to appear (timing issue)
2. Has a different title than expected
3. Being blocked/hidden in the VM

## Fixes Applied

### 1. Enhanced Save Dialog Detection (Lines ~530-610)
```python
# Added 3 retry attempts with progressive delays
# Added 4 different title patterns to match:
- "Save Print Output As" (exact)
- ".*Save.*Print.*" (regex: any Save...Print)
- "Save Print Output" (shorter)
- ".*Save.*Output.*" (regex: any Save...Output)

# Each retry waits longer: 5s, 8s, 11s
```

### 2. Better Error Logging
- Now shows which title pattern is being tried
- Lists all visible windows if dialog not found (for debugging)
- Shows exact error messages

### 3. Multiple Backend Attempts
- Tries win32 first (VM-friendly)
- Falls back to uia
- Then tries Desktop search
- Each with multiple title variations

## How to Use the Debug Tool

### Step 1: Manual Test
1. Run your SAP automation until it's stuck looking for the Save dialog
2. **Manually click the Print OK button** so the Save dialog appears
3. **Keep the Save dialog open**
4. Run the debug script:
   ```bash
   cd sap_automation
   python test_windows_debug.py
   ```

### Step 2: Find the Exact Title
The script will show ALL windows. Look for output like:
```
✓ FOUND: 'Save Print Output As'
  Class: #32770
  Visible: True
  Enabled: True
```

If you see a different title like:
- "Guardar salida de impresión como" (Spanish)
- "Save as" 
- "Speichern unter" (German)
- Or ANY variation

### Step 3: Update the Code (if needed)
If the title is different, edit `sap_automation.py` line ~548:
```python
_dialog_titles = [
    "Save Print Output As",
    "YOUR_EXACT_TITLE_HERE",  # Add this line
    ".*Save.*Print.*",
    "Save Print Output",
    ".*Save.*Output.*",
]
```

## Alternative: Check If Dialog Appears At All

The issue might be that clicking OK on the Print dialog **doesn't trigger** the Save dialog in the VM. To test this:

1. Run automation until it clicks Print OK
2. Wait 10 seconds
3. Look at the VM screen - is there a Save dialog visible?

If **NO** dialog appears, the issue is different - the Print OK click might need adjustment.

## Quick Fix: Increase Wait Time

If the dialog appears but slowly, increase the initial wait:

Edit line ~531 in `sap_automation.py`:
```python
time.sleep(10)  # Change from 5 to 10 seconds
```

## Common VM-Specific Issues

### Issue 1: Dialog Appears Behind Other Windows
**Symptom**: Dialog exists but script can't find it
**Fix**: Already added `.restore()` and `.set_focus()` in the code

### Issue 2: VM Performance
**Symptom**: Everything works but very slowly
**Fix**: Increase sleep times:
- Line ~531: `time.sleep(10)` (currently 5)
- Line ~547: `time.sleep(5 * _retry)` (currently 3)

### Issue 3: Different SAP Version/Language
**Symptom**: Dialog title is in different language or worded differently
**Fix**: Use the debug tool to find exact title, add to `_dialog_titles` list

## Next Steps

1. **Run the updated code** - It now has 3 retries with multiple title patterns
2. **If it still fails**, run `test_windows_debug.py` while Save dialog is open
3. **Report back** with:
   - Does the Save dialog appear visually in the VM?
   - What's the exact title shown by the debug tool?
   - Any error messages

## Code Changes Summary

### File: `sap_automation.py`

**Line ~10-30**: Added `_list_all_windows_debug()` function

**Lines ~530-610**: Complete rewrite of Save dialog detection:
- 3 retry attempts
- 4 title pattern variations
- Progressive delay increases
- Comprehensive error logging
- Window enumeration on failure

**Lines ~464-528**: Print dialog (already working, no changes needed)

## Testing Commands

```bash
# Test in VM
cd C:\Users\agel_rpabot_pm\Desktop\DSM-NIS-\sap_automation
python test_windows_debug.py

# Or run your main automation
python sap_main.py
```

## Expected New Output

You should now see:
```
Connecting to 'Save Print Output As' dialog...
  Trying backend=win32, title='Save Print Output As'...
  Trying backend=win32, title='.*Save.*Print.*'...
  Trying backend=uia, title='Save Print Output As'...
  ✔ Found Save dialog (backend=uia, title='Save Print Output As')
```

Or if it fails after retries:
```
  ✖ Cannot find 'Save Print Output As' dialog

  ═══ All Visible Windows ═══
  1. 'SAP Easy Access' [visible=True, enabled=True]
  2. 'Microsoft Edge' [visible=True, enabled=True]
  ...
```

This will tell us exactly what's available to connect to!
