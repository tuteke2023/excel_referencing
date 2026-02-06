# TB-GL Linker - Session Summary (2026-02-06)

## Overview
Deployed and debugged the TB-GL Linker web application with Claude CLI integration for AI-powered structure detection.

## Deployment

### Service Configuration
- **URL:** https://excel-referencing.ttaccountancy.au
- **Port:** 8530
- **Cloudflare Tunnel:** `excel-referencing` (ID: 98f22c66-4078-4543-874f-ef316af389a3)
- **Systemd Services:**
  - `excel-referencing.service` - Streamlit app
  - `excel-referencing-tunnel.service` - Cloudflare tunnel

### Files Location
- **Repository:** `/home/user/projects/excel_referencing`
- **Tunnel Config:** `/home/user/.cloudflared/config-excel-referencing.yml`

## Bug Fixes

### 1. Claude CLI Detection (web_app.py)
**Issue:** Web app blocked when `CLAUDE_API_URL` not configured.
**Fix:** Added CLI-mode fallback in `check_claude_api_connection()` to detect local Claude CLI.

### 2. "Total" Rows Detected as Account Headers
**Issue:** "Total Bank Fees" was detected as a separate account header, causing the search for "Net Movement" to stop too early.
**Fix:** Added exclusion in `_is_account_header()`:
```python
if value_lower.startswith('total ') or value_lower == 'net movement':
    return False
```

### 3. "Account Type" Column Misdetection
**Issue:** TB header detection matched "Account Type" column instead of "Account" column.
**Fix:** Skip columns with "type" in header and improved column detection logic.

### 4. Claude Returns Generic Patterns
**Issue:** Claude returned `['Net Movement', 'Total']` as summary patterns, causing "Total Bank Fees" to match before "Net movement".
**Fix:** Filter out generic patterns in `_find_net_movement()`:
```python
summary_texts = [t for t in summary_texts if t.lower() not in ['total', 'balance', 'net']]
```

### 5. Static Hyperlink References
**Issue:** HYPERLINK formulas use static cell addresses that don't adjust when rows are inserted/deleted in GL.
**Fix:** Changed from HYPERLINK to direct cell references:
```python
# Before: =HYPERLINK("#'General Ledger Detail'!E1151", 'General Ledger Detail'!E1151)
# After:  ='General Ledger Detail'!E1151
```

## Key Changes Summary

| File | Change |
|------|--------|
| `web_app.py` | CLI fallback, exclude Total rows, filter Claude patterns, direct cell refs |
| `tb_gl_linker_netmovement.py` | Same fixes as web_app.py for CLI consistency |

## Testing

Tested with Boardwalk Dental Pty Ltd files:
- Trial Balance: 68 accounts
- General Ledger: 4894 rows, 67 account sections
- Successfully linked 57 accounts to Net Movement figures

## Version
- Version: 1.2.1
- Claude CLI: 2.1.32

## Authors
- Fixes implemented by: Clawd (AI Assistant)
- Testing by: Teke Tu, Sydelle Tay
