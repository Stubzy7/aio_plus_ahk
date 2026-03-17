# AIO+ AHK Changelog

PORT items from the Python version applied to `AIO_+.ahk`.

---

## 2026-03-18

### Combo Popcorn+Magic-F: OCR empty storage detection
- Drop loop now checks `PcCheckStorageEmpty()` after each pass, matching standalone popcorn behavior.

### Combo wizard: skip drop key prompt if already set
- Checks INI for saved drop key — skips prompt if present. Same for guided popcorn wizard.

### Combo wizard: fix step dialogs not closing
- Step transitions use `Hide()` + destroy guard to prevent dialogs lingering.

### BG Mammoth Drums label
- GUI label and tooltip updated to "BG Mammoth Drums" to match Python.

### Upload char: Items Not Allowed detection fix
- OCR now checks for "Items Not Allowed" **before** "Joining" — prevents false join confirmation when both texts appear in same scan.
- Tooltip hidden before OCR scan to avoid own tooltip text polluting results.

### Upload char: failed join no longer marks server
- If join isn't confirmed after 30 attempts, re-arms instead of proceeding. Server stays available in cycle.

### Upload char: removed server gating (`obCharLastDest`)
- No server is blocked from the F6 cycle. User always has full control over destination.

### Upload char: timer tooltip positioning
- Timer tooltip (ID 2) moved to y=40 to avoid overlap with status tooltip.

### Upload char: OCR scan area defaults
- Updated default timer OCR region (index 5) to cover full player inventory area.

### BG Mammoth Drums label
- GUI label and tooltip updated to "BG Mammoth Drums" to match Python version.

---

## 2026-03-17

### Auto Imprint: `[E] Feed` fuzzy bracket matching
- Added `ImHasFeedPrompt()` helper with fuzzy matching. Replaced strict checks at both detection points.

### Upload char: Items Not Allowed popup detection
- Added OCR check in join loop for "Items Not Allowed" / "not ready for upload" / "can not be transferred". On detection: Esc x2, re-arm at stage 1.

### Upload char: 3-stage timer + inv management flow
- 3-stage flow using `obCharTimerStage` replacing blocking `OBCheckUploadTimer()`.

### Upload char: persistent timer tooltip + F1 full clear
- Timer shown as persistent ToolTip ID 2. F1 calls `OBStopAll()` for full clear.
