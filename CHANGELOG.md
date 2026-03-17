# AIO+ AHK Changelog

PORT items from the Python version applied to `AIO_+.ahk`.

---

## Applied (2026-03-17)

### Detect key: mouse button capture
- **Status**: Already implemented — `MacroDetectRepeatKey()` already has RButton/LButton/MButton support

### Auto Imprint: `[E] Feed` fuzzy bracket matching
- **Problem**: Strict `InStr(ocrText, "[E]")` never matched because OCR misreads brackets as `(E)`, `IE]`, `[E`, etc.
- **Fix**: Added `ImHasFeedPrompt()` helper with fuzzy matching. Replaced strict checks at both detection points.

### "Items Not Allowed" popup detection in upload char join loop
- **Problem**: Join loop had no OCR check for popup text. Looped 30 times, then set `obCharLastDest` before join was confirmed — corrupting server cycle.
- **Fix**: Added OCR check for "Items Not Allowed" / "not ready for upload" / "can not be transferred". On detection: Esc x2 to exit, re-arm at stage 1 (inv management). Moved `obCharLastDest` to after join confirmed.

### Upload char: 3-stage timer + inv management flow
- **Problem**: `OBCheckUploadTimer()` blocked with a countdown loop. No way to manage items. Pressing F after timer just re-detected the same timer.
- **Fix**: 3-stage flow using `obCharTimerStage`:
  - Stage 0 (first F): OCR checks for timer. If found → close transmitter, show timer, re-arm.
  - Stage 1 (second F): Inv management — `OBWaitInvClose()` polls for inv close. User manages items, closes inv. Then "Press F at transmitter".
  - Stage 2 (third F): Re-checks timer. If still active → back to stage 1. If clear → proceeds to Travel.
  - Items Not Allowed popup → set stage to 1. F1 → full clear, reset stage to 0.
