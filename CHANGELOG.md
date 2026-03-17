# AIO+ AHK Changelog

PORT items from the Python version that should be applied to `AIO_+.ahk`.

---

## Pending Ports (from Python 2026-03-17)

### Detect key: add mouse button capture
- **Python fix**: `gui/tab_macro.py` — added mouse button listener to detect key dialog
- **AHK action**: Update detect key to also capture mouse button clicks (LButton, RButton, etc.)

### Auto Imprint: `[E] Feed` detection never triggering
- **Python fix**: `modules/auto_imprint.py` — fuzzy bracket matching for OCR misreads
- **AHK action**: Line ~16107 has strict `[E]` match. Loosen to accept `(E)`, `IE]`, `[E`, etc. Same silent failure likely exists.

### "Items Not Allowed" popup blocks upload + corrupts server cycle
- **Python fix**: `modules/ob_upload.py` — OCR detects popup text in join loop, exits transmitter, re-arms
- **AHK action**: Add OCR/text check for "Items Not Allowed" / "not ready for upload" / "can not be transferred" during join loop. On detection: exit transmitter, re-arm for retry. Don't set last dest until join confirmed.
- **Two popup variants**: (1) "not ready for upload — drop them here?" with Accept/Cancel, (2) "can not be transferred — drop such items" with OK.

### Upload char: 3-stage timer + inv management flow
- **Python fix**: `modules/ob_upload.py` — 3-stage flow with `ob_char_timer_stage`
- **AHK action**: Implement same flow:
  - Stage 0 (first F): OCR checks for timer. If found → close transmitter, show timer, re-arm.
  - Stage 1 (second F): Inv management — user opens inv, removes timed items, closes inv. Poll for inv close, then "Press F at transmitter".
  - Stage 2 (third F): Re-check timer. If still active → back to stage 1. If clear → proceed to Travel.
  - Items Not Allowed popup → set stage to 1.
  - F1 → full clear, reset stage to 0.
