# AIO+ AHK Changelog

---

Popcorn OCR + macro tab improvements
- OCR detects tame vs storage inventories, uses weight for tames
- Tame popcorn stops at saddle weight via weight OCR
- Crafting inventory OCR strips weight data from slot count
- Popcorn count=0 drops everything w OCR instead of keybinds
- Guided macros save as grid params instead of per-slot events in config
- Guided replay drop timing fixed (was too fast)
- Repeat popcorn F key arms/disarms properly w F1
- F1 show UI fully disarms all macro hotkeys
- Popcorn F no longer triggers before macro is armed

---

Macro hotkey mouse detection + repeat toggle-off
- Detect Key popups capture mouse buttons now.
- Repeat/recorded macros toggle off by pressing the hotkey again.

Added Auto Level saddle-only mode
- Auto Saddle works with zero stat points.

Removed all SoundBeep calls

---

BG Left Click for Key Repeat macros
- Key Repeat with `lbutton` as only key clicks in background.
- `[`/`]` adjust interval. Works with spam and timed.

Combo Popcorn+Magic-F: empty storage detection

Combo wizard: skip drop key prompt if already set

Combo wizard: step windows close properly

BG Mammoth Drums label updated

Upload char: Items Not Allowed detection order fixed

Upload char: failed join retries instead of skipping

Upload char: removed server gating from F6 cycle

Upload char: timer tooltip moved

Upload char: OCR scan area defaults updated

---

Auto Imprint: `[E] Feed` detection fixed

Upload char: Items Not Allowed popup handled

Upload char: 3-stage timer + inv management

Upload char: persistent timer tooltip + F1 clears all
