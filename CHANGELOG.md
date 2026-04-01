# AIO+ AHK Changelog

---

Inventory access key is now configurable (default F)
- Anyone who rebinds F key in ARK can set their key in Misc tab > Set Keys
- Access key = open tame/storage inv, Inv key = your own inventory (for sheep script), Drop key = key you popcorn with
- Every mode responds to it: popcorn, craft, magic F, macros, quick hatch, auto level, OB, GMK
- Tooltips and labels update to show the configured key

Sheep script runs fully in background
- Can tab out of ARK and script keeps running
- Clicking back into ARK pauses background clicks so camera works normally
- Start/pause and F1 activate even when not clicked into ARK, set keys to something you wouldn't press by mistake on another monitor
- Sheep keys save to INI and load on startup

---

Fixed Name/Spay E hold missing the radial wheel
Fixed F key being stolen by macro popcorn when hatch or depo modes are active
F1 now keeps hatch/claim/name/depo checkbox selections when showing UI

---

Fixed AutoLvL and all F key features broken by MacroDisarmPopcornF killing the static ~$F:: handler via Hotkey("$f","Off")

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
