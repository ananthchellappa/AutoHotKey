# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

A personal collection of AutoHotkey (AHK) scripts for Windows, centered on a KDE-style
window manager (ALT+drag to move, ALT+right-drag to resize, double-ALT chords) plus
per-application workarounds (Excel, PowerPoint, MSPaint, Notepad++, Tanner S-Edit /
Custom IC Waveform Viewer). `README.md` documents the user-facing hotkeys.

## No build, test, or lint

These are interpreted scripts. There is no toolchain, no dependencies to install, and no
test suite. "Running" a script means, on a Windows machine:

- Copy the chosen `.ahk` file **and the `Lib/` directory** into the same folder (the
  Startup folder for permanent use — `Win+R`, `shell:startup`), then double-click it.
- `CTRL+ALT+F9` (or `CTRL+ALT+RButton`) suspends/resumes the running script.
- `CTRL+ALT+SHIFT+MButton` reloads it — but only while *not* suspended.

The working copy here is edited from WSL/Linux, so scripts cannot be executed or verified
in this environment. Changes must be hand-tested by the user on Windows; do not claim a
hotkey works, only that it is written.

## AHK v1 vs v2 — do not mix

Almost everything is **AutoHotkey v1** legacy syntax (comma-separated commands,
`%var%` dereferencing, `WinGet`, `VarSetCapacity`, labels + `return`). The only v2 file is
`Toggle_TaskBar_Auto_hide.ahk`, which declares `#Requires AutoHotkey v2.0` and uses
`Buffer()`/`NumPut()`/expression syntax. When editing a file, match the dialect already in
it; the two cannot be `#Include`d into each other.

## The three script variants have diverged deliberately

`EasyWindowDrag_KDE.ahk`, `Easy_KDE_Win11`, and `All_Plus_Tanner_SEdit_Workarounds.ahk`
are three copies of the same base script, each maintained for a different machine. They
share ~250 lines of KDE window-drag code but are **not** kept in sync — recent commits
touch exactly one of them at a time. A change to shared behavior must be applied to each
variant explicitly and consciously, not assumed to propagate.

| File | Target | Distinguishing content | `Notify` include |
|---|---|---|---|
| `EasyWindowDrag_KDE.ahk` | the one README tells users to install; most complete | Snipping Tool / Paint launchers, YouTube URL fix, Notepad++ ALT-scroll, MSPaint font size via MSAA, Excel row select | `#Include <Notify>` (resolves from `Lib/`) |
| `Easy_KDE_Win11` | home Win11 box (note: no `.ahk` extension) | `ALT+3` throws the window to the next monitor using native `#+Left/Right` | hardcoded `C:\Users\anant\...` |
| `All_Plus_Tanner_SEdit_Workarounds.ahk` | work box | `#IfWinActive ahk_exe sedit64.exe` and `Custom IC Waveform Viewer` blocks: RMB click-vs-drag disambiguation that enters S-Edit zoom mode, wheel→zoom/pan remaps, `!1` waveform plotting | hardcoded `C:\Users\ananth.chellappa\...` |
| `EasyWindowDrag_KDE_v2.ahk` | **AutoHotkey v2 port** of `EasyWindowDrag_KDE.ahk` (see below) | same hotkeys, v2 syntax | `#Include <Notify_v2>` |

Only `EasyWindowDrag_KDE.ahk` uses the portable `#Include <Notify>` form; the other two
carry absolute per-user paths that are dead on any other machine. Prefer the library form
for new work.

## The v2 port

`EasyWindowDrag_KDE_v2.ahk` is a feature-parity AutoHotkey v2 port of
`EasyWindowDrag_KDE.ahk`. Both are kept; run one or the other, never both (duplicate
hotkeys). The v1 file remains the one README points users at until the port is tested.

It cannot use `Lib/Notify.ahk` or `Lib/AccessibleObject.ahk` — v1 and v2 can't include each
other — so it depends on two v2-only stand-ins:

- `Lib/Notify_v2.ahk` — small stacking toast exposing the same `Notify(Title, Message, Duration)`
  call shape. Deliberately does *not* reproduce gwarble's negative-duration "ExitApp on click"
  behavior, which nothing here used.
- `Lib/Acc_v2.ahk` — minimal MSAA/IAccessible walker (`FromWindow` / `GetDescendants` /
  `Role` / `Name` / `Handle` / `DoDefaultAction`), enough for the two MSPaint helpers. It
  needs none of the v1 chain (`WinApiMacros`/`GUID`/`ComVar`/`RemoteBuff`/`Str`), which
  existed mostly for `ControlGetName`.

Things the port had to change semantically, worth knowing before editing it:

- Hotkey bodies are functions in v2 **and v2 removed super-globals**, so a top-level `global`
  declaration does not reach into them. Every body touching the shared state (`DoubleAlt`,
  `DoubleCtrl`, `notepadScrollAmount`, `Mon`) declares it `global` on its own first line —
  reading, not just writing, requires the declaration. Add one when writing a new hotkey that
  touches them.
- v1 auto-exempted any hotkey containing `Suspend`; v2 requires `#SuspendExempt`, without
  which the suspend toggle cannot resume itself.
- v2 throws where v1 silently no-opped, so calls acting on the window under the cursor are
  wrapped in `try` — a window closing mid-drag would otherwise raise an error dialog.
- `SysGet, MonN, Monitor` became `MonitorGet`, which throws for an absent monitor instead of
  leaving a blank variable; `WindowMove()` still bails unless two monitors were found, and is
  still deliberately a two-monitor routine.
- `ControlGetFocus` returns an HWND in v2 (not a ClassNN), and `ControlSetText`/`WinMove`
  take their parameters in a different order.
- `MSPaintEnterTextMode()` in v1 guarded its body with a `WinExist()` test on an undefined
  local, which always evaluated false so the body always ran; v2 would throw on that, so the
  dead guard was dropped. `IsWow64Process()` was unreferenced and was dropped too.

Call statements in the port are written with explicit parentheses (`SetKeyDelay(-1, -1)`,
not `SetKeyDelay -1, -1`) to sidestep v2's command-style parsing ambiguity. Keep that style.

## WSLg windows (RAIL proxies)

WSLg Linux windows (gedit etc.) are not real Windows windows. Weston inside WSL owns them
and the RDP client projects them onto the desktop as RemoteApp/RAIL proxies — window class
`RAIL_WINDOW`, process `msrdc.exe` (`C:\Program Files\WSL\msrdc.exe`). Two things about
them break naive window-management code, both verified with `WSLg_Probe.ahk`:

- **`WinMove` works — position *and* size — but applies asynchronously.** A `WinGetPos`
  immediately after a `WinMove` returns the *old* rect. Any loop that re-reads the window
  each frame computes from stale numbers and fights itself, which looks exactly like "the
  hotkey does nothing". This is why v1's ALT+drag worked on WSLg (it captures the rect once
  and applies a cumulative offset) while ALT+right-drag resize did not (it re-read every
  iteration and stepped incrementally). Write these loops stateless: capture the original
  rect once, compute from the cumulative mouse delta.
- **`WinMaximize` is a request Weston fulfils, and it picks the monitor itself**, ignoring
  where the local proxy sits — so maximize and throw land on the wrong display. The v2 port
  therefore never calls `WinMaximize`/`WinRestore` on a RAIL window; `RailMaximizeTo()` sizes
  it to the target monitor's work area directly and `RailMaxRect` tracks the un-maximize rect,
  because a hand-sized window still reports `MinMax = 0` and Windows cannot track it.

`IsRailWindow()` gates all of this on the process name, so ordinary windows keep the original
code path untouched.

## Structure of a script file

Each of the big scripts follows the same layout, and new code should go in the matching
section:

1. Directives and globals (`#SingleInstance Force`, `SetWinDelay`, `SysGet, Mon1/Mon2, Monitor`).
2. Plain helper functions (`LaunchPaint()`, `FixYTURL()`, `IsWow64Process()`).
3. Global hotkeys — suspend/reload, launchers, then the KDE core (`!LButton`, `!RButton`,
   `!MButton`, `~Alt`, `~Ctrl`, `~^LButton`).
4. The `windowmove:` subroutine (`gosub`-style label) for two-monitor window throwing.
5. Per-application `#IfWinActive` blocks.
6. Function definitions that depend on `Lib/` (e.g. `MSPaintFontSizeAdd`).

### Gotchas that bite in this file layout

- **Every `#IfWinActive` block must be closed with a bare `#IfWinActive`** (or `#If` for
  the `#If WinActive(...)` form). Forgetting it silently scopes every subsequent hotkey to
  that application.
- **Double-ALT / double-CTRL chords** are implemented by `~Alt::` / `~Ctrl::` setting
  `DoubleAlt` / `DoubleCtrl` from `A_PriorHotKey` and `A_TimeSincePriorHotkey < 400`. Any
  hotkey that consumes the flag must reset it to `false`, or the next click misfires.
- **Monitor logic in `windowmove:` assumes exactly two monitors**, reading `Mon1*`/`Mon2*`
  captured by `SysGet` at load time — it does not react to display changes. AHK numbers
  active monitors 1,2 even when Windows calls them 2,3.
- **Machine-specific hardcoded paths** exist throughout (`#n` → Notepad++ install path,
  `#k` → `K:\projects`, the absolute `#Include`s). Leave them unless asked; they are the
  user's own environment, not bugs.

## Lib/

Vendored third-party v1 libraries, resolved via AHK's `%A_ScriptDir%\Lib` search path:

- `Notify.ahk` — gwarble's tray toast; the repo's only user feedback channel
  (`Notify("Suspended", "", 3)`), used for suspend/resume/reload and CapsLock state.
- `AccessibleObject.ahk` — MSAA/IAccessible wrapper used to reach controls no AHK command
  can (the MSPaint ribbon "Font Size" edit box). It pulls in `WinApiMacros.ahk`,
  `GUID.ahk`, `ComVar.ahk`, `RemoteBuff.ahk` (cross-process memory), and `Str.ahk` via its
  own `#Include`s — include `AccessibleObject` and the rest follow.

Application automation otherwise goes through COM (`ComObjActive("Excel.Application")`,
`"PowerPoint.Application"`), always wrapped in `Try` since the app may not be running.

## Conventions

- Inline comments credit the source of each contributed block (`; from Michael Nelson`,
  `; from Nour Nasser`, `; from Rodfell on AHK forum`, `; from chatGPT`) and often carry a
  date (`; 2/12/26`). Keep attributing new blocks the same way.
- Commit messages are short and describe the hotkey added, frequently naming the source of
  the idea — e.g. "Add ALT+Wheel and ALT+SHIFT+Wheel to use with S-Edit".
- Superseded code is commented out in place rather than deleted, usually with the date and
  the reason it was replaced. Follow that rather than removing history.
