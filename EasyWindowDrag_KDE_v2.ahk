#Requires AutoHotkey v2.0
#SingleInstance Force

; 9/1/26 -- AutoHotkey v2 port of EasyWindowDrag_KDE.ahk.
; The v1 script is kept as-is; run one or the other, not both.
; Depends on Lib\Notify_v2.ahk and Lib\Acc_v2.ahk (v2 stand-ins for the v1
; Notify.ahk and AccessibleObject.ahk, which cannot be included from v2).
;
; Behaviour notes for this port:
;  * Hotkey bodies are functions in v2, and v2 has no super-globals, so every
;    body that touches the shared state (DoubleAlt, DoubleCtrl, Mon,
;    notepadScrollAmount) declares it `global` on its first line.
;  * v2 raises errors where v1 silently did nothing, so window calls that act on
;    whatever is under the cursor are wrapped in try -- a window closing
;    mid-drag used to be a no-op and would otherwise now pop an error dialog.
;  * v1 auto-exempted any hotkey containing Suspend; v2 needs #SuspendExempt,
;    or the suspend/resume toggle could not resume itself.
;  * IsWow64Process() was dead code and is dropped.

#Include <Notify_v2>
#Include <Acc_v2>

InstallKeybdHook()   ; v2 dropped the #InstallKeybdHook directive; it is a function now
SendMode("Input")
SetWorkingDir(A_ScriptDir)
SetControlDelay(-1)
CoordMode("Mouse", "Screen")

; This is the setting that runs smoothest on my
; system. Depending on your video card and cpu
; power, you may want to raise or lower this value.
; (v1 set -1 up top and 2 further down; 2 was what actually took effect.)
SetWinDelay(2)

; from Michael Nelson
; The number of times alt-scrolling will cause in Notepad++
notepadScrollAmount := 3

; Double-Alt / Double-Ctrl chord flags, set by the ~Alt and ~Ctrl hotkeys.
DoubleAlt := false
DoubleCtrl := false

; if you have 2 monitors active - regardless of whether ID#1 is disconnected or not, AHK always sees 1,2, not 2,3 that Windows uses..
; v2 note: MonitorGet throws for a monitor that is not there, where v1's SysGet
; just left Mon2Left blank -- hence the try. WindowMove() still bails unless
; there are two, exactly as before.
Mon := []
loop MonitorGetCount() {
    try {
        MonitorGet(A_Index, &monLeft, &monTop, &monRight, &monBottom)
        Mon.Push({ left: monLeft, top: monTop, right: monRight, bottom: monBottom })
    }
}

; Easy Window Dragging -- KDE style -- by Jonny
; https://www.autohotkey.com
; This script makes it much easier to move or resize a window: 1) Hold down
; the ALT key and LEFT-click anywhere inside a window to drag it to a new
; location; 2) Hold down ALT and RIGHT-click-drag anywhere inside a window
; to easily resize it; 3) Press ALT twice, but before releasing it the second
; time, left-click to minimize the window under the mouse cursor, right-click
; to maximize it, or middle-click to close it.
; This script was inspired by and built on many like it
; in the forum. Thanks go out to ck, thinkstorm, Chris,
; and aurelian for a job well done.
;
; The Double-Alt modifier is activated by pressing
; Alt twice, much like a double-click. Hold the second
; press down until you click.
;
; The shortcuts:
;  Alt + Left Button  : Drag to move a window.
;  Alt + Right Button : Drag to resize a window.
;  Double-Alt + Left Button   : Minimize a window.
;  Double-Alt + Right Button  : Maximize/Restore a window.
;  Double-Alt + Middle Button : Close a window.
;
; You can optionally release Alt after the first
; click rather than holding it down the whole time.


; ============================== helpers ==============================

; from John T - autohotkey can't find snipping tool
LaunchSnippingTool() {
    if WinExist("Snipping Tool") {
        WinActivate()
    } else {
        Run("SnippingTool.exe")
        if (!WinWait("Snipping Tool", , 5))   ; v1 waited forever; 5s keeps the thread from hanging
            return
        WinActivate("Snipping Tool")
    }
    SendInput("^+n")
}

LaunchPaint() {
    if WinExist("Untitled - Paint") {
        if WinActive("Untitled - Paint")
            Run("mspaint.exe")
        else
            WinActivate("Untitled - Paint")
    } else {
        Run("mspaint.exe")
        if (WinWait("Untitled - Paint", , 5))
            WinActivate("Untitled - Paint")
    }
}

FixYTURL() {
    if (InStr(A_Clipboard, "https://www.youtube.com") = 1)
        A_Clipboard := StrReplace(A_Clipboard, "watch?v=", "embed/")
}


; --- WSLg / RemoteApp (RAIL) windows -----------------------------------------
; 9/1/26. WSLg does not create real Windows windows: each Linux window lives in
; the Weston compositor inside WSL and is projected onto the desktop by the RDP
; client (msrdc.exe) as a RemoteApp/RAIL proxy (class RAIL_WINDOW).
;
; What was established by experiment (see WSLg_Resize_Probe.ahk):
;  * The REMOTE side owns the window size. The local proxy can be shrunk below
;    the remote surface -- the content is simply clipped, right and bottom edges
;    vanish -- but it can never be grown. Any size divergence breaks rendering
;    and input routing until the original size is restored: the window looks
;    hung while the Linux app underneath is still running normally.
;  * Position changes via WinMove DO apply locally and are harmless, which is
;    why ALT+drag works. They do not appear to be reported to the server, so
;    Weston still maximizes by its own reckoning of which output the window is
;    on -- hence maximize landing on the wrong monitor.
;  * WinMaximize/WinRestore go through the RAIL protocol and keep the window
;    healthy, even when they pick the wrong monitor.
;
; So: never change the SIZE of one of these windows from the script. Resize is
; skipped, maximize uses the real WinMaximize, and a throw moves without
; maximizing. Fixing resize properly means driving the OS's own modal size loop
; (WM_SYSCOMMAND / SC_SIZE), which is what RAIL forwards -- still under test.

IsRailWindow(hwnd) {
    try {
        exe := WinGetProcessName("ahk_id " hwnd)
        return (exe = "msrdc.exe" || exe = "mstsc.exe")
    } catch
        return false
}

; Moving one of these with WinMove is what puts it in the "bad" state: the
; local proxy slides, the server keeps the old geometry, and the client's
; hit-test regions go stale -- the resize borders stop responding, maximize
; resolves against the wrong monitor, restore lands wrong. Dragging the window
; by a real gesture re-syncs it.
;
; So hand the drag to Windows' own move loop. WM_NCLBUTTONDOWN with HTCAPTION
; is the genuine frame entry point (it works even though these windows have no
; WS_CAPTION -- the loop only needs the hit-test code), the RDP client hooks it,
; and it tracks the physical left button, which the ALT+drag gesture is already
; holding down. Windows then drives the drag and reports the final position to
; the server on release.
RailDragMove(hwnd) {
    static WM_NCLBUTTONDOWN := 0x00A1
    static HTCAPTION := 2
    MouseGetPos(&mx, &my)
    lParam := ((my & 0xFFFF) << 16) | (mx & 0xFFFF)
    try PostMessage(WM_NCLBUTTONDOWN, HTCAPTION, lParam, , "ahk_id " hwnd)
}

; Win+Shift+Left/Right is Windows' own move-to-next-monitor, and the server
; tracks it correctly where our WinMove does not.
RailThrow(hwnd, goRight) {
    try WinActivate("ahk_id " hwnd)
    Sleep(80)
    Send(goRight ? "#+{Right}" : "#+{Left}")
}

RailMonitorOf(hwnd) {
    try
        WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
    catch
        return MonitorGetPrimary()
    cx := x + w / 2
    cy := y + h / 2
    loop MonitorGetCount() {
        MonitorGet(A_Index, &l, &t, &r, &b)
        if (cx >= l && cx < r && cy >= t && cy < b)
            return A_Index
    }
    return MonitorGetPrimary()
}

; WinGetMinMax reads stale on these windows, so ALT+F1 kept re-maximizing
; instead of restoring (the "three presses to restore" symptom). Fall back to
; measuring: if the window covers essentially all of its monitor work area,
; treat it as maximized. That also stays correct when the window was maximized
; or restored from its own title bar, which a flag of ours could not track.
RailIsMaximized(hwnd) {
    try {
        if (WinGetMinMax("ahk_id " hwnd) = 1)
            return true
        WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        MonitorGetWorkArea(RailMonitorOf(hwnd), &l, &t, &r, &b)
        return (w >= (r - l) * 0.95 && h >= (b - t) * 0.95)
    } catch
        return false
}

; Weston resolves a maximize against the monitor the SERVER thinks the window
; is on, and it only learns geometry from real client-driven gestures -- never
; from our WinMove. Sliding the maximized result across to the right monitor
; was tried and is WORSE: moving an already-maximized proxy desyncs it from the
; remote surface and the window hangs. So we let the maximize land wherever
; Weston puts it. Moving a window by its title bar first (a real gesture) does
; teach the server, and then ALT+F1 maximizes on the correct monitor.
RailToggleMax(hwnd) {
    ; Braces are load-bearing: v2's Try has its own optional Else clause, so an
    ; unbraced `try` as the if-branch swallows the else.
    ; Win+Up / Win+Down are the shell's own snap commands. They go through the
    ; path the RDP client and the server both track, where WinMaximize and
    ; WinRestore leave the window half-synced (maximize on the wrong monitor,
    ; restore that does nothing). Win+Down only restores when the window really
    ; is maximized -- on a restored window it minimizes -- so the measured
    ; RailIsMaximized check above is what keeps this safe.
    try WinActivate("ahk_id " hwnd)
    Sleep(80)
    if (RailIsMaximized(hwnd))
        Send("#{Down}")
    else
        Send("#{Up}")
}


; ============================== Ananth ==============================

; v1 auto-exempted hotkeys containing Suspend; v2 does not, so say so or the
; script can never be resumed from the keyboard.
#SuspendExempt
^!F9::
^!RButton::
{
    Suspend(-1)
    Notify(A_IsSuspended ? "Suspended" : "Resumed", "", 3)
}
#SuspendExempt False

; NOTE: not suspend-exempt, so (as in v1) this does nothing while suspended.
^!+MButton::
{
    Notify("Reloading", "", 3)
    Sleep(1000)  ; Give 1 second to read the message
    Reload()
}

^#T::FixYTURL()	; CTRL WIN T
#p::LaunchPaint()
#Escape::LaunchPaint()
#s::LaunchSnippingTool()
; tweaked 7/10/16
#n::Run('"C:\Program Files\Notepad++\notepad++.exe"')
#k::Run('explore "K:\projects"')

^Escape::
{
    try WinMove(0, 0, , , "A")
}

; paste clipboard
^!v::
{
    if (StrLen(A_Clipboard) > 0) {
        SetKeyDelay(15)
        Send("{Text}" A_Clipboard)
        A_Clipboard := ""
        SetKeyDelay(0)
    }
}

; 1/25/16 -- how often did you use the old function from M$?
#Down::
{
    global DoubleAlt
    MouseGetPos(, , &id)
    ; This message is mostly equivalent to WinMinimize,
    ; but it avoids a bug with PSPad.
    try PostMessage(0x112, 0xF020, , , "ahk_id " id)
    DoubleAlt := false
}


; ============================== KDE window drag ==============================

!LButton::
{
    global DoubleAlt
    if (DoubleAlt) {
        MouseGetPos(, , &id)
        ; This message is mostly equivalent to WinMinimize,
        ; but it avoids a bug with PSPad.
        try PostMessage(0x112, 0xF020, , , "ahk_id " id)
        DoubleAlt := false
        return
    }

    ; Get the initial mouse position and window id, and
    ; abort if the window is maximized.
    MouseGetPos(&x1, &y1, &id)
    try {
        if (WinGetMinMax("ahk_id " id))
            return
        ; Get the initial window position.
        WinGetPos(&winX1, &winY1, , , "ahk_id " id)
    } catch
        return

    ; WSLg: let Windows drive the drag, so the server learns the new position.
    if (IsRailWindow(id)) {
        RailDragMove(id)
        return
    }

    while GetKeyState("LButton", "P") {   ; loop until the button is released
        MouseGetPos(&x2, &y2)             ; Get the current mouse position.
        ; Offset from the initial mouse position, applied to the window position.
        try
            WinMove(winX1 + (x2 - x1), winY1 + (y2 - y1), , , "ahk_id " id)
        catch
            break
    }
}

!RButton::
{
    global DoubleAlt
    if (DoubleAlt) {
        MouseGetPos(, , &id)
        ; Toggle between maximized and restored state.
        try {
            if (IsRailWindow(id))
                RailToggleMax(id)
            else if (WinGetMinMax("ahk_id " id))
                WinRestore("ahk_id " id)
            else
                WinMaximize("ahk_id " id)
        }
        DoubleAlt := false
        return
    }

    ; Get the initial mouse position and window id, and
    ; abort if the window is maximized.
    MouseGetPos(&x1, &y1, &id)
    try {
        if (WinGetMinMax("ahk_id " id))
            return
        ; Get the initial window position and size.
        WinGetPos(&winX, &winY, &winW, &winH, "ahk_id " id)
    } catch
        return

    ; Define the window region the mouse is currently in.
    ; The four regions are Up and Left, Up and Right, Down and Left, Down and Right.
    winLeft := (x1 < winX + winW / 2) ? 1 : -1
    winUp := (y1 < winY + winH / 2) ? 1 : -1

    ; Every frame is computed from the ORIGINAL rect plus the cumulative mouse
    ; delta -- the window is never read back. v1 re-read WinGetPos each
    ; iteration and stepped incrementally, which is fine for a window that
    ; resizes synchronously but breaks on WSLg/RAIL windows: the read-back is
    ; stale, so the loop fought itself and the resize appeared dead. (ALT+drag
    ; always worked on WSLg precisely because it never re-read the window.)
    ; Mathematically identical to v1 for ordinary windows.
    while GetKeyState("RButton", "P") {
        MouseGetPos(&x2, &y2)
        dx := x2 - x1   ; Cumulative offset from the initial mouse position.
        dy := y2 - y1
        ; Then, act according to the defined region.
        newX := winX + (winLeft + 1) / 2 * dx
        newY := winY + (winUp + 1) / 2 * dy
        newW := winW - winLeft * dx
        newH := winH - winUp * dy
        try
            WinMove(newX, newY, newW, newH, "ahk_id " id)
        catch
            break
    }
}

; "Alt + MButton" may be simpler, but I
; like an extra measure of security for
; an operation like this.
!MButton::
{
    global DoubleAlt
    if (DoubleAlt) {
        MouseGetPos(, , &id)
        try WinClose("ahk_id " id)
        DoubleAlt := false
    }
}

; This detects "double-clicks" of the alt key.
~Alt::
{
    global DoubleAlt
    DoubleAlt := (A_PriorHotkey = "~Alt" && A_TimeSincePriorHotkey < 400)
    Sleep(0)
    KeyWait("Alt")  ; This prevents the keyboard's auto-repeat feature from interfering.
}

; der Hero Herr Gwarble : http://www.gwarble.com/ahk/Notify/
~CapsLock::
{
    if GetKeyState("CapsLock", "T")
        Notify("CAPS", "CAPS ON", 3)
    else
        Notify("CAPS", "caps off", 3)
}

;; preparing for Rodfell's throwing script
; This detects "double-clicks" of the CTRL key.
~Ctrl::
{
    global DoubleCtrl
    DoubleCtrl := (A_PriorHotkey = "~Ctrl" && A_TimeSincePriorHotkey < 400)
    Sleep(0)
    KeyWait("Ctrl")  ; This prevents the keyboard's auto-repeat feature from interfering.
}

;;;;;;;;;;;;;;; from Rodfell on AHK forum
; ~ prefix added by Houli Wang
~^LButton::
{
    global DoubleCtrl
    KeyWait("LButton")
    if (DoubleCtrl) {
        DoubleCtrl := false
        MouseGetPos(, , &windowToMove)
        WindowMove(windowToMove)
    }
}

; Was the "windowmove:" label + gosub in v1; Gosub and labels are gone in v2.
; Still deliberately a two-monitor routine, as the original was.
WindowMove(windowToMove) {
    global Mon
    if (Mon.Length < 2)
        return

    try {
        WinGetPos(&x1, &y1, &w1, &h1, "ahk_id " windowToMove)
        winState := WinGetMinMax("ahk_id " windowToMove)
    } catch
        return

    ; works out if centre of window is on monitor 1 (m1=1) or monitor 2 (m1=2)
    cx := x1 + w1 / 2
    cy := y1 + h1 / 2
    m1 := (cx > Mon[1].left && cx < Mon[1].right && cy > Mon[1].top && cy < Mon[1].bottom) ? 1 : 2
    m2 := (m1 = 1) ? 2 : 1   ; m2 is the monitor the window will be moved to
    from := Mon[m1]
    to := Mon[m2]

    fromW := Abs(from.right - from.left)
    fromH := Abs(from.bottom - from.top)
    toW := Abs(to.right - to.left)
    toH := Abs(to.bottom - to.top)

    ratioX := (fromW - w1 < 5) ? 0 : Abs((x1 - from.left) / (fromW - w1))   ; where the window fits on x axis
    ratioY := (fromH - h1 < 5) ? 0 : Abs((y1 - from.top) / (fromH - h1))    ; where the window fits on y axis
    x2 := to.left + ratioX * (toW - w1)   ; where the window will fit on x axis in normal situation
    y2 := to.top + ratioY * (toH - h1)
    w2 := w1
    h2 := h1   ; width and height will stay the same when moving unless reason not to lower in script

    ; if x axis takes up whole axis OR won't fit on new screen
    if (fromW - w1 < 5 || Abs(to.right - to.left - w1) < 5) {
        x2 := to.left
        w2 := toW
    }
    if (fromH - h1 < 5 || toH - h1 < 5) {
        y2 := to.top
        h2 := toH
    }

    ; RAIL windows: move only. Their size must not be touched, and a maximize
    ; would be resolved by Weston against its own idea of the current output.
    if (IsRailWindow(windowToMove)) {
        RailThrow(windowToMove, to.left > from.left)
        return
    }

    try {
        if (winState) {   ; move maximized window
            WinRestore("ahk_id " windowToMove)
            WinMove(to.left, to.top, , , "ahk_id " windowToMove)
            WinMaximize("ahk_id " windowToMove)
        } else {
            ; adjustments for windows that are not fully on the initial monitor
            if (x1 < from.left)
                x2 := to.left
            if (x1 + w1 > from.right)
                x2 := to.right - w2
            if (y1 < from.top)
                y2 := to.top
            if (y1 + h1 > from.bottom)
                y2 := to.bottom - h2
            WinMove(x2, y2, w2, h2, "ahk_id " windowToMove)   ; move non-maximized window
            WinMaximize("ahk_id " windowToMove)
        }
    }
}


; ============================== per-app ==============================

;;;;;;;;;;;;;;; from Rodfell on AHK forum
#HotIf WinActive("ahk_class XLMAIN")
^!WheelUp::Send("^{PgUp}")
^!WheelDown::Send("^{PgDn}")
+WheelDown::SendInput("{ScrollLock}{Right}{ScrollLock}")
+WheelUp::SendInput("{ScrollLock}{Left}{ScrollLock}")
^BackSpace::Send("^+{Left}{BackSpace}")	; to get delete word backward in Excel
#HotIf

#HotIf WinActive("ahk_class Chrome_WidgetWin_1")
^!WheelUp::Send("^{PgUp}")
^!WheelDown::Send("^{PgDn}")
#HotIf

#HotIf WinActive("ahk_class ApplicationFrameWindow")
^!WheelUp::Send("^{PgUp}")
^!WheelDown::Send("^{PgDn}")
#HotIf

#HotIf WinActive("ahk_exe POWERPNT.EXE")
^;::
{
    SendInput(FormatTime(, "MM/dd/yy"))
}

; trying Areeb's solution
^+B::
{
    try {
        ppt := ComObjActive("PowerPoint.Application")
        if (ppt.ActiveWindow.Selection.Type = 2) {
            try ppt.ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Font.Color.RGB := 0xFF0000
        }
        if (ppt.ActiveWindow.Selection.Type = 3)
            ppt.ActiveWindow.Selection.TextRange.Font.Color.RGB := 0xFF0000
    }
}

^+C::
{
    try {
        ppt := ComObjActive("PowerPoint.Application")
        if (ppt.ActiveWindow.Selection.Type = 2) {
            try ppt.ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Font.Name := "consolas"
        }
        if (ppt.ActiveWindow.Selection.Type = 3)
            ppt.ActiveWindow.Selection.TextRange.Font.Name := "consolas"
    }
}
#HotIf

$!Escape::
{
    prevMatchMode := A_TitleMatchMode
    SetTitleMatchMode(2)
    if WinActive("Citrix XenApp")
        Send("+{Escape}")
    else
        try WinMoveBottom("A")
    SetTitleMatchMode(prevMatchMode)
}

#HotIf WinActive("ahk_class mintty")
^+L::Send("clear{enter}")
#HotIf

#HotIf WinActive("Untitled - Paint")
#F4::
{
    try ProcessClose(WinGetPID("A"))
}
#HotIf

; Inspired by using Alt-F1 on KDE to toggle maximization :)
$!F1::
{
    global DoubleAlt
    prevMatchMode := A_TitleMatchMode
    SetTitleMatchMode(2)
    if WinActive("Citrix XenApp") {
        Send("!{F1}")
    } else {
        MouseGetPos(, , &id)
        ; Toggle between maximized and restored state.
        try {
            if (IsRailWindow(id))
                RailToggleMax(id)
            else if (WinGetMinMax("ahk_id " id))
                WinRestore("ahk_id " id, , "ATL-WIN7")   ; ExcludeTitle : "ATL-WIN7"
            else
                WinMaximize("ahk_id " id)
        }
        DoubleAlt := false
    }
    SetTitleMatchMode(prevMatchMode)
}

; from Michael Nelson
; Trigger on alt scroll up (!WheelUp), it will not trigger itself ($) and AHK will not cause the key to be supressed (~)
~$!WheelUp::
{
    global notepadScrollAmount
    if WinActive("ahk_class Notepad++") {          ; if Notepad++
        loop notepadScrollAmount - 1               ; Loop X many more times to meet the notepadScrollAmount desired
            Send("{WheelUp}")                       ; Sends a wheelup command to windows
    }
}

; Same as wheelup but for wheeldown
~$!WheelDown::
{
    global notepadScrollAmount
    if WinActive("ahk_class Notepad++") {
        loop notepadScrollAmount - 1
            Send("{WheelDown}")
    }
}

; from Nour Nasser, UPWK
#HotIf WinActive("ahk_class MSPaintApp")
^[::
{
    SetKeyDelay(-1, -1)
    MSPaintFontSizeAdd(-1)
}

^]::
{
    SetKeyDelay(-1, -1)
    MSPaintFontSizeAdd(1)
}

^!d::
{
    try {
        if (!StatusBarGetText(1, "A"))
            return
    } catch
        return

    dateTime := FormatTime(, "MM/dd/yyyy,HH:mm")   ; update format if needed

    MSPaintEnterTextMode()
    Click()
    Sleep(100)
    ; v2's ControlGetFocus returns an HWND, not a ClassNN.
    try {
        focused := ControlGetFocus("A")
        if (focused)
            ControlSetText(dateTime, focused)      ; v2 takes the text first
    }
}
#HotIf

MSPaintFontSizeAdd(Value) {
    static FontSizeEditHandle := 0

    if (!FontSizeEditHandle || !WinExist("ahk_id " FontSizeEditHandle)) {
        FontSizeEditHandle := 0
        win := AccessibleObject.FromWindow(WinActive("ahk_class MSPaintApp"))
        if (!win)
            return
        for ctl in win.GetDescendants() {
            if (ctl.Role = 0x2A && ctl.Name = "Font Size") {   ; 0x2A = ROLE_SYSTEM_TEXT
                FontSizeEditHandle := ctl.Handle
                break
            }
        }
        if (!FontSizeEditHandle)
            return
    }

    try {
        ControlFocus(FontSizeEditHandle)
        fontSize := ControlGetText(FontSizeEditHandle)
        if (!IsInteger(fontSize))
            return
        ControlSetText(Max(1, Integer(fontSize) + Value), FontSizeEditHandle)
        ControlSend("{Enter}", FontSizeEditHandle)
    } catch {
        FontSizeEditHandle := 0   ; ribbon rebuilt underneath us; re-find next time
    }
}

MSPaintEnterTextMode() {
    ; The v1 version guarded this with a WinExist() test on an undefined local,
    ; which always evaluated false, so the body always ran. v2 would throw on
    ; that undefined variable, so the dead guard is simply dropped.
    win := AccessibleObject.FromWindow(WinActive("ahk_class MSPaintApp"))
    if (!win)
        return
    for ctl in win.GetDescendants() {
        if (ctl.Role = 0x2B && ctl.Name = "Text") {   ; 0x2B = ROLE_SYSTEM_PUSHBUTTON
            ctl.DoDefaultAction()
            return
        }
    }
}

#HotIf WinActive("ahk_exe EXCEL.EXE")
; want : CTRL+ALT+LeftClick - select entire row
; NOTE : to prevent Research pane from opening, in your PERSONAL.XLSB,
; in ThisWorkbook, under WorkBook_Open() : Application.CommandBars("Research").Enabled = False
^!LButton::
{
    Click("Left")   ; Click the left mouse button
    try {
        XL := ComObjActive("Excel.Application")
        XL.ActiveCell.EntireRow.Select()
    }
}
#HotIf
