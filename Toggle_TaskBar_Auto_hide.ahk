#Requires AutoHotkey v2.0
#SingleInstance Force


; Hotkey: Ctrl+Win+T   (change if you want)
#^t::ToggleTaskbarAutoHide()

ToggleTaskbarAutoHide() {
    ; APPBARDATA and flags for SHAppBarMessage (ABM_GETSTATE / ABM_SETSTATE)
    static ABM_GETSTATE := 0x00000004
    static ABM_SETSTATE := 0x0000000A
    static ABS_AUTOHIDE := 0x00000001

    ; Build APPBARDATA struct (size from MS docs)
    cbSize := 2 * A_PtrSize + 2 * 4 + 16 + A_PtrSize
    appBar := Buffer(cbSize, 0)
    NumPut("UInt", cbSize, appBar, 0)  ; cbSize field

    ; Ask shell for current taskbar state
    state := DllCall(
        "Shell32\SHAppBarMessage",
        "UInt", ABM_GETSTATE,
        "Ptr",  appBar,
        "UInt"  ; return type
    )

    if (state = "") {
        MsgBox "Could not query taskbar state (SHAppBarMessage failed)."
        return
    }

    ; Flip the AUTOHIDE bit and write it back
    newState := state ^ ABS_AUTOHIDE
    NumPut("UInt", newState, appBar, cbSize - A_PtrSize) ; lParam field

    DllCall(
        "Shell32\SHAppBarMessage",
        "UInt", ABM_SETSTATE,
        "Ptr",  appBar
    )
}
