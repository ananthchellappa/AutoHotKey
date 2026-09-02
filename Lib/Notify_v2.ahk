#Requires AutoHotkey v2.0
;
; Notify_v2.ahk -- small stacking toast for AutoHotkey v2.
;
; Stands in for gwarble's v1 Notify() (Lib\Notify.ahk), which is 400+ lines of
; v1 Gui commands and does not port cleanly. Only the call shape this repo
; actually uses is reproduced:
;
;   Notify(Title, Message, Duration)      ; Duration in seconds
;
; Differences from gwarble's original, on purpose:
;   * Duration 0 means "stay until clicked" (the v1 version also flashed).
;   * The v1 "negative duration = ExitApp on click/timeout" behaviour is NOT
;     implemented; nothing in this repo used it.
;   * The Options string is accepted and ignored, so existing calls still parse.
;
; Toasts stack upward from the bottom-right of the primary monitor's work area,
; newest at the bottom. Click one to dismiss it early.
;

Notify(Title := "", Message := "", Duration := 30, Options := "") {
    return NotifyToast.Show(Title, Message, Duration)
}

class NotifyToast
{
    static Active := []
    static Width := 280
    static EdgeGap := 12
    static StackGap := 8

    static Show(Title, Message, Duration) {
        ; -DPIScale keeps Move/GetPos in real pixels, which is what the
        ; work-area maths below assumes.
        g := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale +E0x08000000")  ; WS_EX_NOACTIVATE
        g.BackColor := "202020"
        g.MarginX := 14
        g.MarginY := 12

        textWidth := this.Width - 2 * g.MarginX
        dismiss := (*) => NotifyToast.Dismiss(g)

        if (Title != "") {
            g.SetFont("s10 Bold cFFFFFF", "Segoe UI")
            g.Add("Text", "w" textWidth, Title).OnEvent("Click", dismiss)
        }
        if (Message != "") {
            g.SetFont("s9 Norm cDDDDDD", "Segoe UI")
            g.Add("Text", "w" textWidth, Message).OnEvent("Click", dismiss)
        }
        if (Title = "" && Message = "")
            g.Add("Text", "w" textWidth, " ")

        ; Show hidden at an explicit position first, so a later Show() cannot
        ; auto-centre the window.
        g.Show("Hide AutoSize x0 y0")

        this.Active.Push(g)
        this.Restack()
        g.Show("NoActivate")

        if (Duration > 0)
            SetTimer(dismiss, -Round(Duration * 1000))

        return g
    }

    static Dismiss(g) {
        for i, item in this.Active {
            if (item = g) {
                this.Active.RemoveAt(i)
                break
            }
        }
        try
            g.Destroy()
        this.Restack()
    }

    static Restack() {
        try
            MonitorGetWorkArea(, &left, &top, &right, &bottom)
        catch
            return

        ; Oldest sits at the bottom and each newer one stacks above it, so the
        ; most recent notification is always on top of the older ones.
        y := bottom - this.EdgeGap
        for g in this.Active {
            try {
                g.GetPos(, , &w, &h)
                y -= h
                g.Move(right - w - this.EdgeGap, y)
                y -= this.StackGap
            }
        }
    }
}
