#SingleInstance Force  ; from Nour Nasser
#InstallKeybdHook ; from Nour
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
; from Michael Nelson
; The number of times alt-scrolling will cause in Notepad++
notepadScrollAmount := 3

SetBatchLines -1 ; Nour Nasser
SetControlDelay -1
SetWinDelay -1 ; from Nour

; >>>>>>>>>>>>>>>>>>>>>>>>> others, please see my custom adds at end and
; >>>>>>>>>>>>>>>>>>>>>>>>> delete if necessary..
; Easy Window Dragging -- KDE style (requires XP/2k/NT) -- by Jonny
; http://www.autohotkey.com
; This script makes it much easier to move or resize a window: 1) Hold down
; the ALT key and LEFT-click anywhere inside a window to drag it to a new
; location; 2) Hold down ALT and RIGHT-click-drag anywhere inside a window
; to easily resize it; 3) Press ALT twice, but before releasing it the second
; time, left-click to minimize the window under the mouse cursor, right-click
; to maximize it, or middle-click to close it.
; This script was inspired by and built on many like it
; in the forum. Thanks go out to ck, thinkstorm, Chris,
; and aurelian for a job well done.
; Change history:
; November 07, 2006: Optimized resizing code in !RButton, courtesy of bluedawn.
; February 05, 2006: Fixed double-alt (the ~Alt hotkey) to work with latest versions of AHK.
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
; from John T - autohotkey can't find snipping tool
; Determines if we are running a 32 bit program (autohotkey) on 64 bit Windows
;;;; SetTitleMatchMode, 2   ;-- set this 4/7/15 to trap "Notepad"; later decided to
;;;;                        ;-- look for err_final.txt - Notepad
IsWow64Process()
{
   hProcess := DllCall("kernel32\GetCurrentProcess")
   ret := DllCall("kernel32\IsWow64Process", "UInt", hProcess, "UInt *", bIsWOW64)
   return ret & bIsWOW64
}

FixYTURL()
{
	If (InStr(Clipboard, "https://www.youtube.com") = 1)
	{
		Clipboard := StrReplace(Clipboard, "watch?v=", "embed/")
	}
}

; Launch snipping tool using correct path based on 64 bit or 32 bit Windows
LaunchSnippingTool()
{
 IfWinExist , Snipping Tool
 {
  WinActivate , Snipping Tool
  SendInput ^+n
 } else {
  Run SnippingTool.exe
  WinWait , Snipping Tool
  WinActivate , Snipping Tool
  SendInput ^+n
 }
}
LaunchPaint()
{
 IfWinExist , Untitled - Paint
 {
  IfWinActive, Untitled - Paint
  {
   Run mspaint.exe
  } else {
   WinActivate , Untitled - Paint
  }
 } else {
  Run mspaint.exe
  WinWait , Untitled - Paint
  WinActivate , Untitled - Paint
 }
}

If (A_AhkVersion < "1.0.39.00")
{
    MsgBox,20,,This script may not work properly with your version of AutoHotkey. Continue?
    IfMsgBox,No
    ExitApp
}


; This is the setting that runs smoothest on my
; system. Depending on your video card and cpu
; power, you may want to raise or lower this value.
SetWinDelay,2
CoordMode,Mouse
; if you have 2 monitors active - regardless of whether ID#1 is disconnected or not, AHK always sees 1,2, not 2,3 that Windows uses..
SysGet, Mon1, Monitor, 1
SysGet, Mon2, Monitor, 2
 
; Ananth
^!F9::
^!RButton::
	Suspend
	if A_IsSuspended 
	{  
		Notify("Suspended", "", 3)
    } else
    {
		Notify("Resumed", "", 3)
    }
Return

^#T::FixYTURL()	; CTRL WIN T
#p::LaunchPaint()
#Escape::LaunchPaint()
#s::LaunchSnippingTool()
; tweaked 7/10/16
;#s::Run SnippingTool.exe
#n::Run "C:\Program Files\Notepad++\notepad++.exe"
#k::Run, explore "K:\projects"
^Escape::
WinMove,A,, 0, 0
return

; paste clipboard
^!v::
SetKeyDelay, 15
if ( (StrLen(Clipboard) > 0) )
{
	Send {Text} %clipboard%
	clipboard = 
}
SetKeyDelay, 0
return


#Down::
    MouseGetPos,,,KDE_id
    ; This message is mostly equivalent to WinMinimize,
    ; but it avoids a bug with PSPad.
    PostMessage,0x112,0xf020,,,ahk_id %KDE_id%
    DoubleAlt := false
    return
; 1/25/16 -- how often did you use the old function from M$?

!LButton::
If DoubleAlt
{
    MouseGetPos,,,KDE_id
    ; This message is mostly equivalent to WinMinimize,
    ; but it avoids a bug with PSPad.
    PostMessage,0x112,0xf020,,,ahk_id %KDE_id%
    DoubleAlt := false
    return
}
; Get the initial mouse position and window id, and
; abort if the window is maximized.
MouseGetPos,KDE_X1,KDE_Y1,KDE_id
WinGet,KDE_Win,MinMax,ahk_id %KDE_id%
If KDE_Win
    return
; Get the initial window position.
WinGetPos,KDE_WinX1,KDE_WinY1,,,ahk_id %KDE_id%
Loop
{
    GetKeyState,KDE_Button,LButton,P ; Break if button has been released.
    If KDE_Button = U
        break
    MouseGetPos,KDE_X2,KDE_Y2 ; Get the current mouse position.
    KDE_X2 -= KDE_X1 ; Obtain an offset from the initial mouse position.
    KDE_Y2 -= KDE_Y1
    KDE_WinX2 := (KDE_WinX1 + KDE_X2) ; Apply this offset to the window position.
    KDE_WinY2 := (KDE_WinY1 + KDE_Y2)
    WinMove,ahk_id %KDE_id%,,%KDE_WinX2%,%KDE_WinY2% ; Move the window to the new position.
}
return
!RButton::
If DoubleAlt
{
    MouseGetPos,,,KDE_id
    ; Toggle between maximized and restored state.
    WinGet,KDE_Win,MinMax,ahk_id %KDE_id%
    If KDE_Win
        WinRestore,ahk_id %KDE_id%
    Else
        WinMaximize,ahk_id %KDE_id%
    DoubleAlt := false
    return
}
; Get the initial mouse position and window id, and
; abort if the window is maximized.
MouseGetPos,KDE_X1,KDE_Y1,KDE_id
WinGet,KDE_Win,MinMax,ahk_id %KDE_id%
If KDE_Win
    return
; Get the initial window position and size.
WinGetPos,KDE_WinX1,KDE_WinY1,KDE_WinW,KDE_WinH,ahk_id %KDE_id%
; Define the window region the mouse is currently in.
; The four regions are Up and Left, Up and Right, Down and Left, Down and Right.
If (KDE_X1 < KDE_WinX1 + KDE_WinW / 2)
   KDE_WinLeft := 1
Else
   KDE_WinLeft := -1
If (KDE_Y1 < KDE_WinY1 + KDE_WinH / 2)
   KDE_WinUp := 1
Else
   KDE_WinUp := -1
Loop
{
    GetKeyState,KDE_Button,RButton,P ; Break if button has been released.
    If KDE_Button = U
        break
    MouseGetPos,KDE_X2,KDE_Y2 ; Get the current mouse position.
    ; Get the current window position and size.
    WinGetPos,KDE_WinX1,KDE_WinY1,KDE_WinW,KDE_WinH,ahk_id %KDE_id%
    KDE_X2 -= KDE_X1 ; Obtain an offset from the initial mouse position.
    KDE_Y2 -= KDE_Y1
    ; Then, act according to the defined region.
    WinMove,ahk_id %KDE_id%,, KDE_WinX1 + (KDE_WinLeft+1)/2*KDE_X2  ; X of resized window
                            , KDE_WinY1 +   (KDE_WinUp+1)/2*KDE_Y2  ; Y of resized window
                            , KDE_WinW  -     KDE_WinLeft  *KDE_X2  ; W of resized window
                            , KDE_WinH  -       KDE_WinUp  *KDE_Y2  ; H of resized window
    KDE_X1 := (KDE_X2 + KDE_X1) ; Reset the initial position for the next iteration.
    KDE_Y1 := (KDE_Y2 + KDE_Y1)
}
return
; "Alt + MButton" may be simpler, but I
; like an extra measure of security for
; an operation like this.
!MButton::
If DoubleAlt
{
    MouseGetPos,,,KDE_id
    WinClose,ahk_id %KDE_id%
    DoubleAlt := false
    return
}
return
; This detects "double-clicks" of the alt key.
~Alt::
DoubleAlt := A_PriorHotKey = "~Alt" AND A_TimeSincePriorHotkey < 400
Sleep 0
KeyWait Alt  ; This prevents the keyboard's auto-repeat feature from interfering.
return


; der Hero Herr Gwarble : http://www.gwarble.com/ahk/Notify/
~CapsLock::
 {
   if getkeystate("capslock","t")
    {
      Notify("CAPS", "CAPS ON", 3)
    }
   else
    {
      Notify("CAPS", "caps off", 3)
    }
 }
return

;; preparing for Rodfell's throwing script
; This detects "double-clicks" of the CTRL key.
~Ctrl::
DoubleCtrl := A_PriorHotKey = "~Ctrl" AND A_TimeSincePriorHotkey < 400
Sleep 0
KeyWait Ctrl  ; This prevents the keyboard's auto-repeat feature from interfering.
return
;;;;;;;;;;;;;;; from Rodfell on AHK forum
;SysGet, Mon1, Monitor, 1
;SysGet, Mon2, Monitor, 2
;MsgBox, Left: %Mon2Left% -- Top: %Mon2Top% -- Right: %Mon2Right% -- Bottom %Mon2Bottom%.

; ~ prefix added by Houli Wang
~^LButton::
KeyWait, Lbutton
If DoubleCtrl
{
 DoubleCtrl := false
 ; verified that this is called as expected..
 mousegetpos,,,windowtomove
 gosub windowmove
 return
}
;Send {Ctrl Down}{Click down} ; thanks to Trik's response to Alan Stancliff
return
windowmove:
;MsgBox, Left: %Mon2Left% -- Top: %Mon2Top% -- Right: %Mon2Right% -- Bottom %Mon2Bottom%.
if ("" == Mon2Left )
{
 ;MsgBox No Way Jose
 return
}

wingetpos,x1,y1,w1,h1,ahk_id %windowtomove%
winget,winstate,minmax,ahk_id %windowtomove%
m1:=(x1+w1/2>mon1left) and (x1+w1/2<mon1right) and (y1+h1/2>mon1top) and (y1+h1/2<mon1bottom) ? 1:2   ;works out if centre of window is on monitor 1 (m1=1) or monitor 2 (m1=2)
m2:=m1=1 ? 2:1  ;m2 is the monitor the window will be moved to
ratiox:=abs(mon%m1%right-mon%m1%left)-w1<5 ? 0:abs((x1-mon%m1%left)/(abs(mon%m1%right-mon%m1%left)-w1))  ;where the window fits on x axis
ratioy:=abs(mon%m1%bottom-mon%m1%top)-h1<5 ? 0:abs((y1-mon%m1%top)/(abs(mon%m1%bottom-mon%m1%top)-h1))   ;where the window fits on y axis
x2:=mon%m2%left+ratiox*(abs(mon%m2%right-mon%m2%left)-w1)   ;where the window will fit on x axis in normal situation
y2:=mon%m2%top+ratioy*(abs(mon%m2%bottom-mon%m2%top)-h1)
w2:=w1  
h2:=h1   ;width and height will stay the same when moving unless reason not to lower in script
if abs(mon%m1%right-mon%m1%left)-w1<5 or abs(mon%m2%right-mon%m2%left-w1)<5   ;if x axis takes up whole axis OR won't fit on new screen
   {
   x2:=mon%m2%left  
   w2:=abs(mon%m2%right-mon%m2%left)
   }
if abs(mon%m1%bottom-mon%m1%top)-h1<5 or abs(mon%m2%bottom-mon%m2%top)-h1<5
   {
   y2:=mon%m2%top
   h2:=abs(mon%m2%bottom-mon%m2%top)
   }
if winstate   ;move maximized window
   {
   winrestore,ahk_id %windowtomove%
   winmove,ahk_id %windowtomove%,,mon%m2%left,mon%m2%top
   winmaximize,ahk_id %windowtomove%
   }
else
   {
   if (x1<mon%m1%left)
      x2:=mon%m2%left   ;adjustments for windows that are not fully on the initial monitor (m1)
   if (x1+w1>mon%m1%right)
      x2:=mon%m2%right-w2
   if (y1<mon%m1%top)
      y2:=mon%m2%top
   if (y1+h1>mon%m1%bottom)
      y2:=mon%m2%bottom-h2
   winmove,ahk_id %windowtomove%,,x2,y2,w2,h2   ;move non-maximized window
   winmaximize, ahk_id %windowtomove%
   }
return
;;;;;;;;;;;;;;; from Rodfell on AHK forum
#IfWinActive ahk_class XLMAIN
   ^!WheelUp::Send ^{PgUp}
   ^!WheelDown::Send ^{PgDn}
+WheelDown::SendInput {ScrollLock}{Right}{ScrollLock}
+WheelUp::SendInput {ScrollLock}{Left}{ScrollLock}
^BackSpace::Send ^+{Left}{BackSpace}	; to get delete word backward in Excel
#IfWinActive

#IfWinActive ahk_class Chrome_WidgetWin_1
   ^!WheelUp::Send ^{PgUp}
   ^!WheelDown::Send ^{PgDn}
#IfWinActive
#IfWinActive ahk_class ApplicationFrameWindow
   ^!WheelUp::Send ^{PgUp}
   ^!WheelDown::Send ^{PgDn}
#IfWinActive

#IfWinActive ahk_exe POWERPNT.EXE
 ^;::
  FormatTime , date, , MM/dd/yy
  SendInput %date%
 Return

; trying Areeb's solution 
^+B::
Try
	{
	ppt := ComObjActive("PowerPoint.Application")
	If (ppt.ActiveWindow.Selection.Type = 2)
		{
		Try
			ppt.ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Font.Color.RGB:=0xFF0000
		}
	
	If (ppt.ActiveWindow.Selection.Type = 3)
	ppt.ActiveWindow.Selection.TextRange.Font.Color.RGB:=0xFF0000
	}
Return

^+C::
Try
	{
	ppt := ComObjActive("PowerPoint.Application")
	If (ppt.ActiveWindow.Selection.Type = 2)
		{
		Try
			ppt.ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Font.Name:= "consolas"
		}
	
	If (ppt.ActiveWindow.Selection.Type = 3)
	ppt.ActiveWindow.Selection.TextRange.Font.Name:= "consolas"
	}

Return

#IfWinActive

;#Include C:\Users\ac104265\Documents\Autohotkey\Notify.ahk
#Include <Notify>
$!Escape::
 SetTitleMatchMode, 2
 IfWinActive, Citrix XenApp
 {
  Send +{Escape}
 } else {
  Winset, Bottom, , A
 }
 SetTitleMatchMode, 1
    return

;#IfWinActive ahk_class ConsoleWindowClass
#IfWinActive ahk_class mintty
   ^+L::Send clear{enter}
#IfWinActive

#IfWinActive, Untitled - Paint
   #F4::
WinGet, paintPID, PID, A
Process, Close, %paintPID%
   Return
#IfWinActive

; Inspired by using Alt-F1 on KDE to toggle maximization :)
$!F1::
 SetTitleMatchMode, 2
 IfWinActive, Citrix XenApp
 {
  Send !{F1}
 } else {
  MouseGetPos,,,KDE_id
  ; Toggle between maximized and restored state.
  WinGet,KDE_Win,MinMax,ahk_id %KDE_id%
  If KDE_Win
   WinRestore,ahk_id %KDE_id%,,ATL-WIN7 ; ExcludeTitle : "ATL-WIN7"
  Else
   WinMaximize,ahk_id %KDE_id%
  DoubleAlt := false
 }
 SetTitleMatchMode, 1
 return

; from Michael Nelson

; Trigger on alt scroll up (!WheelUp), it will not trigger itself ($) and AHK will not cause the key to be supressed (~)
~$!WheelUp::
if WinActive("ahk_class Notepad++") { ; if Notepad++
    loop % notepadScrollAmount-1  ; Loop X many more times to meet the notepadScrollAmount desired
        send {WheelUp} ; Sends a wheelup command to windows
}
return

; Same as wheelup but for wheeldown
~$!WheelDown::
if WinActive("ahk_class Notepad++") {
    loop % notepadScrollAmount-1
        send {WheelDown}
}
return

#Include <AccessibleObject>

; from Nour Nasser, UPWK
#If WinActive("AHK_Class MSPaintApp")

^[::
	SetKeyDelay -1, -1
	MSPaintFontSizeAdd(-1)
Return

^]::
	SetKeyDelay -1, -1
	MSPaintFontSizeAdd(1)
Return

#If


MSPaintFontSizeAdd(Value) {
	Static FontSize_Edit_Handle := 0
	If (!WinExist("AHK_ID " . FontSize_Edit_Handle))
	{
		Win := AccessibleObject.FromWindow(WinActive("AHK_Class MSPaintApp"))
		If (!Win)
			Return
		FontSize_Edit := ""
		For I, Control In Win.GetDescendants()
		{
			If (Control.Role = 0x2A && Control.Name = "Font Size")
			{
				FontSize_Edit := Control
				Break
			}
		}
		If (!FontSize_Edit)
			Return

		FontSize_Edit_Handle := FontSize_Edit.Handle
	}

	ControlFocus,, AHK_ID %FontSize_Edit_Handle%
	ControlGetText FontSize,, AHK_ID %FontSize_Edit_Handle%
	FontSize += Value
	FontSize := Max(1, FontSize)
	ControlSetText,, %FontSize%, AHK_ID %FontSize_Edit_Handle%
	ControlSend,, {Enter}, AHK_ID %FontSize_Edit_Handle%
}
