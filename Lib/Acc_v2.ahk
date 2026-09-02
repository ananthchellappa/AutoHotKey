#Requires AutoHotkey v2.0
;
; Acc_v2.ahk -- minimal MSAA / IAccessible helper for AutoHotkey v2.
;
; Replaces the v1 Lib\AccessibleObject.ahk for the only thing this repo asks of
; it: walking a window's accessibility tree to find a control by Role + Name,
; then getting its HWND or invoking its default action (the MSPaint ribbon
; "Font Size" box and "Text" tool).
;
; It deliberately implements just that slice, so the v1 dependency chain
; (WinApiMacros / GUID / ComVar / RemoteBuff / Str) is not needed -- those
; existed mostly for ControlGetName's cross-process WM_GETCONTROLNAME buffer,
; which nothing here uses.
;
; API kept source-compatible with the v1 class where it is used:
;   AccessibleObject.FromWindow(hWnd)  -> object or 0
;   .Name  .Role  .State  .ChildCount  .Handle
;   .GetChildren()      -> Array
;   .GetDescendants()   -> lazy depth-first enumerator (single loop variable)
;   .DoDefaultAction(ChildId := 0)
;
; Like the v1 original, hidden and disabled children are skipped by default.
;

class AccessibleObject
{
    static IID_IAccessible := AccGuidFromString("{618736E0-3C3D-11CF-810C-00AA00389B71}")

    static IncludeHidden := false
    static IncludeDisabled := false

    static OBJID_CLIENT := 0xFFFFFFFC

    ; MSAA state bits used for filtering
    static STATE_UNAVAILABLE := 0x1
    static STATE_INVISIBLE := 0x8000

    IAccessible := 0
    Index := 0

    __New(IAccessible) {
        this.IAccessible := IAccessible
    }

    static FromWindow(hWnd, dwId := 0xFFFFFFFC) {
        if (!hWnd)
            return 0

        pAcc := 0
        hr := DllCall("Oleacc\AccessibleObjectFromWindow"
            , "Ptr", hWnd
            , "UInt", dwId
            , "Ptr", AccessibleObject.IID_IAccessible
            , "Ptr*", &pAcc
            , "Int")
        if (hr < 0 || !pAcc)
            return 0

        ; ComObjFromPtr takes ownership of the reference -- do not release it here.
        return AccessibleObject(ComObjFromPtr(pAcc))
    }

    Name {
        get {
            try
                return this.IAccessible.accName[0]
            catch
                return ""
        }
    }

    Role {
        get {
            try
                return this.IAccessible.accRole[0]
            catch
                return 0
        }
    }

    State {
        get {
            try
                return this.IAccessible.accState[0]
            catch
                return 0
        }
    }

    ChildCount {
        get {
            try
                return this.IAccessible.accChildCount
            catch
                return 0
        }
    }

    Handle {
        get {
            if (!this.HasOwnProp("_Handle")) {
                hWnd := 0
                try
                    DllCall("Oleacc\WindowFromAccessibleObject"
                        , "Ptr", ComObjValue(this.IAccessible)
                        , "Ptr*", &hWnd)
                this._Handle := hWnd
            }
            return this._Handle
        }
    }

    DoDefaultAction(ChildId := 0) {
        try
            this.IAccessible.accDoDefaultAction(ChildId)
    }

    GetChildren() {
        ; VARIANT is 16 bytes on x86, 24 on x64.
        static SizeOf_VARIANT := 8 + 2 * A_PtrSize

        children := []
        count := this.ChildCount
        if (count < 1)
            return children

        varArray := Buffer(SizeOf_VARIANT * count * 2, 0)
        obtained := 0
        result := DllCall("Oleacc\AccessibleChildren"
            , "Ptr", ComObjValue(this.IAccessible)
            , "Int", 0
            , "Int", count
            , "Ptr", varArray
            , "Int*", &obtained
            , "Int")
        if ((result != 0 && result != 1) || !obtained)
            return children

        loop obtained {
            offset := (A_Index - 1) * SizeOf_VARIANT
            ; Only VT_DISPATCH (0x9) entries are real objects; VT_I4 entries are
            ; "simple elements" that have no window handle of their own.
            if (NumGet(varArray, offset, "UShort") != 0x9)
                continue
            pChild := NumGet(varArray, offset + 8, "Ptr")
            if (!pChild)
                continue

            child := AccessibleObject(ComObjFromPtr(pChild))
            state := child.State
            if ((!AccessibleObject.IncludeHidden && (state & AccessibleObject.STATE_INVISIBLE))
                || (!AccessibleObject.IncludeDisabled && (state & AccessibleObject.STATE_UNAVAILABLE)))
                continue

            children.Push(child)
        }

        for i, child in children
            child.Index := i
        return children
    }

    ; Lazy depth-first walk, same order as the v1 DescendantsEnumerator.
    ; Use one loop variable:  for ctl in win.GetDescendants()
    GetDescendants() {
        stack := this.GetChildren()
        return NextDescendant

        NextDescendant(&value, *) {
            if (!stack.Length)
                return false
            item := stack.RemoveAt(1)
            kids := item.GetChildren()
            if (kids.Length)
                stack.InsertAt(1, kids*)
            value := item
            return true
        }
    }

    GetDescendantByRoleAndName(Role, Name) {
        for ctl in this.GetDescendants() {
            if (ctl.Role = Role && ctl.Name = Name)
                return ctl
        }
        return 0
    }
}

AccGuidFromString(sGuid) {
    buf := Buffer(16, 0)
    DllCall("Ole32\CLSIDFromString", "Str", sGuid, "Ptr", buf, "Int")
    return buf
}
