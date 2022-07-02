#Include WinApiMacros.ahk
#Include GUID.ahk
#Include ComVar.ahk
#Include RemoteBuff.ahk
#Include Str.ahk

ControlGetName(hWnd) {
	Static WM_GETCONTROLNAME := DllCall("RegisterWindowMessage", "Str", "WM_GETCONTROLNAME", "UInt")
	Static MaxChars := 260
	Static BuffSize := MaxChars * 2

	Buff := RemoteBuff.Create(BuffSize, "AHK_ID " hWnd)
	If (Buff)
	{
		Success := DllCall("SendMessage", "Ptr", hWnd, "UInt", WM_GETCONTROLNAME, "Ptr", MaxChars, "Ptr", Buff.Pointer, "Ptr")
		If (Success)
		{
			VarSetCapacity(Text, BuffSize)
			Buff.Read(Text)
			Return StrGet(&Text)
		}
	}
}

ControlGetType(hWnd) {
	Static WM_GETCONTROLTYPE := DllCall("RegisterWindowMessage", "Str", "WM_GETCONTROLTYPE", "UInt")
	Static MaxChars := 512
	Static BuffSize := MaxChars * 2

	Buff := RemoteBuff.Create(BuffSize, "AHK_ID " hWnd)
	If (Buff)
	{
		Success := DllCall("SendMessage", "Ptr", hWnd, "UInt", WM_GETCONTROLTYPE, "Ptr", MaxChars, "Ptr", Buff.Pointer, "Ptr")
		If (Success)
		{
			VarSetCapacity(Text, BuffSize)
			Buff.Read(Text)
			Return StrGet(&Text)
		}
	}
}

Class AccessibleObject
{
	; Static hOle32 := DllCall("LoadLibrary", "Str", "Ole32")
	Static hOleacc := DllCall("LoadLibrary", "Str", "Oleacc")
	Static IID_IAccessible := New GUID("{618736E0-3C3D-11CF-810C-00AA00389B71}")


	Static IncludeHidden := False
	Static IncludeDisabled := False
	Static IncludeWindows := True


	FromWindow(hWnd, dwID:=0xFFFFFFFC) {
		pIAccessible := 0
		HResult := DllCall("Oleacc\AccessibleObjectFromWindow"
			, "Ptr", hWnd
			, "UInt", dwID
			, "Ptr", AccessibleObject.IID_IAccessible.Pointer
			, "Ptr*", pIAccessible)
		If (Failed(HResult) || !pIAccessible)
			Return

		IAccessible := ComObject(pIAccessible)
		ObjRelease(pIAccessible)

		Return New AccessibleObject(IAccessible)
	}

	FromPoint(X, Y) {
		Static SizeOf_VARIANT := A_PtrSize = 8 ? 24 : 16

		POINT := (X & 0xFFFFFFFF) | ((Y & 0xFFFFFFFF) << 32)
		VarSetCapacity(VarChild, SizeOf_VARIANT, 0)

		pIAccessible := 0
		HResult := DllCall("Oleacc\AccessibleObjectFromPoint"
			, "Int64", POINT
			, "Ptr*", pIAccessible
			, "Ptr", &VarChild)
		If (Failed(HResult))
			Return
		DllCall("OleAut32\VariantClear", "Ptr", &VarChild)
		If (!pIAccessible)
			Return

		IAccessible := ComObject(pIAccessible)
		ObjRelease(pIAccessible)

		Return New AccessibleObject(IAccessible)
	}


	__New(IAccessible) {
		This.IAccessible := IAccessible
	}

	_TryGet(MemberName, ChildId:="") {
		Try
		{
			If (ChildId != "")
				Return This.IAccessible[MemberName, ChildId]
			Else
				Return This.IAccessible[MemberName]
		}
		Catch e
		{
			If (InStr(e.Message, "0x80004001") || InStr(e.Message, "0x80020003") || InStr(e.Message, "0x8007000E"))
				Return
			Throw e
		}
	}

	Handle
	{
		Get
		{
			If (!This.HasKey("_Handle"))
			{
				Handle := 0
				DllCall("Oleacc\WindowFromAccessibleObject", "Ptr", ComObjValue(This.IAccessible), "Ptr*", Handle)
				This._Handle := Handle
			}
			Return This._Handle
		}
	}

	Name
	{
		Get
		{
			Return This._TryGet("accName", 0)
		}
	}

	ControlName
	{
		Get
		{
			If (!This.HasKey("_ControlName"))
			{
				This._ControlName := ""
				If (This.Handle)
				{
					This._ControlName := ControlGetName(This.Handle)
				}
			}
			Return This._ControlName
		}
	}

	ControlType
	{
		Get
		{
			If (!This.HasKey("_ControlType"))
			{
				This._ControlType := ""
				If (This.Handle)
				{
					This._ControlType := ControlGetType(This.Handle)
				}
			}
			Return This._ControlType
		}
	}

	ClassName
	{
		Get
		{
			If (!This.HasKey("_ClassName"))
			{
				This._ClassName := ""
				If (This.Handle)
				{
					WinGetClass ClassName, % "AHK_ID " This.Handle
					This._ClassName := ClassName
				}
			}
			Return This._ClassName
		}
	}

	Description
	{
		Get
		{
			Return This._TryGet("accDescription", 0)
		}
	}

	FocusIndex
	{
		Get
		{
			Value := This.IAccessible.accFocus
			If (!IsObject(Value))
				Return Value
		}
	}

	Focus
	{
		Get
		{
			Value := This.IAccessible.accFocus
			If (IsObject(Value))
				Return New AccessibleObject(Value)
		}
	}

	Help
	{
		Get
		{
			Return This._TryGet("accHelp", 0)
		}
	}

	KeyboardShortcut
	{
		Get
		{
			Return This._TryGet("accKeyboardShortcut", 0)
		}
	}

	ChildCount
	{
		Get
		{
			Return This.IAccessible.accChildCount
		}
	}

	DefaultAction
	{
		Get
		{
			Return This._TryGet("accDefaultAction", 0)
		}
	}

	Role
	{
		Get
		{
			Return This.IAccessible.accRole(0)
		}
	}

	RoleText
	{
		Get
		{
			Static MaxChars := 150
			Static BuffSize := MaxChars * 2

			VarSetCapacity(TextBuff, BuffSize)
			DllCall("Oleacc\GetRoleText", "UInt", This.Role, "Str", TextBuff, "UInt", MaxChars)
			Return TextBuff
		}
	}

	StateText
	{
		Get
		{
			Static MaxChars := 150
			Static BuffSize := MaxChars * 2

			VarSetCapacity(TextBuff, BuffSize)

			States := []
			State := This.State
			Loop 32
			{
				Flag := 0x1 << (A_Index - 1)
				If (State & Flag)
				{
					DllCall("Oleacc\GetStateText", "UInt", Flag, "Str", TextBuff, "UInt", MaxChars)
					States.Push(TextBuff)
				}
			}
			Return StrJoin(", ", States*)
		}
	}

	State
	{
		Get
		{
			Return This.IAccessible.accState(0)
		}
	}

	SelectedIndexes
	{
		Get
		{
			Selection := This._TryGet("accSelection")
			If (Selection = "")
				Return []
			
			If (!IsObject(Selection)) ; child id
			{
				Return [Selection]
			}
			Else
			{
				If (ComObjType(Selection) = 9) ; IDispatch
				{
					Return []
				}
				Else ; IEnumVARIANT
				{
					Items := []
					While (Selection.Next(Item))
					{
						If (!IsObject(Item))
						{
							Items.Push(Item)
						}
					}
					Return Items
				}
			}
		}
	}

	SelectedItems
	{
		Get
		{
			Selection := This._TryGet("accSelection")
			If (Selection = "")
				Return []
			
			If (IsObject(Selection))
			{
				If (ComObjType(Selection) = 9) ; IDispatch
				{
					Return [New AccessibleObject(Selection)]
				}
				Else ; IEnumVARIANT
				{
					Items := []
					While (Selection.Next(Item))
					{
						If (IsObject(Item))
						{
							Items.Push(New AccessibleObject(Item))
						}
					}
					Return Items
				}
			}
			Else
			{
				Return []
			}
		}
	}

	Parent
	{
		Get
		{
			Parent := This.IAccessible.accParent
			If (IsObject(Parent))
			{
				Parent := New AccessibleObject(Parent)
				If (!AccessibleObject.IncludeWindows && Parent.Role = 9 && Parent.ChildCount = 7)
				{
					Parent := Parent.Parent
				}
				Return Parent
			}
			Return ""
		}
	}

	Value
	{
		Get
		{
			Return This._TryGet("accValue", 0)
		}
		Set
		{
			Return This.IAccessible.accValue(0) := Value
		}
	}

	DoDefaultAction(ChildId:=0) {
		This.IAccessible.accDoDefaultAction(ChildId)
	}

	HitTest(X, Y) {
		Return This.IAccessible.accHitTest(X, Y)
	}

	Location(ByRef X, ByRef Y, ByRef Width, ByRef Height) {
		oX := New ComVar(0x3)
		oY := New ComVar(0x3)
		oWidth := New ComVar(0x3)
		oHeight := New ComVar(0x3)
		This.IAccessible.accLocation(oX.Pointer, oY.Pointer, oWidth.Pointer, oHeight.Pointer, 0)
		X := oX.Value
		Y := oY.Value
		Width := oWidth.Value
		Height := oHeight.Value
	}

	GetChildren() {
		Static SizeOf_VARIANT := A_PtrSize = 8 ? 24 : 16

		Children := []
		ChildCount := This.IAccessible.accChildCount
		If (ChildCount = 0)
			Return Children, ErrorLevel := 0

		VarSetCapacity(VarArray, SizeOf_VARIANT * ChildCount * 2, 0)
		pIAccessible := ComObjQuery(This.IAccessible, AccessibleObject.IID_IAccessible.String)
		If (!pIAccessible)
			Return Children, ErrorLevel := 1

		Result := DllCall("Oleacc\AccessibleChildren", "Ptr", pIAccessible, "Int", 0, "Int", ChildCount, "Ptr", &VarArray, "Int*", Obtained, "Int")
		ObjRelease(pIAccessible)
		If ((Result != 0 && Result != 1) || Obtained = 0)
			Return Children, ErrorLevel := 2

		Loop %Obtained%
		{
			pVar := &VarArray + (A_Index - 1) * SizeOf_VARIANT
			Type := NumGet(pVar + 0, "Int")
			If (Type = 0x9)
			{
				pChild_IDispatch := NumGet(pVar + 8, "Ptr")

				Child_IDispatch := ComObject(pChild_IDispatch)
				ObjRelease(pChild_IDispatch)

				Child := New AccessibleObject(Child_IDispatch)
				If ((!AccessibleObject.IncludeHidden && Child.IsHidden())
					|| (!AccessibleObject.IncludeDisabled && Child.IsDisabled()))
				{
					Continue
				}

				If (!AccessibleObject.IncludeWindows && Child.Role = 9 && Child.ChildCount = 7)
				{
					Child := New AccessibleObject(Child.IAccessible.accChild(-4))
				}
				Children.Push(Child)
			}
		}
		For I, Child In Children
			Child.Index := I
		Return Children, ErrorLevel := 0
	}

	GetChildrenIds() {
		Static SizeOf_VARIANT := A_PtrSize = 8 ? 24 : 16

		Children := []
		ChildCount := This.IAccessible.accChildCount
		If (ChildCount = 0)
			Return Children, ErrorLevel := 0

		VarSetCapacity(VarArray, SizeOf_VARIANT * ChildCount * 2, 0)
		pIAccessible := ComObjQuery(This.IAccessible, AccessibleObject.IID_IAccessible.String)
		If (!pIAccessible)
			Return Children, ErrorLevel := 1

		Result := DllCall("Oleacc\AccessibleChildren", "Ptr", pIAccessible, "Int", 0, "Int", ChildCount, "Ptr", &VarArray, "Int*", Obtained, "Int")
		ObjRelease(pIAccessible)
		If ((Result != 0 && Result != 1) || Obtained = 0)
			Return Children, ErrorLevel := 2

		Loop %Obtained%
		{
			pVar := &VarArray + (A_Index - 1) * SizeOf_VARIANT
			Type := NumGet(pVar + 0, "Int")

			If (Type = 0x3)
			{
				ChildId := NumGet(pVar + 8, "Int")
				Children.Push(ChildId)
			}
		}
		Return Children, ErrorLevel := 0
	}

	Select(ChildId:=0) {
		This.IAccessible.accSelect(0x2, ChildId)
	}

	AddSelection(ChildId) {
		This.IAccessible.accSelect(0x8, ChildId)
	}

	SetFocus(ChildId:=0) {
		This.IAccessible.accSelect(0x1, ChildId)
	}

	IsDisabled() {
		Return (This.State & 0x1)
	}

	IsHidden() {
		Return (This.State & 0x8000)
	}

	GetDescendants() {
		Return New AccessibleObject.Descendants(This)
	}

	GetDescendantByName(Name) {
		For I, Child In This.GetDescendants()
		{
			If (Child.Name = Name)
				Return Child
		}
	}

	GetDescendantByControlName(ControlName) {
		For I, Child In This.GetDescendants()
		{
			If (Child.ControlName = ControlName)
				Return Child
		}
	}

	GetDescendantByRole(Role) {
		For I, Child In This.GetDescendants()
		{
			If (Child.Role = Role)
				Return Child
		}
	}

	Class Descendants
	{
		__New(oAcc) {
			This.oAcc := oAcc
		}

		_NewEnum() {
			Return New AccessibleObject.DescendantsEnumerator(This.oAcc)
		}
	}

	Class DescendantsEnumerator
	{
		__New(oAcc) {
			This.oAcc := oAcc
			This.Children := ""
			This.LastChild := ""
			This.Index := 0
		}

		Next(ByRef Index, ByRef Value) {
			If (!This.Children)
				This.Children := This.oAcc.GetChildren()
			If (This.LastChild)
			{
				Children := This.LastChild.GetChildren()
				If (Children.Length())
					This.Children.InsertAt(1, Children*)
			}
			Child := This.Children.RemoveAt(1)
			If (!Child)
				Return False
			This.LastChild := Child
			This.Index += 1
			Index := This.Index
			Value := Child
			Return True
		}
	}
}