Class RemoteBuff
{
	Create(Size, WinTitle:="", WinText:="", ExcludeTitle:="", ExcludeText:="") {
		WinGet PID, PID, %WinTitle%, %WinText%, %ExcludeTitle%, %ExcludeText%
		If (PID = "")
			Return
		hProcess := DllCall("OpenProcess", "UInt", 0x0008 | 0x0010 | 0x0020, "Int", 0, "UInt", PID, "Ptr")
		If (hProcess)
		{
			pBuff := DllCall("VirtualAllocEx", "Ptr", hProcess, "Ptr", 0, "UInt", Size, "UInt", 0x1000, "UInt", 0x04, "Ptr")
			If (pBuff)
			{
				Return New RemoteBuff(hProcess, pBuff, Size, True)
			}

			DllCall("CloseHandle", "Ptr", hProcess)
		}
	}

	Get(Address, Size, WinTitle:="", WinText:="", ExcludeTitle:="", ExcludeText:="") {
		WinGet PID, PID, %WinTitle%, %WinText%, %ExcludeTitle%, %ExcludeText%
		If (PID = "")
			Return
		hProcess := DllCall("OpenProcess", "UInt", 0x0008 | 0x0010 | 0x0020, "Int", 0, "UInt", PID, "Ptr")
		If (hProcess)
		{
			Return New RemoteBuff(hProcess, Address, Size, False)
		}
	}

	__New(hProcess, pBuff, Size, OwnBuff) {
		This.hProcess := hProcess
		This.pBuff := pBuff
		This.Size := Size
		This.OwnBuff := OwnBuff
	}

	__Delete() {
		If (This.OwnBuff)
			DllCall("VirtualFreeEx", "Ptr", This.hProcess, "Ptr", This.pBuff, "UInt", 0, "UInt", 0x8000, "Int")
		DllCall("CloseHandle", "Ptr", This.hProcess)
	}

	Pointer
	{
		Get
		{
			Return This.pBuff
		}
	}

	Read(ByRef Buff, Offset:=0, Length:=-1) {
		If (Length = -1)
		{
			Length := Min(This.Size - Offset, VarSetCapacity(Buff))
		}
		Return DllCall("ReadProcessMemory", "Ptr", This.hProcess, "Ptr", This.pBuff + Offset
			, "Ptr", &Buff, "UInt", Length, "Ptr", 0, "Int")
	}

	Write(ByRef Buff, Offset:=0, Length:=-1) {
		If (Length = -1)
		{
			Length := Min(This.Size - Offset, VarSetCapacity(Buff))
		}
		Return DllCall("WriteProcessMemory", "Ptr", This.hProcess, "Ptr", This.pBuff + Offset
			, "Ptr", &Buff, "UInt", Length, "Ptr", 0, "Int")
	}

	ReadPtr(pBuff, Offset:=0, Length:=-1) {
		If (Length = -1)
		{
			Length := This.Size - Offset
		}
		Return DllCall("ReadProcessMemory", "Ptr", This.hProcess, "Ptr", This.pBuff + Offset
			, "Ptr", pBuff, "UInt", Length, "Ptr", 0, "Int")
	}

	WritePtr(pBuff, Offset:=0, Length:=-1) {
		If (Length = -1)
		{
			Length := This.Size - Offset
		}
		Return DllCall("WriteProcessMemory", "Ptr", This.hProcess, "Ptr", This.pBuff + Offset
			, "Ptr", pBuff, "UInt", Length, "Ptr", 0, "Int")
	}
}
