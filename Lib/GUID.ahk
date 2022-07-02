CreateUUID() {
	VarSetCapacity(UUID, 16, 0)
	DllCall("Rpcrt4\UuidCreate", "Ptr", &UUID)
	Return GUIDToStr(&UUID)
}

GUIDFromStr(sGUID, ByRef GUID) {
	VarSetCapacity(GUID, 16)
	Return DllCall("Ole32\CLSIDFromString", "Str", sGUID, "Ptr", &GUID, "Ptr")
}

GUIDToStr(pGUID) {
	pString := 0
	DllCall("Ole32\StringFromCLSID", "Ptr", pGUID, "Ptr*", pString, "Ptr")
	sGUID := StrGet(pString)
	DllCall("Ole32\CoTaskMemFree", "Ptr", pString)
	Return sGUID
}

GUIDIsEqual(pGUID1, pGUID2) {
	If (pGUID1 = pGUID2)
		Return True
	Loop 2
	{
		If (NumGet(pGUID1 + (A_Index - 1) * 8, "Int64") != NumGet(pGUID2 + (A_Index - 1) * 8, "Int64"))
			Return False
	}
	Return True
}

Class GUID
{
	__New(sGUID) {
		This._String := sGUID
		This.SetCapacity("_GUID", 16)
		This._Pointer := This.GetAddress("_GUID")
		DllCall("Ole32\CLSIDFromString", "Str", sGUID, "Ptr", This._Pointer, "Ptr")
	}

	String {
		Get {
			Return This._String
		}
	}

	Pointer {
		Get {
			Return This._Pointer
		}
	}
}


