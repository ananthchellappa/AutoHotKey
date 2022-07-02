Class ComVar
{
	__New(Type := 0xC) {
		This._InternalArray := ComObjArray(Type, 1)
		DllCall("OleAut32\SafeArrayAccessData", "Ptr", ComObjValue(This._InternalArray), "Ptr*", pInternalArrayData)
		This._Pointer := ComObject(0x4000 | Type, pInternalArrayData)
	}

	Pointer
	{
		Get
		{
			Return This._Pointer
		}
	}

	Value
	{
		Get
		{
			Return This._InternalArray[0]
		}
		Set
		{
			Return This._InternalArray[0] := Value
		}
	}

	__Delete() {
		DllCall("OleAut32\SafeArrayUnaccessData", "Ptr", ComObjValue(This._InternalArray))
	}
}