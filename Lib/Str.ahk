
StrJoin(Separator, String1:="", Strings*) {
	JoinedString := String1
	For I, String In Strings
	{
		JoinedString .= Separator . String
	}
	Return JoinedString
}

StrStartsWith(String, SearchString, CaseSensitive := False) {
	If (CaseSensitive)
		Return SubStr(String, 1, StrLen(SearchString)) == SearchString
	Else
		Return SubStr(String, 1, StrLen(SearchString)) = SearchString
}

StrEndsWith(String, SearchString, CaseSensitive := False) {
	If (CaseSensitive)
		Return SubStr(String, -StrLen(SearchString) + 1) == SearchString
	Else
		Return SubStr(String, -StrLen(SearchString) + 1) = SearchString
}

StrGetLine(String, LineNumber, EOL:="`n") {
	LineStartPos := InStr(String, EOL,,, LineNumber - 1)
	If (!LineStartPos)
		Return ""
	LineStartPos += StrLen(EOL)
	LineEndPos := InStr(String, EOL,, LineStartPos)
	If (LineEndPos)
		Return SubStr(String, LineStartPos, LineEndPos - LineStartPos)
	Else
		Return SubStr(String, LineStartPos)
}

StrReplaceLine(String, LineNumber, ReplaceLine, EOL:="`n") {
	LineStartPos := InStr(String, EOL,,, LineNumber - 1)
	If (!LineStartPos)
		Return String
	LineStartPos += StrLen(EOL)
	LineEndPos := InStr(String, EOL,, LineStartPos)
	If (LineEndPos)
		Return SubStr(String, 1, LineStartPos - 1) . ReplaceLine . SubStr(String, LineEndPos)
	Else
		Return SubStr(String, 1, LineStartPos - 1) . ReplaceLine
}

StrRepeat(String, RepeatCount) {
	FinalStr := ""
	Loop %RepeatCount%
		FinalStr .= String
	Return FinalStr
}

LineFromPos(String, Pos, EOL:="`n") {
	StrLen := StrLen(String)
	If (Pos > StrLen)
		Return 0
	LineNumber := 1
	While (Pos := InStr(String, EOL,, Pos - StrLen - 1))
		LineNumber++
	Return LineNumber
}

ColumnFromPos(String, Pos, EOL:="`n") {
	LinePos := InStr(String, EOL,, Pos - StrLen(String) - 1)
	Return Pos - LinePos
}

StrPutEx(String, ByRef Target, Encoding) {
	Length := StrPut(String, Encoding)
	cbLength := Length * (Encoding = "UTF-16" || Encoding = "CP1200" ? 2 : 1)
	VarSetCapacity(Target, cbLength)
	StrPut(String, &Target, Encoding)
	Return cbLength
}

UTF8_GetBytes(String) {
	BytesCount := StrPutEx(String, Buff, "UTF-8") - 1 ; ignore the null terminator
	Bytes := []
	Loop %BytesCount%
		Bytes.Push(NumGet(Buff, A_Index - 1, "UChar"))
	Return Bytes
}

ANSI_GetBytes(String) {
	BytesCount := StrPutEx(String, Buff, "CP0") - 1 ; ignore the null terminator
	Bytes := []
	Loop %BytesCount%
		Bytes.Push(NumGet(Buff, A_Index - 1, "UChar"))
	Return Bytes
}

ASCII_GetBytes(String) {
	BytesCount := StrPutEx(String, Buff, "CP1252") - 1 ; ignore the null terminator
	Bytes := []
	Loop %BytesCount%
		Bytes.Push(NumGet(Buff, A_Index - 1, "UChar"))
	Return Bytes
}

Unicode_GetBytes(String) {
	BytesCount := StrPutEx(String, Buff, "CP1200") - 2 ; ignore the null terminator
	Bytes := []
	Loop %BytesCount%
		Bytes.Push(NumGet(Buff, A_Index - 1, "UChar"))
	Return Bytes
}