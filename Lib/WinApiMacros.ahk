; some win api macros
Succeeded(hr) {
	Return hr >= 0
}

Failed(hr) {
	Return hr < 0
}

HiWord(dwValue) {
	Return (dwValue >> 16) & 0xFFFF
}

LoWord(dwValue) {
	Return dwValue & 0xFFFF
}

Get_X_LParam(lParam) {
	Return (lParam & 0xFFFF) << 48 >> 48
}

Get_Y_LParam(lParam) {
	Return ((lParam >> 16) & 0xFFFF) << 48 >> 48
}

MakeLong(wLow, wHigh) {
	Return (wLow & 0xFFFF) | ((wHigh & 0xFFFF) << 16)
}

MakeWord(wLow, wHigh) {
	Return (wLow & 0xFF) | ((wHigh & 0xFF) << 8)
}

RGB(byRed, byGreen, byBlue) {
	Return (byRed & 0xFF) | ((byGreen & 0xFF) << 8) | ((byBlue & 0xFF) << 16)
}

SetRect(lprc, xLeft, yTop, xRight, yBottom) {
	NumPut((xLeft & 0xFFFFFFFF) | (yTop << 32), lprc+0, "UInt64")
	NumPut((xRight & 0xFFFFFFFF) | (yBottom << 32), lprc+8, "UInt64")
}