Option Strict Off
Option Explicit On
Module modNickSpoof
	'modNickSpoof.bas
	'
	'Nick Spoofing module for Starcraft Brood War 1.08
	'First Created: 07-30-00
	'Last Modified: 05-18-01
	'by: Yoshi
	'http://www.scbackstab.com
	
	'dimention retained nickname globally
	Public RetainedNick As String
	'dimention the position of txtName
	Public NamePos As Short
	'for retrieving the game process handle
	Declare Function FindWindow Lib "user32"  Alias "FindWindowA"(ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
	'get processID from hwnd returned above
	Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer
	'return handle to target process
	Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
	'for writing game memory
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup1016.htm
	Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Integer, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer
	'for reading game memory
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup1016.htm
	Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Integer, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer
	'close each open handle
	Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
	'for determining OS version
	'UPGRADE_WARNING: Structure OSVERSIONINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup1050.htm
	Declare Function GetVersionExA Lib "kernel32" (ByRef lpVersionInformation As OSVERSIONINFO) As Short
	'User defined type for OSVERSIONINFO
	Public Structure OSVERSIONINFO
		Dim dwOSVersionInfoSize As Integer
		Dim dwMajorVersion As Integer
		Dim dwMinorVersion As Integer
		Dim dwBuildNumber As Integer
		Dim dwPlatformId As Integer
		'2 = Windows NT
		<VBFixedString(128),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=128)> Dim szCSDVersion As String
	End Structure
	'Necessary flags for NT
	Public Const STANDARD_RIGHTS_REQUIRED As Integer = &HF0000
	Public Const SYNCHRONIZE As Integer = &H100000
	Public Const PROCESS_ALL_ACCESS As Boolean = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFFs)
	
	
	Function ReadNick() As String
		'Dimention variables
		'Brood War Window Handle
		Dim bwWnd As Integer
		'Brood War Process ID
		Dim bwPid As Integer
		'Brood War Process Handle
		Dim bwPhand As Integer
		'Length returned
		Dim LengthReturned As Integer
		'String for containing nickname
		Dim strNick As VB6.FixedLengthString = New VB6.FixedLengthString(21)
		'end dimentioning
		
		'Get Brood War processid
		bwWnd = FindWindow(Nothing, "Brood War")
		'make sure Brood War is open
		If (bwWnd = 0) Then
			ReadNick = "Error 01: Brood War is not open."
			Exit Function
		End If
		'Convert window handle to process id
		GetWindowThreadProcessId(bwWnd, bwPid)
		
		'Get process handle
		bwPhand = OpenProcess(PROCESS_ALL_ACCESS, False, bwPid)
		If (bwPhand = 0) Then
			ReadNick = "Error 02: Can't get process handle."
			Exit Function
		End If
		
		'Read memory at the address of the name stored:
		' Version 1.07: 0x1902b0a0
		' Version 1.08O: 0x1902E4D8
		' Version 1.08 Final: 0x190314e8
		
		ReadProcessMemory(bwPhand, &H190314E8, strNick.Value, 21, LengthReturned)
		If (LengthReturned = 0) Then
			ReadNick = "Error 03: You must connect to Battle.net."
			CloseHandle(bwPhand)
			Exit Function
		End If
		'Close the handle we opened.
		CloseHandle(bwPhand)
		'And finally, return the nickname read.
		ReadNick = strNick.Value
	End Function
	Function ChangeNick(ByRef NewNick As String) As String
		Dim parse_Renamed As Object
		'Dimention variables
		'Brood War Window Handle
		Dim bwWnd As Integer
		'Brood War Process ID
		Dim bwPid As Integer
		'Brood War Process Handle
		Dim bwPhand As Integer
		'Length returned
		Dim LengthReturned As Integer
		'String used to write to memory
		Dim strNewName As VB6.FixedLengthString = New VB6.FixedLengthString(21)
		'end dimentioning
		'Get Textbox data
		strNewName.Value = NewNick
		For parse_Renamed = 1 To 21
			'Convert any trailing spaces in the string to an ascii 00 value.
			'UPGRADE_WARNING: Couldn't resolve default property of object parse_Renamed. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup1037.htm
			If (parse_Renamed > Len(NewNick)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object parse_Renamed. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup1037.htm
				If Mid(strNewName.Value, parse_Renamed, 1) = " " Then
					'UPGRADE_WARNING: Couldn't resolve default property of object parse_Renamed. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup1037.htm
					Mid(strNewName.Value, parse_Renamed, 1) = Chr(0)
				End If
			End If
		Next parse_Renamed
		'Get Brood War processid
		bwWnd = FindWindow(Nothing, "Brood War")
		'make sure Brood War is open
		If (bwWnd = 0) Then
			ChangeNick = "Error 01: Brood War window not found. Brood War must be open to spoof your nickname."
			Exit Function
		End If
		'Convert window handle to process id
		GetWindowThreadProcessId(bwWnd, bwPid)
		
		'Get process handle
		bwPhand = OpenProcess(PROCESS_ALL_ACCESS, False, bwPid)
		If (bwPhand = 0) Then
			ChangeNick = "Error 02: Unable to retreive Brood War process handle."
			Exit Function
		End If
		
		'Write textbox to memory location:
		' Version 1.07: 0x1902b0a0
		' Version 1.08O: 0x1902E4D8
		' Version 1.08 Final: 0x190314e8
		' Version 1.08 Final Channel: 0x19030094
		WriteProcessMemory(bwPhand, &H190314E8, strNewName.Value, 21, LengthReturned)
		If (LengthReturned = 0) Then
			ChangeNick = "Error 04: No data entered into Brood War. You must be connected to Battle.net to spoof your nickname."
			CloseHandle(bwPhand)
			Exit Function
		End If
		'And finally, close the handle we opened.
		CloseHandle(bwPhand)
	End Function
	Function IsAcceptable(ByRef Ascii As Short) As Boolean
		'Check if Ascii is acceptable in battle.net nicknames
		If Ascii = 8 Then IsAcceptable = True : Exit Function
		If Ascii = 33 Then IsAcceptable = True : Exit Function
		If Ascii = 36 Then IsAcceptable = True : Exit Function
		If Ascii > 38 And Ascii < 42 Then IsAcceptable = True : Exit Function
		If Ascii = 43 Then IsAcceptable = True : Exit Function
		If Ascii > 44 And Ascii < 47 Then IsAcceptable = True : Exit Function
		If Ascii > 47 And Ascii < 60 Then IsAcceptable = True : Exit Function
		If Ascii = 61 Then IsAcceptable = True : Exit Function
		If Ascii = 64 Then IsAcceptable = True : Exit Function
		If Ascii > 64 And Ascii < 92 Then IsAcceptable = True : Exit Function
		If Ascii > 92 And Ascii < 127 Then IsAcceptable = True : Exit Function
		IsAcceptable = False
	End Function
	Function EncodeToASCII(ByRef BWColor As String) As Short
		If BWColor = "Blue" Then EncodeToASCII = 14
		If BWColor = "Green" Then EncodeToASCII = 15
		If BWColor = "Highlighted Green" Then EncodeToASCII = 16
		If BWColor = "Black" Then EncodeToASCII = 17
		If BWColor = "White" Then EncodeToASCII = 18
		If BWColor = "Red" Then EncodeToASCII = 19
		If BWColor = "Invisible" Then EncodeToASCII = 20
	End Function
	Function DecodeToBWCompatible(ByRef EncodedASCII As Short) As Short
		If EncodedASCII = 14 Then DecodeToBWCompatible = 2 'blue
		If EncodedASCII = 15 Then DecodeToBWCompatible = 3 'green
		If EncodedASCII = 16 Then DecodeToBWCompatible = 4 'highlighted green
		If EncodedASCII = 17 Then DecodeToBWCompatible = 5 'black
		If EncodedASCII = 18 Then DecodeToBWCompatible = 6 'white
		If EncodedASCII = 19 Then DecodeToBWCompatible = 7 'red
		If EncodedASCII = 20 Then DecodeToBWCompatible = 8 'invisible
	End Function
	Function getVersion() As Integer
		Dim osinfo As OSVERSIONINFO
		Dim retvalue As Short
		osinfo.dwOSVersionInfoSize = 148
		osinfo.szCSDVersion = Space(128)
		retvalue = GetVersionExA(osinfo)
		getVersion = osinfo.dwPlatformId
	End Function
End Module