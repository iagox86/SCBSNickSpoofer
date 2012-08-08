Option Strict Off
Option Explicit Off
Namespace NickSpoofer
	Module modNickSpoof
		'modNickSpoof.bas
		' 
		'Nick Spoofing module for Starcraft Brood War 1.07
		'Last Modified: 07-30-00
		'by: Yoshi
		'http://www.scbackstab.com
		
		'dimention retained nickname globally
		Public  RetainedNick As String
		'dimention the position of txtName
		Public  NamePos As Short
		'for retrieving the game process handle
		Declare Function FindWindow Lib "user32"  Alias "FindWindowA"(ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
		'get processID from hwnd returned above
		Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Integer, ByRef lpdwProcessId As Integer) As Integer
		'return handle to target process
		Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
		'for writing game memory
		Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Integer, ByVal lpBaseAddress As Object, ByVal lpBuffer As Object, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer
		'for reading game memory
		Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Integer, ByVal lpBaseAddress As Object, ByVal lpBuffer As Object, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer
		'close each open handle
		Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
		Function ReadNick() As String
			Dim PROCESS_ALL_ACCESS As Object
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
			'UPGRADE_WARNING: A boolean is being converted to a numeric type. The numeric type of True has changed from -1 in VB6 to 1 in VB7.
			bwPhand = OpenProcess(PROCESS_ALL_ACCESS, False, bwPid)
			If (bwPhand = 0) Then
				ReadNick = "Error 02: Can't get process handle."
				Exit Function
			End If
			'Read memory at the address of the name stored:
			' Version 1.07: 0x1902b0a0
			' Version 1.08O: 0x1902E4D8
			
			ReadProcessMemory(bwPhand, &H1902B0A0, strNick.Value, 21, LengthReturned)
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
			Dim PROCESS_ALL_ACCESS As Object
			Dim parse As Object
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
			For parse = 1 To 21
				'Convert any trailing spaces in the string to an ascii 00 value.
				If (parse > Len(NewNick)) Then
					If Mid(strNewName.Value, parse, 1) = " " Then Mid(strNewName.Value, parse, 1) = Chr(0)
				End If
			Next parse
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
			'UPGRADE_WARNING: A boolean is being converted to a numeric type. The numeric type of True has changed from -1 in VB6 to 1 in VB7.
			bwPhand = OpenProcess(PROCESS_ALL_ACCESS, False, bwPid)
			If (bwPhand = 0) Then
				ChangeNick = "Error 02: Unable to retreive Brood War process handle."
				Exit Function
			End If
			'Write textbox to memory location:
			' Version 1.07: 0x1902b0a0
			' Version 1.08O: 0x1902E4D8
			WriteProcessMemory(bwPhand, &H1902B0A0, strNewName.Value, 21, LengthReturned)
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
			If Ascii = 8 Then
				IsAcceptable = True : Exit Function
			End If
			If Ascii = 33 Then
				IsAcceptable = True : Exit Function
			End If
			If Ascii = 36 Then
				IsAcceptable = True : Exit Function
			End If
			If Ascii > 38 And Ascii < 42 Then
				IsAcceptable = True : Exit Function
			End If
			If Ascii = 43 Then
				IsAcceptable = True : Exit Function
			End If
			If Ascii > 44 And Ascii < 47 Then
				IsAcceptable = True : Exit Function
			End If
			If Ascii > 47 And Ascii < 60 Then
				IsAcceptable = True : Exit Function
			End If
			If Ascii = 61 Then
				IsAcceptable = True : Exit Function
			End If
			If Ascii = 64 Then
				IsAcceptable = True : Exit Function
			End If
			If Ascii > 64 And Ascii < 92 Then
				IsAcceptable = True : Exit Function
			End If
			If Ascii > 92 And Ascii < 127 Then
				IsAcceptable = True : Exit Function
			End If
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
			If EncodedASCII = 14 Then DecodeToBWCompatible = 2 ' blue
			If EncodedASCII = 15 Then DecodeToBWCompatible = 3 ' green
			If EncodedASCII = 16 Then DecodeToBWCompatible = 4 ' highlighted green
			If EncodedASCII = 17 Then DecodeToBWCompatible = 5 ' black
			If EncodedASCII = 18 Then DecodeToBWCompatible = 6 ' white
			If EncodedASCII = 19 Then DecodeToBWCompatible = 7 ' red
			If EncodedASCII = 20 Then DecodeToBWCompatible = 8 ' invisible
		End Function
	End Module
End Namespace