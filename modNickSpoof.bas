Attribute VB_Name = "modNickSpoof"
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
Public NamePos As Integer
'for retrieving the game process handle
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'get processID from hwnd returned above
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'return handle to target process
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'for writing game memory
Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'for reading game memory
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'close each open handle
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'for determining OS version
Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
'User defined type for OSVERSIONINFO
    Public Type OSVERSIONINFO
       dwOSVersionInfoSize As Long
       dwMajorVersion As Long
       dwMinorVersion As Long
       dwBuildNumber As Long
       dwPlatformId As Long           '1 = Windows 95.
                                      '2 = Windows NT

       szCSDVersion As String * 128
    End Type
'Necessary flags for NT
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)


Function ReadNick() As String
'Dimention variables
'Brood War Window Handle
Dim bwWnd As Long
'Brood War Process ID
Dim bwPid As Long
'Brood War Process Handle
Dim bwPhand As Long
'Length returned
Dim LengthReturned As Long
'String for containing nickname
Dim strNick As String * 21
'end dimentioning

'Get Brood War processid
bwWnd = FindWindow(vbNullString, "Brood War")
'make sure Brood War is open
    If (bwWnd = 0) Then
        ReadNick = "Error 01: Brood War is not open."
        Exit Function
    End If
'Convert window handle to process id
GetWindowThreadProcessId bwWnd, bwPid

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

ReadProcessMemory bwPhand, &H190314E8, strNick, 21, LengthReturned
    If (LengthReturned = 0) Then
        ReadNick = "Error 03: You must connect to Battle.net."
        CloseHandle bwPhand
        Exit Function
    End If
'Close the handle we opened.
CloseHandle bwPhand
'And finally, return the nickname read.
ReadNick = strNick
End Function
Function ChangeNick(NewNick As String) As String
'Dimention variables
'Brood War Window Handle
Dim bwWnd As Long
'Brood War Process ID
Dim bwPid As Long
'Brood War Process Handle
Dim bwPhand As Long
'Length returned
Dim LengthReturned As Long
'String used to write to memory
Dim strNewName As String * 21
'end dimentioning
'Get Textbox data
strNewName = NewNick
For parse = 1 To 21
    'Convert any trailing spaces in the string to an ascii 00 value.
    If (parse > Len(NewNick)) Then
        If Mid$(strNewName, parse, 1) = " " Then Mid$(strNewName, parse, 1) = Chr$(0)
    End If
Next parse
'Get Brood War processid
bwWnd = FindWindow(vbNullString, "Brood War")
'make sure Brood War is open
    If (bwWnd = 0) Then
        ChangeNick = "Error 01: Brood War window not found. Brood War must be open to spoof your nickname."
        Exit Function
    End If
'Convert window handle to process id
GetWindowThreadProcessId bwWnd, bwPid

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
WriteProcessMemory bwPhand, &H190314E8, strNewName, 21, LengthReturned
    If (LengthReturned = 0) Then
           ChangeNick = "Error 04: No data entered into Brood War. You must be connected to Battle.net to spoof your nickname."
           CloseHandle bwPhand
           Exit Function
    End If
'And finally, close the handle we opened.
CloseHandle bwPhand
End Function
Function IsAcceptable(Ascii As Integer) As Boolean
'Check if Ascii is acceptable in battle.net nicknames
If Ascii = 8 Then IsAcceptable = True: Exit Function
If Ascii = 33 Then IsAcceptable = True: Exit Function
If Ascii = 36 Then IsAcceptable = True: Exit Function
If Ascii > 38 And Ascii < 42 Then IsAcceptable = True: Exit Function
If Ascii = 43 Then IsAcceptable = True: Exit Function
If Ascii > 44 And Ascii < 47 Then IsAcceptable = True: Exit Function
If Ascii > 47 And Ascii < 60 Then IsAcceptable = True: Exit Function
If Ascii = 61 Then IsAcceptable = True: Exit Function
If Ascii = 64 Then IsAcceptable = True: Exit Function
If Ascii > 64 And Ascii < 92 Then IsAcceptable = True: Exit Function
If Ascii > 92 And Ascii < 127 Then IsAcceptable = True: Exit Function
IsAcceptable = False
End Function
Function EncodeToASCII(BWColor As String) As Integer
If BWColor = "Blue" Then EncodeToASCII = 14
If BWColor = "Green" Then EncodeToASCII = 15
If BWColor = "Highlighted Green" Then EncodeToASCII = 16
If BWColor = "Black" Then EncodeToASCII = 17
If BWColor = "White" Then EncodeToASCII = 18
If BWColor = "Red" Then EncodeToASCII = 19
If BWColor = "Invisible" Then EncodeToASCII = 20
End Function
Function DecodeToBWCompatible(EncodedASCII As Integer) As Integer
If EncodedASCII = 14 Then DecodeToBWCompatible = 2  'blue
If EncodedASCII = 15 Then DecodeToBWCompatible = 3  'green
If EncodedASCII = 16 Then DecodeToBWCompatible = 4  'highlighted green
If EncodedASCII = 17 Then DecodeToBWCompatible = 5  'black
If EncodedASCII = 18 Then DecodeToBWCompatible = 6  'white
If EncodedASCII = 19 Then DecodeToBWCompatible = 7  'red
If EncodedASCII = 20 Then DecodeToBWCompatible = 8  'invisible
End Function
Function getVersion() As Long
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer
osinfo.dwOSVersionInfoSize = 148
osinfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(osinfo)
getVersion = osinfo.dwPlatformId
End Function
