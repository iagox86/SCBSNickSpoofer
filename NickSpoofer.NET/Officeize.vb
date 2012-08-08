Option Strict Off
Option Explicit On
Module Officeize
	'UPGRADE_WARNING: Structure NOTIFYICONDATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup1050.htm
	Declare Function Shell_NotifyIcon Lib "shell32.dll"  Alias "Shell_NotifyIconA"(ByVal dwMessage As Integer, ByRef lpData As NOTIFYICONDATA) As Integer
	
	Public Structure NOTIFYICONDATA
		Dim cbSize As Integer
		Dim hwnd As Integer
		Dim uID As Integer
		Dim uFlags As Integer
		Dim uCallbackMessage As Integer
		Dim hIcon As Integer
		<VBFixedString(64),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr,SizeConst:=64)> Dim szTip As String
	End Structure
	
	Public Const NIM_ADD As Short = &H0s
	Public Const NIM_MODIFY As Short = &H1s
	Public Const NIM_DELETE As Short = &H2s
	Public Const NIF_MESSAGE As Short = &H1s
	Public Const NIF_ICON As Short = &H2s
	Public Const NIF_TIP As Short = &H4s
	
	'Make your own constant, e.g.:
	Public Const NIF_DOALL As Boolean = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
	
	Public Const WM_MOUSEMOVE As Short = &H200s
	Public Const WM_LBUTTONDBLCLK As Short = &H203s
	Public Const WM_LBUTTONDOWN As Short = &H201s
	Public Const WM_RBUTTONDOWN As Short = &H204s
	
	
	
	Private Declare Function GetWindowLong Lib "user32"  Alias "GetWindowLongA"(ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
	Private Declare Function SetWindowLong Lib "user32"  Alias "SetWindowLongA"(ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal CX As Integer, ByVal CY As Integer, ByVal wFlags As Integer) As Integer
	
	Private Const GWL_EXSTYLE As Short = (-20)
	Private Const WS_EX_CLIENTEDGE As Short = &H200s
	Private Const WS_EX_STATICEDGE As Integer = &H20000
	
	Private Const SWP_FRAMECHANGED As Short = &H20s
	Private Const SWP_NOMOVE As Short = &H2s
	Private Const SWP_NOOWNERZORDER As Short = &H200s
	Private Const SWP_NOSIZE As Short = &H1s
	Private Const SWP_NOZORDER As Short = &H4s
	
	
	Public Function AddOfficeBorder(ByVal hwnd As Integer) As Object
		
		Dim lngRetVal As Integer
		'Retrieve the current border style
		lngRetVal = GetWindowLong(hwnd, GWL_EXSTYLE)
		
		'Calculate border style to use
		lngRetVal = lngRetVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
		
		'Apply the changes
		SetWindowLong(hwnd, GWL_EXSTYLE, lngRetVal)
		SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED)
		
	End Function
End Module