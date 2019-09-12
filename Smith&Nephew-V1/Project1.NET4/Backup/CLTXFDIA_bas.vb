Option Strict Off
Option Explicit On
Module CLTXFDIA1
	
	'   '* Windows API Functions Declarations
	'    Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
	'    Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
	'    Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
	'    Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
	'    Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
	
	'Constanst used by GetWindow
	'    Global Const GW_CHILD = 5
	'    Global Const GW_HWNDFIRST = 0
	'    Global Const GW_HWNDLAST = 1
	'    Global Const GW_HWNDNEXT = 2
	'    Global Const GW_HWNDPREV = 3
	'    Global Const GW_OWNER = 4
	'Global Declarations
	'
	
	Public MainForm As cltxfdia
	
	Public g_sPathJOBST As String
	
	Public g_sSexTXF As String
	Public g_sUnits As String
	Public g_sDiagnosisTXF As String
	
	'Globals used to create DRAFIX Macros
	'
	Public fNum As Short
	'
	Public CC As String 'The comma (,)
	Public NL As String 'The new line character
	Public TABBED As String 'The tab char character
	Public QQ As String 'Double quotes (")
	Public QCQ As String 'QQ & CC & QQ
	Public QC As String 'QQ & CC
	Public CQ As String 'CC & QQ
End Module