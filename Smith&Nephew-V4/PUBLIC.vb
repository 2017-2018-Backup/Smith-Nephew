Option Strict Off
Option Explicit On
'UPGRADE_NOTE: Public was upgraded to Public_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
Module Public_Renamed
	'Project:   COMMON
	'File:      Public.BAS
	'Purpose:   Declaration of APIs and Constants
	'           Used by all CAD Systm Programmes
	'
	'Version:   1.00
	'Date:      10.Dec.98
	'Author:    Gary George
	'
	'-------------------------------------------------------
	'REVISIONS:
	'Date       By      Action
	'-------------------------------------------------------
	'
	'Notes:-
	
	Public Declare Function GetActiveWindow Lib "user32" () As Integer
	Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
	Public Declare Function GetWindowText Lib "user32"  Alias "GetWindowTextA"(ByVal hwnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer
	Public Declare Function GetWindowTextLength Lib "user32"  Alias "GetWindowTextLengthA"(ByVal hwnd As Integer) As Integer
	Public Declare Function GetWindowsDirectory Lib "kernel32"  Alias "GetWindowsDirectoryA"(ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Public Declare Function GetTempFileName Lib "kernel32"  Alias "GetTempFileNameA"(ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer
    Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Integer, ByVal lpBuffer As String) As Integer
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Boolean


    'Constanst used by GetWindow
    Public Const GW_CHILD As Short = 5
	Public Const GW_HWNDFIRST As Short = 0
	Public Const GW_HWNDLAST As Short = 1
	Public Const GW_HWNDNEXT As Short = 2
	Public Const GW_HWNDPREV As Short = 3
	Public Const GW_OWNER As Short = 4
	
	'MsgBox constant
	Public Const IDCANCEL As Short = 2
	Public Const IDYES As Short = 6
	Public Const IDNO As Short = 7
	Public Const IDOK As Short = 1 ' OK button pressed
	Public Const IDABORT As Short = 3 ' Abort button pressed
	Public Const IDRETRY As Short = 4 ' Retry button pressed
	Public Const IDIGNORE As Short = 5 ' Ignore button pressed
	
	
	Public Const MAX_PATH As Short = 260
	
	Public Function fnGetDrafixWindowTitleText() As String
		
		'Returns the Drafix Window Title Text
		'
		'If Drafix task found
		'        Return the Drafix Window Title Text
		'    else
		'        Return an empty string
		'N.B.
		'    Returns first Drafix task found.
		'    This is OK for Drafix 2.1e as only a single instance
		'    of Drafix Windows CAD is allowed.
		
		Dim sTask As String
		Dim X As Object
		Dim nLength As Integer
		Dim nCurrWnd As Integer


        'Get the nWnd of the first item in the master list
        'so we can process the task list entries (top level only)
        nCurrWnd = GetWindow(ARMDIA1.MainForm.Handle.ToInt32, GW_HWNDFIRST)

        'Loop to locate Drafix CAD task
        While nCurrWnd <> 0
			
			'Extract details of task
			nLength = GetWindowTextLength(nCurrWnd)
			sTask = Space(nLength + 1)
			nLength = GetWindowText(nCurrWnd, sTask, nLength + 1)
			'If task is "Drafix" then return Task title text
			If Left(sTask, 6) = "Drafix" Then
				fnGetDrafixWindowTitleText = sTask
				Exit Function
			End If
			
			'Get next task from master list
			nCurrWnd = GetWindow(nCurrWnd, GW_HWNDNEXT)

            'Process Windows events
            'UPGRADE_ISSUE: DoEvents does not return a value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8D115264-E27F-4472-A684-865A00B5E826"'
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            System.Windows.Forms.Application.DoEvents() 'X = System.Windows.Forms.Application.DoEvents()

        End While
		
	End Function

    Public Function fnPathJOBST() As String
        'Function to find the installation directory of the JOBST CAD
        'system.
        'This information in stored in the [JOBST] section of the DRAFIX.INI
        'file saved in the Windows directory.
        '
        '    [JOBST]
        '    PathJOBST=C:\JOBST
        '        .   .
        '        .   .
        '
        'Returns:-
        '    the value given in the entry PathJOBST
        ' or
        '    an empty string if entry not found.
        '
        '

        'API Variables
        Dim lpBuffer As New VB6.FixedLengthString(144) 'Minimum recommended wrt GetWindowsDirectory()
        Dim nBufferSize As Integer
        Dim nSize As Integer
        nBufferSize = 143

        Dim WindowsDir As String

        'Get the path to the Windows Directory to locate DRAFIX.INI
        '
        nSize = GetWindowsDirectory(lpBuffer.Value, nBufferSize)
        WindowsDir = Left(lpBuffer.Value, nSize)

        'Get the path to the installed JOBST CAD System from
        'DRAFIX.INI
        '
        nSize = GetPrivateProfileString("JOBST", "PathJOBST", "", lpBuffer.Value, nBufferSize, WindowsDir & "\DRAFIX.INI")
        fnPathJOBST = Left(lpBuffer.Value, nSize)

    End Function

    Public Function fnGetSettingsPath(ByRef sKey As String) As String
        'Function to find the installation directory of the JOBST CAD
        'system.
        'API Variables
        Dim lpBuffer As New VB6.FixedLengthString(144) 'Minimum recommended wrt GetWindowsDirectory()
        Dim nBufferSize As Integer
        Dim nSize As Integer
        nBufferSize = 143
        Dim WindowsDir As String

        'Get the path to the Windows Directory to locate DRAFIX.INI
        nSize = GetWindowsDirectory(lpBuffer.Value, nBufferSize)
        WindowsDir = Left(lpBuffer.Value, nSize)

        'Get the path to the installed JOBST CAD System from
        Dim sDrawPath As String = "C:\ACAD2018_SMITH-NEPHEW\Interface"
        ''nSize = GetPrivateProfileString("DRAW", "PathDRAW", "", sDrawPath, nBufferSize, sDrawPath & "\SETTINGS.INI")
        nSize = GetPrivateProfileString("DRAW", sKey, "", lpBuffer.Value, nBufferSize, sDrawPath & "\SETTINGS.INI")
        fnGetSettingsPath = Left(lpBuffer.Value, nSize)
    End Function
    Public Function fnWriteValueInSettingsFile(ByRef sKey As String, ByRef sValue As String)
        'Get the path to the installed JOBST CAD System from
        Dim sDrawPath As String = "C:\ACAD2018_SMITH-NEPHEW\Interface"
        ''nSize = GetPrivateProfileString("DRAW", "PathDRAW", "", sDrawPath, nBufferSize, sDrawPath & "\SETTINGS.INI")
        Dim bIsWrite As Boolean = WritePrivateProfileString("DRAW", sKey, sValue, sDrawPath & "\SETTINGS.INI")
    End Function
End Module