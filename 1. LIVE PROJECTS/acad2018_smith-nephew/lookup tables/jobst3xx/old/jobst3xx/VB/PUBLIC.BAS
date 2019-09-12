Attribute VB_Name = "Public"
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

    Public Declare Function GetActiveWindow Lib "user32" () As Long
    Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
    Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
    Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

   
   'Constanst used by GetWindow
    Public Const GW_CHILD = 5
    Public Const GW_HWNDFIRST = 0
    Public Const GW_HWNDLAST = 1
    Public Const GW_HWNDNEXT = 2
    Public Const GW_HWNDPREV = 3
    Public Const GW_OWNER = 4

   'MsgBox constant
    Public Const IDCANCEL = 2
    Public Const IDYES = 6
    Public Const IDNO = 7
    Public Const IDOK = 1                  ' OK button pressed
    Public Const IDABORT = 3               ' Abort button pressed
    Public Const IDRETRY = 4               ' Retry button pressed
    Public Const IDIGNORE = 5              ' Ignore button pressed


    Public Const MAX_PATH = 260

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

    Dim sTask$, X  As Variant
    Dim nLength As Long
    Dim nCurrWnd As Long
    

   'Get the nWnd of the first item in the master list
   'so we can process the task list entries (top level only)
    nCurrWnd = GetWindow(MainForm.hwnd, GW_HWNDFIRST)

   'Loop to locate Drafix CAD task
    While nCurrWnd <> 0

       'Extract details of task
        nLength = GetWindowTextLength(nCurrWnd)
        sTask = Space$(nLength + 1)
        nLength = GetWindowText(nCurrWnd, sTask, nLength + 1)
        'If task is "Drafix" then return Task title text
        If Left(sTask, 6) = "Drafix" Then
            fnGetDrafixWindowTitleText = sTask
            Exit Function
        End If
       
       'Get next task from master list
        nCurrWnd = GetWindow(nCurrWnd, GW_HWNDNEXT)

       'Process Windows events
        X = DoEvents()

    Wend

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
    Dim lpBuffer As String * 144 'Minimum recommended wrt GetWindowsDirectory()
    Dim nBufferSize As Long
    Dim nSize As Long
    nBufferSize = 143

    Dim WindowsDir As String

   'Get the path to the Windows Directory to locate DRAFIX.INI
   '
    nSize = GetWindowsDirectory(lpBuffer, nBufferSize)
    WindowsDir = Left$(lpBuffer, nSize)

   'Get the path to the installed JOBST CAD System from
   'DRAFIX.INI
   '
    nSize = GetPrivateProfileString("JOBST", "PathJOBST", "", lpBuffer, nBufferSize, WindowsDir & "\DRAFIX.INI")
    fnPathJOBST = Left$(lpBuffer, nSize)

End Function

