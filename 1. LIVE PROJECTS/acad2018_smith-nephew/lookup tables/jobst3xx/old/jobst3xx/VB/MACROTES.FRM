VERSION 5.00
Begin VB.Form MacroTest 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1800
   ClientLeft      =   888
   ClientTop       =   1512
   ClientWidth     =   3372
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1800
   ScaleWidth      =   3372
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   492
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run Macro"
      Height          =   492
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1092
   End
End
Attribute VB_Name = "MacroTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   JOBSTART.MAK
'Purpose:   General procedure to start JOBST CAD modules.
'
'           Ensures that no more than one instance of
'           of a JOBST CAD system VB module can exist.
'           This stops the problem of DRAFIX macros opening
'           a DDE conversation with the wrong instance.
'
'
'Version:   3.01
'Date:      10.Jun.94
'Author:    Gary George
'---------------------------------------------------------------------------------------------
'REVISIONS:
'Date       By      Action
'---------------------------------------------------------------------------------------------
'Dec 98     GG      Ported to VB 5

   '* Windows API Functions Declarations
'    Private Declare Function GetWindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
'    Private Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
'    Private Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFilename$)
'
'   'Constanst used by GetWindow
'    Const GW_CHILD = 5
'    Const GW_HWNDFIRST = 0
'    Const GW_HWNDLAST = 1
'    Const GW_HWNDNEXT = 2
'    Const GW_HWNDPREV = 3
'    Const GW_OWNER = 4

    Public MainForm As MacroTest
    
    Option Explicit

Private Sub Command1_Click()
    On Error Resume Next

    Dim sDrafixTask As String
    Dim sTask As String
    Dim sMacro As String
    Dim sCommand As String
    Dim sDrafixText As String
    Dim sTaskText As String
    Dim PathJOBST As String
    
 
    sCommand = "@C:\CADLINK\Imagea.D{enter}"
   
    sDrafixText = "Drafix CAD Professional"

    PR_GetDrafixandTaskWindowTitleText sDrafixTask, sTask, sDrafixText, sTaskText
    
   'Start the required task by using the appropriate
   'drafix macro
    If sDrafixTask <> "" Then
        AppActivate sDrafixTask
        SendKeys sCommand, True
    Else
        MsgBox "Can't find a specific drawing to work with!", 16, "Starting CAD Module"
    End If

 

  
End Sub

Private Sub Command2_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
 
    If App.PrevInstance Then End

    
    Set MainForm = MacroTest
    
    Show

End Sub

Private Sub PR_GetDrafixandTaskWindowTitleText(sDrafixTask As String, sInstanceTask As String, sDrafixText As String, sInstanceText As String)
   'Returns the
   '    Drafix Window Title Text
   '    Requested Instance Window Title Text
   '
   'N.B. Warning
   '
   '    Returns last Drafix task found.
   '    This is OK for Drafix 2.1e as only a single instance
   '    of Drafix Windows CAD is allowed.
   '    Dangerous for Drafix 3.

    Dim sTask As String
    Dim nLength As Long
    Dim nCurrWnd As Long
    Dim X As Long
    Dim nDrafixText As Long
    Dim nInstanceText As Long

   'Get the nWnd of the first item in the master list
   'so we can process the task list entries (top level only)
    nCurrWnd = GetWindow(MainForm.hwnd, GW_HWNDFIRST)

   'Loop to locate Drafix CAD task and Requested Instance
    sDrafixTask = ""
    nDrafixText = Len(sDrafixText)
    sInstanceTask = ""
    nInstanceText = Len(sInstanceText)

    While nCurrWnd <> 0

       'Extract details of task
        nLength = GetWindowTextLength(nCurrWnd)
        sTask = Space$(nLength + 1)
        nLength = GetWindowText(nCurrWnd, sTask, nLength + 1)
        
        If Left(sTask, nDrafixText) = sDrafixText Then sDrafixTask = sTask
        If Left(sTask, nInstanceText) = sInstanceText Then sInstanceTask = sTask
       
       'Get next task from master list
        nCurrWnd = GetWindow(nCurrWnd, GW_HWNDNEXT)

       'Process Windows events
        X = DoEvents()

    Wend

End Sub

