VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form merge 
   Caption         =   "Merge Drawings"
   ClientHeight    =   3708
   ClientLeft      =   696
   ClientTop       =   1992
   ClientWidth     =   5304
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
   LinkMode        =   1  'Source
   LinkTopic       =   "merge"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3708
   ScaleWidth      =   5304
   Begin VB.CommandButton CAD_File 
      Cancel          =   -1  'True
      Caption         =   "Find ..."
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3180
      Width           =   1275
   End
   Begin VB.TextBox txtCurrentCADFile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6300
      TabIndex        =   12
      Text            =   "txtCurrentCADFile"
      Top             =   2520
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   6300
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      DialogTitle     =   "Open CAD File"
      FontSize        =   0
      MaxFileSize     =   256
      PrinterDefault  =   0   'False
   End
   Begin VB.Frame frmCADFile 
      Caption         =   "Drawing to Merge"
      Height          =   825
      Left            =   180
      TabIndex        =   2
      Top             =   2160
      Width           =   4935
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Use Find to select drawing to merge!"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   4395
      End
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3780
      TabIndex        =   0
      Top             =   3180
      Width           =   1275
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3180
      Width           =   1275
   End
   Begin VB.Frame frmWODetails 
      Caption         =   "Current Drawing"
      Height          =   1815
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   4935
      Begin VB.TextBox txtOrderDate 
         Height          =   315
         Left            =   3360
         TabIndex        =   4
         Text            =   "txtOrderDate"
         Top             =   1260
         Width           =   1335
      End
      Begin VB.TextBox txtWorkOrder 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Text            =   "txtWorkOrder"
         Top             =   1260
         Width           =   1035
      End
      Begin VB.TextBox txtPatientName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Text            =   "txtPatientName"
         Top             =   780
         Width           =   3615
      End
      Begin VB.TextBox txtFileNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Text            =   "txtFileNo"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label labPatientDetails 
         Caption         =   "File No:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   8
         Top             =   405
         Width           =   675
      End
      Begin VB.Label labPatientDetails 
         Caption         =   "Name:"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   9
         Top             =   810
         Width           =   615
      End
      Begin VB.Label labWODetails 
         Caption         =   "W/O Date:"
         Height          =   240
         Index           =   2
         Left            =   2280
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label labWODetails 
         Caption         =   "W/Order:"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Left            =   6360
      Top             =   1740
   End
End
Attribute VB_Name = "merge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   Merge.MAK
'Purpose:   To allow the user to locate a cad drawing file to
'           merge with the current darwing
'Version:   1.00
'Date:      06.Feb.98
'Author:    Gary George
'---------------------------------------------------------------------------------------------
'REVISIONS:
'Date       By      Action
'---------------------------------------------------------------------------------------------
'Dec 98     GG      Ported to VB5
'
'NOTE:-
'
'
    Dim m_sDrawingToMergeFullPath   As String
    Dim m_sDrawingToMergeTitleOnly  As String
    Option Explicit

Private Sub CAD_File_Click()
   'Find an existing CAD file to merge
   '
   'Path to drawings
    Dim lpBuffer As String * 144 'Minimum recommended wrt GetWindowsDirectory()
    Dim nBufferSize%, nSize%
    Dim WindowsDir$

    nBufferSize% = 143

   'Get the path to the Windows Directory to locate DRAFIX.INI
   '
    nSize% = GetWindowsDirectory(lpBuffer, nBufferSize%)
    WindowsDir$ = Left$(lpBuffer, nSize%)

   'Get the path to the Drawing directory from
   'DRAFIX.INI
   '
    nSize% = GetPrivateProfileString("Path", "PathDrawing", "C:\", lpBuffer, nBufferSize%, WindowsDir$ & "\DRAFIX.INI")

    CMDialog1.InitDir = Left$(lpBuffer, nSize%)
    CMDialog1.Filter = "Drafix CAD (*.cad)|*.cad"
    CMDialog1.DialogTitle = "CAD File to Merge"
    CMDialog1.filename = ""
    CMDialog1.DefaultExt = "cad"
    CMDialog1.Action = 1
    
    If CMDialog1.filename <> "" Then
        If Dir$(CMDialog1.filename) = "" Then
            MsgBox "CAD File not found.", 48, "Merge Drawings"
        Else
            Label1.Caption = CMDialog1.filename
            m_sDrawingToMergeFullPath = CMDialog1.filename
            m_sDrawingToMergeTitleOnly = CMDialog1.FileTitle
        End If
    End If
End Sub

Private Sub Cancel_Click()
    End
End Sub

Private Function FN_EscapeSlashesInString(sAssignedString As Variant) As String
   'Search through the string looking for " (double quote characater)
   'If found use \ (Backslash) to escape it
   '
    Dim ii%
    Dim Char$
    Dim sEscapedString$
    
    FN_EscapeSlashesInString = ""
    
    For ii = 1 To Len(sAssignedString)
        Char = Mid$(sAssignedString, ii, 1)
        If Char = "\" Then
            sEscapedString = sEscapedString & "\" & Char
        Else
            sEscapedString = sEscapedString & Char
        End If
    Next ii
    
    FN_EscapeSlashesInString = sEscapedString

End Function


Private Sub Form_LinkClose()
   'Disable TIMEOUT Timer
    Timer1.Enabled = False
   
    Show

End Sub

Private Sub Form_Load()
   'Hide form to allow DDE to happen
   'Show form on LinkClose
    Hide

   'Clear Text Boxes
   'Work Order and TXF file Name
    txtWorkOrder = ""

   'Patient Details
    txtFileNo = ""
    txtPatientName = ""
    txtOrderDate = ""
    txtCurrentCADFile = ""

   'Position to center of screen
    Left = (Screen.Width - merge.Width) / 2
    Top = (Screen.Height - merge.Height) / 2

    Set MainForm = merge
    
   'Get the path to the JOBST system installation directory
    g_sPathJOBST = fnPathJOBST()
    
   'Ensure that we timeout after approx 6 seconds
   'The timer is disabled on link close
    Timer1.Interval = 6000
    Timer1.Enabled = False

End Sub

Private Sub OK_Click()
   'Check that with the user that data is correct
   'If OK the update drawing
    Dim sDrafixTask         As String
    Dim sDrawingToMerge     As String
    Dim sMessage        As String
    
   'Disable to stop double click
    OK.Enabled = False
    sMessage = ""

   'Check for existance of file
    If m_sDrawingToMergeFullPath = "" Then GoTo ExitSub_OK_Click
    
    If m_sDrawingToMergeFullPath = txtCurrentCADFile.Text Then sMessage = "Selected Drawing File is the same as the current Drawing File!" & NL

    If Dir$(m_sDrawingToMergeFullPath) = "" Then sMessage = sMessage & "CAD Drawing File not found!" & NL
    
    If UCase(Right$(m_sDrawingToMergeFullPath, 4)) <> ".CAD" Then sMessage = sMessage & "The selected file is not a CAD Drawing File!" & NL
    
    If Len(sMessage) > 0 Then
        MsgBox sMessage, 48, "Errors selecting file"
        GoTo ExitSub_OK_Click
    End If

    PR_DRAFIX_Macro "C:\JOBST\DRAW.D"

    sDrafixTask = fnGetDrafixWindowTitleText()

    'Start the required task by using the appropriate
    'drafix macro
    If sDrafixTask <> "" Then
        AppActivate sDrafixTask
        SendKeys "@C:\JOBST\DRAW.D{enter}", True
        End
    Else
        MsgBox "Unable to find a running copy of DRAFIX!", 16, "Error - Merging Drawing"
    End If
    
ExitSub_OK_Click:

    OK.Enabled = True

End Sub

Private Sub PR_DRAFIX_Macro(sDrafixFile As String)
   'Create a DRAFIX macro file
   'NOTES
   '
    Dim sString$
    
   'Open file
    fNum = FreeFile
    Open sDrafixFile For Output As fNum

   'Initialise String globals
    CC = Chr$(44) 'The comma (,)
    NL = Chr$(10) 'The new line character
    QQ = Chr$(34) 'Double quotes (")
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ
     
   'Write header information etc. to the DRAFIX macro file
   '
    Print #fNum, "//DRAFIX Macro created - "; Date$; "  "; Time$
    Print #fNum, "//Patient    - "; txtPatientName.Text; CC; " "; txtFileNo.Text; CC
    Print #fNum, "//Work Order - "; txtWorkOrder.Text
    Print #fNum, "//by Visual Basic"
   
'    Print #fNum, "@" & FN_EscapeSlashesInString(g_sPathJOBST) & "\\CADLINK\\MERGE.D "; QQ; FN_EscapeSlashesInString((m_sDrawingToMergeFullPath)); QQ; ";"
    Print #fNum, "@"; g_sPathJOBST & "\CADLINK\MERGE.D "; m_sDrawingToMergeFullPath; ";"
    
   'Close the Macro File
    Print #fNum, ""
    Print #fNum, "// -- End of MACRO --"
    Close #fNum
   
End Sub

Private Sub PR_GetDrafixWindowTitleText(sDrafixText$, sDrafixTask$)
   'Returns the
   '    Drafix Window Title Text in sDrafixTask$
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
    Dim X As Integer
    Dim nDrafixText%, nInstanceText%

   'Get the nWnd of the first item in the master list
   'so we can process the task list entries (top level only)
    nCurrWnd = GetWindow(MainForm.hwnd, GW_HWNDFIRST)

   'Loop to locate Drafix CAD task and Requested Instance
    sDrafixTask = ""
    nDrafixText = Len(sDrafixText)

    While nCurrWnd <> 0

       'Extract details of task
        nLength = GetWindowTextLength(nCurrWnd)
        sTask = Space$(nLength + 1)
        nLength = GetWindowText(nCurrWnd, sTask, nLength + 1)
        
        If Left(sTask, nDrafixText) = sDrafixText Then sDrafixTask = sTask
       
       'Get next task from master list
        nCurrWnd = GetWindow(nCurrWnd, GW_HWNDNEXT)

       'Process Windows events
        X = DoEvents()

    Wend

End Sub

Private Sub PR_ReadLine(ByVal fFile As Integer, sLine As String)
    'Read a line, character at a time up to either the
    '   NL character
    'or
    '   CRNL characters
    '
    Dim NewLine$, CR$, Char As String * 1
    NewLine = Chr$(10)
    CR = Chr$(13)
    sLine = ""

    Do Until EOF(fFile)
        Char = Input$(1, #fFile)
        If Char = NewLine Then Exit Do
        If Char <> CR Then sLine = sLine & Char
    Loop

End Sub

Private Sub Timer1_Timer()
   'In case there is no Link_Close event
   'The programme will time out after approx 5 secs
    End
End Sub

