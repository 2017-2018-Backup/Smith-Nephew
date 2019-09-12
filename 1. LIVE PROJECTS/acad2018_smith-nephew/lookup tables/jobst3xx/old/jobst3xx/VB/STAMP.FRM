VERSION 5.00
Begin VB.Form stamp 
   Caption         =   "Stamps"
   ClientHeight    =   4305
   ClientLeft      =   2970
   ClientTop       =   3090
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "stamp"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4305
   ScaleWidth      =   6585
   Begin VB.CheckBox chkArrow 
      Caption         =   "Arrow"
      Height          =   315
      Left            =   5280
      TabIndex        =   14
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtStampType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7920
      TabIndex        =   11
      Top             =   660
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton Edit_USER 
      Caption         =   "Edit USER"
      Height          =   315
      Left            =   5220
      TabIndex        =   10
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stamps For"
      Height          =   2295
      Left            =   5100
      TabIndex        =   4
      Top             =   780
      Width           =   1335
      Begin VB.OptionButton OptionUSER 
         Caption         =   "USER"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1860
         Width           =   975
      End
      Begin VB.OptionButton OptionGeneral 
         Caption         =   "General"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1500
         Width           =   1095
      End
      Begin VB.OptionButton OptionWaistHt 
         Caption         =   "Waist Ht"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1140
         Width           =   1095
      End
      Begin VB.OptionButton OptionLeg 
         Caption         =   "Leg"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   780
         Width           =   1095
      End
      Begin VB.OptionButton OptionArm 
         Caption         =   "Arm"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   1155
      End
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3300
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtStamp 
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   360
      Width           =   4755
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ListBox StampList 
      Height          =   2595
      Left            =   180
      TabIndex        =   0
      Top             =   840
      Width           =   4755
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DDE Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   15
      Top             =   360
      Width           =   1035
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Stamp Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Stamp Text"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "stamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   '* Windows API Functions Declarations
'    Private Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
'    Private Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
'    Private Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
'
'   'Constanst used by GetWindow
'    Const GW_CHILD = 5
'    Const GW_HWNDFIRST = 0
'    Const GW_HWNDLAST = 1
'    Const GW_HWNDNEXT = 2
'    Const GW_HWNDPREV = 3
'    Const GW_OWNER = 4

Private Sub Cancel_Click()
    End
End Sub

Private Sub Edit_USER_Click()
    EditUSER.Show 1
   'Force click event to re-read values from file
    OptionUSER.Value = False
    OptionUSER.Value = True
End Sub


Private Function FN_TextCheck(sLine) As String
   'Checks for the following DRAFIX special characters and escape
   'them using the backslash \
   'This function is required as the characters have a special
   'meaning to the DRAFIX interpreter.
   '
   'Check for :-
   '    (") - single double quote
   '    (\) - the backslash
   '
    Dim sNewline As String, sChar As String
    Dim QUOTE As String, BACKSLASH As String
    Dim ii As Integer

    QUOTE = Chr$(34)
    BACKSLASH = "\"

    For ii = 1 To Len(sLine)
        sChar = Mid$(sLine, ii, 1)
        If sChar = BACKSLASH Or sChar = QUOTE Then
        sNewline = sNewline + BACKSLASH + sChar
        Else
            sNewline = sNewline + sChar
        End If
    Next ii

    FN_TextCheck = sNewline

End Function

Private Sub Form_LinkClose()
    
    Select Case txtStampType.Text
        Case "Arm"
            OptionArm.Value = True
        Case "Waist Ht"
            OptionWaistHt.Value = True
        Case "Leg"
            OptionLeg.Value = True
        Case "General"
            OptionGeneral.Value = True
        Case "USER"
            OptionUSER.Value = True
    End Select

    Screen.MousePointer = 0    ' Change pointer to hourglass.

    Show
End Sub

Private Sub Form_Load()
    Hide
   'Check if a previous instance is running
   'If it is warn user and exit
    If App.PrevInstance Then
        MsgBox "The Stamp Module is already running!" + Chr(13) + "Use ALT-TAB and Cancel it.", 16, "Error Starting Stamps"
        End
    End If

   'Maintain while loading DDE data
    Screen.MousePointer = 11    ' Change pointer to hourglass.
   'Reset in Form_LinkClose
    
    Left = (Screen.Width - stamp.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - stamp.Height) / 2  ' Center form vertically.
    Edit_USER.Enabled = False
    
    Set MainForm = stamp
    
End Sub

Private Sub OK_Click()
    prInsertStamp
End Sub

Private Sub OptionArm_Click()
'Stamps for LEG
    Edit_USER.Enabled = False
    StampList.Clear
    StampList.AddItem "1/2" & Chr$(34) & " ELASTIC"
    StampList.AddItem "2" & Chr$(34) & " ELASTIC"
    StampList.AddItem "SILICONE ELASTIC"
    StampList.AddItem "NO ELASTIC"
    StampList.AddItem ""
    StampList.AddItem "LINING"
    StampList.AddItem "INSIDE LINING"
    StampList.AddItem "OUTSIDE LINING"
    StampList.AddItem "FULL LINING"
    StampList.AddItem ""
    StampList.AddItem "REINFORCED ELBOW"
End Sub

Private Sub OptionGeneral_Click()
'Stamps for GENERAL use
    Edit_USER.Enabled = False
    StampList.Clear
    'StampList.AddItem "  "
End Sub

Private Sub OptionLeg_Click()
'Stamps for Legs
    Edit_USER.Enabled = False
    StampList.Clear
    StampList.AddItem "REINFORCED KNEE"
    StampList.AddItem "BEHIND KNEE LINING"
    StampList.AddItem "REINFORCED TOP OF THIGH"
    StampList.AddItem "NO ELASTIC"
    StampList.AddItem "ELASTIC"
End Sub

Private Sub OptionUser_Click()
    Dim sLine

'Stamps created and stored by the USER
    Edit_USER.Enabled = True
    StampList.Clear

'Load USER Stamps from file
    hFileNumber = FreeFile
    If Dir$("C:\JOBST\STAMPS.USR") <> "" Then
        Open "C:\JOBST\STAMPS.USR" For Input As hFileNumber
        Do While Not EOF(hFileNumber)
            Line Input #hFileNumber, sLine
            StampList.AddItem sLine
            Loop
        Close hFileNumber
    End If
End Sub

Private Sub OptionWaistHt_Click()
'Stamps for WAIST HT
    Edit_USER.Enabled = False
    StampList.Clear
    StampList.AddItem "BEHIND KNEE LINING"
    StampList.AddItem "REINFORCED KNEE"
    StampList.AddItem "REINFORCED INNER\nTHIGH PERINEUM"
    StampList.AddItem "REINFORCED TOP OF THIGH"
    StampList.AddItem ""
    StampList.AddItem "ATTACH SUSPENDERS"
    StampList.AddItem "ATTACH VELCRO TABS\n& SEND OTHER PIECES"
    StampList.AddItem "ATTACH VELCRO TABS\nTO VEST & WAIST HEIGHT"
    StampList.AddItem ""
    StampList.AddItem "2in ELASTIC WAIST BAND"
    StampList.AddItem "DOUBLE WAIST BAND"
    StampList.AddItem ""
    StampList.AddItem "NO ELASTIC"
    StampList.AddItem "ELASTIC"
    StampList.AddItem ""
    StampList.AddItem "CUT OUT \nON LEFT ONLY"

End Sub

Private Sub prInsertStamp()
    Dim nNewLineStart As Integer
    Dim fNum As Integer
    Dim sStampLine1 As String, sStampLine2 As String
    Dim sPathJOBST

    sPathJOBST = fnPathJOBST()

    If txtStamp <> "" Then
       'Split into two lines if required
        nNewLineStart = InStr(1, txtStamp.Text, "\n", 0)
        If nNewLineStart > 0 Then
            sStampLine1 = FN_TextCheck(Left$(txtStamp.Text, nNewLineStart - 1))
            sStampLine2 = FN_TextCheck(Mid$(txtStamp.Text, nNewLineStart + 2))
        Else
            sStampLine1 = FN_TextCheck(txtStamp)
            sStampLine2 = ""
        End If

       'Create drawing macro
       'Open file
        fNum = FreeFile
        Open "C:\JOBST\DRAW.D" For Output As fNum
        
       'Write headder Information to file
        Print #fNum, "//DRAFIX Macro created - "; Date$; "  "; Time$
        Print #fNum, "//by Visual Basic STAMP.EXE"

       'Declare and set DRAFIX variables
       'N.B Chr$(34) is 'Double quotes (") character
        Print #fNum, "STRING sStamp1, sStamp2;"
        Print #fNum, "sStamp1="; Chr$(34); sStampLine1; Chr$(34); ";"
        Print #fNum, "sStamp2="; Chr$(34); sStampLine2; Chr$(34); ";"

       'Select macro depending on Arrow check box
       'Note:-
       '    Two different macros are used to speed the insertion of text
       '    during the simple case with no arrows.
        If chkArrow.Value = 1 Then
            Print #fNum, "@" & sPathJOBST & "\stamps\stmparrw.d;"
        Else
            Print #fNum, "@" & sPathJOBST & "\stamps\stmpdraw.d;"
        End If

        Close #fNum
       
       'Activate DRAFIX Windows CAD
        AppActivate fnGetDrafixWindowTitleText()

       'Start drawing macro
        SendKeys "@c:\jobst\draw.d{enter}", True
        
        End
    Else
        MsgBox "You have not yet selected a Stamp", 48, "Stamp Selection Error"
    End If
End Sub

Private Sub StampList_Click()
    txtStamp.Text = StampList.List(StampList.ListIndex)
End Sub

Private Sub StampList_DblClick()
    txtStamp.Text = StampList.List(StampList.ListIndex)
    If txtStamp.Text <> "" Then prInsertStamp
End Sub

