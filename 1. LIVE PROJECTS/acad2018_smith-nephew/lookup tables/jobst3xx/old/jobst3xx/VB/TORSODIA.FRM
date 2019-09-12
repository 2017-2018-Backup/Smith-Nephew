VERSION 5.00
Begin VB.Form torsodia 
   Caption         =   "TORSO Band - Dialogue"
   ClientHeight    =   4455
   ClientLeft      =   3765
   ClientTop       =   2370
   ClientWidth     =   7650
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
   LinkTopic       =   "torsodia"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   7650
   Begin VB.CommandButton Tab 
      Appearance      =   0  'Flat
      Caption         =   "Tab"
      Default         =   -1  'True
      Height          =   375
      Left            =   9480
      TabIndex        =   12
      Top             =   5880
      Width           =   675
   End
   Begin VB.Frame Frame4 
      Caption         =   "Design Choices"
      Height          =   1455
      Left            =   120
      TabIndex        =   49
      Top             =   2880
      Width           =   3735
      Begin VB.ComboBox cboClosure 
         Height          =   315
         Left            =   1260
         TabIndex        =   13
         Text            =   "cboClosure"
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox cboFabric 
         Height          =   315
         Left            =   180
         TabIndex        =   50
         Text            =   "cboFabric"
         Top             =   900
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Closure:"
         Height          =   285
         Index           =   32
         Left            =   240
         TabIndex        =   14
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Fabric:"
         Height          =   285
         Index           =   33
         Left            =   180
         TabIndex        =   51
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label Label1 
         Height          =   285
         Index           =   26
         Left            =   1980
         TabIndex        =   52
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   8880
      Top             =   5880
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   10
      Left            =   8880
      TabIndex        =   53
      Text            =   "Fabric"
      Top             =   5460
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   9
      Left            =   8880
      TabIndex        =   54
      Text            =   "Closure"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   8
      Left            =   8880
      TabIndex        =   55
      Text            =   "Back Neck"
      Top             =   4860
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   7
      Left            =   8880
      TabIndex        =   56
      Text            =   "Front Neck"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   6
      Left            =   8880
      TabIndex        =   57
      Text            =   "Right Axilla"
      Top             =   4260
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   5
      Left            =   8880
      TabIndex        =   15
      Text            =   "Left Axilla"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   8880
      TabIndex        =   16
      Text            =   "Right Bracup"
      Top             =   3660
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   8880
      TabIndex        =   17
      Text            =   "Left Bra Cup"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   8880
      TabIndex        =   18
      Text            =   "Nipple Reduction"
      Top             =   3060
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   8880
      TabIndex        =   19
      Text            =   "EOS reduction"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtCombo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   8880
      TabIndex        =   21
      Text            =   "Waist Reduction"
      Top             =   2460
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidVB 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8880
      TabIndex        =   23
      Text            =   "txtUidVB"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtWorkOrder 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8880
      TabIndex        =   24
      Text            =   "txtWorkOrder"
      Top             =   930
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUidMPD 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8880
      TabIndex        =   26
      Text            =   "txtUidMPD"
      Top             =   630
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Measurements below waist"
      Height          =   1125
      Left            =   120
      TabIndex        =   39
      Top             =   1620
      Width           =   3735
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   8
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   5
         Top             =   660
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   7
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "7"
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Circ. at EOS"
         Height          =   285
         Index           =   18
         Left            =   180
         TabIndex        =   48
         Top             =   720
         Width           =   1635
      End
      Begin VB.Label Label1 
         Height          =   285
         Index           =   17
         Left            =   1980
         TabIndex        =   47
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   8
         Left            =   2940
         TabIndex        =   46
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Chest to EOS"
         Height          =   285
         Index           =   16
         Left            =   180
         TabIndex        =   45
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label1 
         Height          =   285
         Index           =   15
         Left            =   1980
         TabIndex        =   44
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   7
         Left            =   2940
         TabIndex        =   43
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Measurements"
      Height          =   1425
      Left            =   120
      TabIndex        =   38
      Top             =   90
      Width           =   3735
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   6
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   3
         Top             =   960
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   5
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   2
         Top             =   630
         Width           =   555
      End
      Begin VB.TextBox txtCir 
         Height          =   315
         Index           =   4
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   1
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   6
         Left            =   2940
         TabIndex        =   42
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Waist circ."
         Height          =   285
         Index           =   13
         Left            =   180
         TabIndex        =   40
         Top             =   1020
         Width           =   1635
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   5
         Left            =   2940
         TabIndex        =   37
         Top             =   690
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Chest circ."
         Height          =   285
         Index           =   11
         Left            =   180
         TabIndex        =   34
         Top             =   690
         Width           =   1635
      End
      Begin VB.Label lblCir 
         Caption         =   "inches"
         Height          =   165
         Index           =   4
         Left            =   2940
         TabIndex        =   32
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Chest to Waist"
         Height          =   285
         Index           =   9
         Left            =   180
         TabIndex        =   28
         Top             =   360
         Width           =   1635
      End
   End
   Begin VB.CommandButton OK 
      Caption         =   "Draw"
      Height          =   375
      Left            =   4380
      TabIndex        =   20
      Top             =   3660
      Width           =   1275
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5940
      TabIndex        =   22
      Top             =   3660
      Width           =   1275
   End
   Begin VB.Frame Frame5 
      Caption         =   "Patient Details"
      Height          =   2055
      Left            =   4020
      TabIndex        =   0
      Top             =   90
      Width           =   3495
      Begin VB.TextBox txtDiagnosis 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   31
         Top             =   1140
         Width           =   2535
      End
      Begin VB.TextBox txtUnits 
         Height          =   315
         Left            =   2580
         TabIndex        =   27
         Top             =   300
         Width           =   795
      End
      Begin VB.TextBox txtSex 
         Height          =   315
         Left            =   2520
         TabIndex        =   35
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtAge 
         Height          =   315
         Left            =   840
         TabIndex        =   33
         Top             =   1560
         Width           =   525
      End
      Begin VB.TextBox txtFileNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   25
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtPatientName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   29
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label22 
         Caption         =   "Diag:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Units:"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Sex:"
         Height          =   195
         Index           =   30
         Left            =   2040
         TabIndex        =   8
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "File No:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Age:"
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   1620
         Width           =   375
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Combo Box DDE "
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   36
      Left            =   8880
      TabIndex        =   30
      Top             =   2250
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Vest Body"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   35
      Left            =   9000
      TabIndex        =   36
      Top             =   1500
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Patient Details"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   34
      Left            =   9000
      TabIndex        =   41
      Top             =   420
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "torsodia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   TORSODIA.MAK
'Purpose:   TORSO Band Dialogue
'
'
'Version:   3.01
'Date:      13th Jan 1998
'Author:    Gary George
'Copyright  C-Gem Ltd
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'Dec 98     GG      Ported to VB5
'
'Notes:-
'
'
'
'
'   'Windows API Functions Declarations
'    Private Declare Function GetActiveWindow Lib "User" () As Integer
'    Private Declare Function IsWindow Lib "User" (ByVal hwnd As Integer) As Integer
'    Private Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
'    Private Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
'    Private Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
'    Private Declare Function GetNumTasks Lib "Kernel" () As Integer
'    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
'    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
   
   
'   'Constanst used by GetWindow
'    Const GW_CHILD = 5
'    Const GW_HWNDFIRST = 0
'    Const GW_HWNDLAST = 1
'    Const GW_HWNDNEXT = 2
'    Const GW_HWNDPREV = 3
'    Const GW_OWNER = 4

   'MsgBox constant
'    Const IDCANCEL = 2
'    Const IDYES = 6
'    Const IDNO = 7


    
    Option Explicit



Private Sub Cancel_Click()
    Dim Response As Integer, sTask$, sCurrentValues$
    
   'Check if data has been modified
    sCurrentValues = FN_ValuesString()

    If sCurrentValues <> g_sChangeChecker Then
        Response = MsgBox("Changes have been made, Save changes before closing", 35, "CAD - Glove Dialogue")
        Select Case Response
            Case IDYES
                PR_CreateMacro_Save "c:\jobst\draw.d"
                sTask = fnGetDrafixWindowTitleText()
                If sTask <> "" Then
                    AppActivate sTask
                    SendKeys "@c:\jobst\draw.d{enter}", True
                    End
                Else
                    MsgBox "Can't find a Drafix Drawing to update!", 16, "TORSO Band - Dialogue"
                End If
            Case IDNO
                End
            Case IDCANCEL
                Exit Sub
        End Select
    Else
        End
    End If
End Sub

Private Function FN_EscapeQuotesInString(sAssignedString As Variant) As String
   'Search through the string looking for " (double quote characater)
   'If found use \ (Backslash) to escape it
   '
    Dim ii%
    Dim Char$
    Dim sEscapedString$
    
    FN_EscapeQuotesInString = ""
    
    For ii = 1 To Len(sAssignedString)
        Char = Mid$(sAssignedString, ii, 1)
        If Char = QQ Then
            sEscapedString = sEscapedString & "\" & Char
        Else
            sEscapedString = sEscapedString & Char
        End If
    Next ii
    
    FN_EscapeQuotesInString = sEscapedString

End Function

Private Function FN_OpenSave(sDrafixFile As String, sType As String, sName As Variant, sFileNo As Variant) As Integer
   'Open the DRAFIX macro file
   'Return the file number

   'Open file
    fNum = FreeFile
    Open sDrafixFile For Output As fNum
    FN_OpenSave = fNum

   'Write header information etc. to the DRAFIX macro file
   '
    Print #fNum, "//DRAFIX Macro created - "; Date$; "  "; Time$
    Print #fNum, "//Patient - "; sName; ", "; sFileNo; ""
    Print #fNum, "//by Visual Basic, TORSO Band"
    Print #fNum, "//type - "; sType; ""


End Function



Private Function FN_ValidateData() As Integer
   'This function is used only to make gross checks
   'for missing data.
   'It does not perform any sensibility checks on the
   'data
    Dim sError$, ii%, nn%
    
   'Initialise
    FN_ValidateData = False
    sError = ""

    ReDim sCircum(0 To 8)
'    sCircum(0) = "Left shoulder circ."
'    sCircum(1) = "Right shoulder circ."
'    sCircum(2) = "Neck circ."
'    sCircum(3) = "Shoulder width"
    sCircum(4) = "Chest to waist"
    sCircum(5) = "Chest circ."
    sCircum(6) = "Waist circ."
    sCircum(7) = "Chest to EOS"
    sCircum(8) = "Circ. at EOS"

   'Vest measurements
'    For ii = 0 To 1
'        If Val(txtCir(ii).Text) = 0 Then
'            sError = sError & "Missing dimension for " & sCircum(ii) & "!" & NL
'        End If
'    Next ii

    For ii = 4 To 6
        If Val(txtCir(ii).Text) = 0 Then
            sError = sError & "Missing dimension for " & sCircum(ii) & "!" & NL
        End If
    Next ii

   'EOS Measurements (if one given both must be given)
    If Val(txtCir(7).Text) = 0 And Val(txtCir(8).Text) <> 0 Then
            sError = sError & "Missing dimension for " & sCircum(7) & "!" & NL
    End If
    If Val(txtCir(8).Text) = 0 And Val(txtCir(8).Text) <> 0 Then
            sError = sError & "Missing dimension for " & sCircum(8) & "!" & NL
    End If

    If cboClosure.Text = "" Then
            sError = sError & "Closure not given! " & NL
    End If

    If cboFabric.Text = "" Then
            sError = sError & "Fabric not given! " & NL
    End If
    
    If sError <> "" Then
        MsgBox sError, 16, "TORSO Band - Dialogue"
        FN_ValidateData = False
    Else
        FN_ValidateData = True
    End If

End Function

Private Function FN_ValuesString() As String
   'Create a string of all the values given in the
   'Text and Combo boxes.
   '
   'Ignore patient details
   '
    Dim ii%, sString$
    
    FN_ValuesString = ""
    sString = ""

    For ii = 4 To 8
        sString = sString & txtCir(ii)
    Next ii

    sString = sString & cboClosure.Text
    
    sString = sString & cboFabric.Text

    FN_ValuesString = sString

End Function


Private Sub Form_LinkClose()
    Dim ii%

   'Stop the timer used to ensure that the Dialogue dies
   'if the DRAFIX macro fails to establish a DDE Link
    Timer1.Enabled = False

   'Check that a "MainPatientDetails" Symbol has been
   'found
    If txtUidMPD = "" Then
        MsgBox "No Patient Details have been found in drawing!", 16, "Error, VEST Body - Dialogue"
        End
    End If

    If txtCombo(9) <> "" Then cboClosure.Text = txtCombo(9) Else cboClosure.ListIndex = 0
    cboFabric.Text = txtCombo(10)

   'Set up units
    If txtUnits = "cm" Then
        g_nUnitsFac = 10 / 25.4
    Else
        g_nUnitsFac = 1
    End If

   'Display dimesions sizes in inches
    For ii = 4 To 8
        txtCir_LostFocus (ii)
    Next ii


   'Save the values in the text to a string
   'this can then be used to check if they have changed
   'on use of the close button
    g_sChangeChecker = FN_ValuesString()
    Screen.MousePointer = 0    ' Change pointer to hourglass.
    
    Show
End Sub

Private Sub Form_Load()
    Dim ii%

    Hide

   'Start a timer
   'The Timer is disabled in LinkClose
   'If after 6 seconds the timer event will "End" the programme
   'This ensures that the dialogue dies in event of a failure
   'on the drafix macro side
    Timer1.Interval = 6000  'Approx 6 Seconds
    Timer1.Enabled = True

   'Check if a previous instance is running
   'If it is warn user and exit
    If App.PrevInstance Then
        MsgBox "TORSO Band Dialogue is already running!" + Chr(13) + "Use ALT-TAB and Cancel it.", 16, "Error Starting VEST Body - Dialogue"
        End
    End If

   'Maintain while loading DDE data
    Screen.MousePointer = 11    ' Change pointer to hourglass.
   'Reset in Form_LinkClose
    
   'Position to center of screen
    Left = (Screen.Width - torsodia.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - torsodia.Height) / 2  ' Center form vertically.

    Set MainForm = torsodia

   'Initialize globals
    QQ = Chr$(34) 'Double quotes (")
    NL = Chr$(13) 'New Line
    CC = Chr$(44) 'The comma (,)
    QCQ = QQ & CC & QQ
    QC = QQ & CC
    CQ = CC & QQ

    g_nUnitsFac = 1
    g_PathJOBST = fnPathJOBST()

   'Clear fields
   'Circumferences and lengths
    For ii = 4 To 8
        txtCir(ii) = ""
    Next ii

   'The data from these DDE text boxes is copied
   'to the combo boxes on Link close
   '
    txtCombo(0).Text = ""
    txtCombo(1).Text = ""
    txtCombo(9).Text = ""
    txtCombo(10).Text = ""
    
   'closure
'    cboClosure.AddItem "Velcro"
'    cboClosure.AddItem "Zip"
    cboClosure.AddItem "Front Velcro"
    cboClosure.AddItem "Front Velcro (Reversed)"
    cboClosure.AddItem "Back Velcro"
    cboClosure.AddItem "Back Velcro (Reversed)"
    cboClosure.AddItem "Front Zip"
    cboClosure.AddItem "Back Zip"
   

   'Patient details
    txtFileNo = ""
    txtUnits = ""
    txtPatientName = ""
    txtDiagnosis = ""
    txtAge = ""
    txtSex = ""
    txtWorkOrder = ""


   'UID of symbols
    txtUidMPD = ""
    txtUidVB = ""

    PR_GetComboListFromFile cboFabric, g_PathJOBST & "\FABRIC.DAT"


End Sub

Private Sub OK_Click()
    Dim sTask$
   'Don't allow multiple clicking
   '
    OK.Enabled = False
    If FN_ValidateData() Then
        Screen.MousePointer = 11
        Hide

        PR_CreateMacro_Draw "c:\jobst\draw.d"
        
        sTask = fnGetDrafixWindowTitleText()
        If sTask <> "" Then
            AppActivate sTask
              SendKeys "@c:\jobst\draw.d{enter}", True
            Screen.MousePointer = 0
            End
        Else
            MsgBox "Can't find a Drafix Drawing to update!", 16, "TORSO Band - Dialogue"
        End If
    End If
    OK.Enabled = True
    Screen.MousePointer = 0

End Sub

Private Sub PR_CreateMacro_Draw(sFile As String)
   'Since this is a fairly straight forward module we
   'will keep calculation and drawing in the same procedure
   '
   'The assumption is that all data has been validated before
   'this procedure is called

   'Figured dimemsions
    Dim nEOSCir             As Double
    Dim nEOStoShoulder      As Double       'Misnomer due to mods 04.Jun.98

    Dim nWaistCir           As Double
    Dim nWaisttoShoulder    As Double       'Misnomer due to mods 04.Jun.98

    Dim nChestCir           As Double
'    Dim nRightShoulderCir   As Double
'    Dim nLeftShoulderCir    As Double
'    Dim nChesttoShoulder    As Double

    Dim nLowShoulderLine    As Double

   'Points
    Dim xyO         As xy
    Dim xyEOS       As xy
    Dim xyEOSatCL   As xy
    Dim xyChest     As xy
    Dim xyChestatCL As xy
    Dim xyWaist     As xy
    Dim xyWaistatCL As xy

   'Figuring factors
    Dim nCirFactor          As Double
'    Dim nShoulderCirFactor  As Double
    Dim nSeamAllowance      As Double

   'Flags
    Dim bEOSGiven   As Integer

   'other
    Dim sText       As String
    Dim sWorkOrder  As String

    Dim nZiplength  As Double

    Dim xyText      As xy

   'Set figuring factors, flags, defaults etc...
    bEOSGiven = False
    nCirFactor = 0.9
'    nShoulderCirFactor = 2.5
    nSeamAllowance = 0.1875
    PR_MakeXY xyO, 0, 0

   'Figure given dimensions
   'N.B.
   'txtCir(0)  = Left Shoulder cir.     (not required)
   'txtCir(1)  = Right Shoulder cir.    (not required)
   'txtCir(2)  = Neck cir.              (not required)
   'txtCir(3)  = Shoulder width         (not required)
   'txtCir(4)  = Chest to waist
   'txtCir(5)  = Chest cir.
   'txtCir(6)  = Waist cir.
   'txtCir(7)  = Chest to EOS           (optional, but must be given if EOS Cir. below given
   'txtCir(8)  = EOS cir.               (optional)

   'EOS (This is optional)
    If Val(txtCir(8)) > 0 Then
        bEOSGiven = True
        nEOSCir = fnDisplayToInches(txtCir(8)) * nCirFactor / 4
        nEOStoShoulder = fnDisplayToInches(txtCir(7))
    End If

   'Waist
    nWaistCir = fnDisplayToInches(txtCir(6)) * nCirFactor / 4
    nWaisttoShoulder = fnDisplayToInches(txtCir(4))

   'Chest
    nChestCir = fnDisplayToInches(txtCir(5)) * nCirFactor / 4

'    nRightShoulderCir = fnDisplayToInches(txtCir(1))
'    nLeftShoulderCir = fnDisplayToInches(txtCir(0))
'
'    If Abs(nRightShoulderCir - nLeftShoulderCir) > 1 Then
'       'use the highest
'        nRightShoulderCir = round(nRightShoulderCir / nShoulderCirFactor)
'        nLeftShoulderCir = round(nLeftShoulderCir / nShoulderCirFactor)
'        nChesttoShoulder = min(nRightShoulderCir, nLeftShoulderCir)
'    Else
'       'use the average
'        nChesttoShoulder = round(((nRightShoulderCir + nLeftShoulderCir) / 2) / nShoulderCirFactor)
'    End If
'   'Drop "nChestToShoulder" by 1/2" to match changes requested to the vest
'   'on 16.Oct.97
'    nChesttoShoulder = nChesttoShoulder - .5

   'High shoulder line
    If bEOSGiven Then
        nLowShoulderLine = nEOStoShoulder
    Else
        nLowShoulderLine = nWaisttoShoulder
    End If

   'Key Points
   'End of Support (If given) and Waist points.
    If bEOSGiven Then
        xyEOSatCL.X = xyO.X
        xyEOSatCL.Y = xyO.Y
        xyEOS.X = xyO.X
        xyEOS.Y = nEOSCir + nSeamAllowance + xyO.Y
        xyWaist.X = nLowShoulderLine - nWaisttoShoulder + xyO.X
        xyWaist.Y = nWaistCir + nSeamAllowance + xyO.Y
        xyWaistatCL.X = xyWaist.X
        xyWaistatCL.Y = xyO.Y
    Else
        xyWaistatCL.X = xyO.X
        xyWaistatCL.Y = xyO.Y
        xyWaist.X = xyO.X
        xyWaist.Y = nWaistCir + nSeamAllowance + xyO.Y
    End If

   'Chest points
'    xyChestatCL.X = nLowShoulderLine - nChesttoShoulder + xyO.X
    xyChestatCL.X = nLowShoulderLine + xyO.X
    xyChestatCL.Y = xyO.Y
    xyChest.X = xyChestatCL.X
    xyChest.Y = nChestCir + nSeamAllowance + xyO.Y

   'Zippers
    If bEOSGiven Then
        nZiplength = FN_CalcLength(xyEOSatCL, xyChestatCL)
    Else
        nZiplength = FN_CalcLength(xyChestatCL, xyWaistatCL)
    End If
    nZiplength = (nZiplength - 0.125) / 0.95

   'DRAW DRAW DRAW DRAW DRAW DRAW DRAW DRAW
   '(Join the dots etc..)
    
    fNum = FN_Open(sFile)

   'Draw Text items
    PR_Setlayer "Notes"
   
   'Zippers
    PR_SetTextData 2, 32, -1, -1, -1    'Horiz-Cen, Vertical-Bottom
    sText = fnInchesToText(nZiplength) & "\"" " & cboClosure.Text
    PR_CalcMidPoint xyChestatCL, xyWaistatCL, xyText
    PR_MakeXY xyText, xyText.X, xyText.Y + 0.25
    PR_DrawText sText, xyText, 0.1

   'Elastic Top
    g_nCurrTextAngle = 270
    PR_CalcMidPoint xyChestatCL, xyChest, xyText
    PR_MakeXY xyText, xyText.X - 0.25, xyText.Y
    PR_DrawText "TOP ELASTIC", xyText, 0.1

   'Elastic Bottom
    g_nCurrTextAngle = 90
    If bEOSGiven Then
        PR_CalcMidPoint xyEOSatCL, xyEOS, xyText
    Else
        PR_CalcMidPoint xyWaistatCL, xyWaist, xyText
    End If
    PR_MakeXY xyText, xyText.X + 0.25, xyText.Y
    PR_DrawText "BOTTOM ELASTIC", xyText, 0.1

   'Patient details
    g_nCurrTextAngle = 0
    PR_SetTextData 1, 32, -1, -1, -1    'Horiz-Left, Vertical-Bottom
    PR_CalcMidPoint xyChest, xyWaistatCL, xyText
    If txtWorkOrder.Text = "" Then
        sWorkOrder = "-"
    Else
        sWorkOrder = txtWorkOrder.Text
    End If
    sText = txtPatientName.Text & "\n" & sWorkOrder & "\n" & Trim$(Mid$(cboFabric.Text, 4))
    PR_DrawText sText, xyText, 0.1
    
   'Remaining patient details in black on layer construct
    PR_Setlayer "Construct"
    PR_MakeXY xyText, xyText.X, xyText.Y - 0.8
    sText = txtFileNo.Text & "\n" & txtDiagnosis.Text & "\n" & txtAge.Text & "\n" & txtSex.Text
    PR_DrawText sText, xyText, 0.1
    
   'Draw construction lines
    If bEOSGiven Then PR_DrawLine xyWaistatCL, xyWaist
    
   'Draw profile
    PR_Setlayer "TemplateLeft"
    PR_StartPoly

     If bEOSGiven Then
        PR_AddVertex xyEOSatCL, 0
        PR_AddVertex xyEOS, 0
        PR_AddVertex xyWaist, 0
        PR_AddVertex xyChest, 0
        PR_AddVertex xyChestatCL, 0
        PR_AddVertex xyEOSatCL, 0
     Else
        PR_AddVertex xyWaistatCL, 0
        PR_AddVertex xyWaist, 0
        PR_AddVertex xyChest, 0
        PR_AddVertex xyChestatCL, 0
        PR_AddVertex xyWaistatCL, 0
     End If
    PR_EndPoly
    
    PR_UpdateDB

    Close #fNum

End Sub

Private Sub PR_CreateMacro_Save(sDrafixFile As String)
   'fNum is a global variable use in subsequent procedures
    fNum = FN_OpenSave(sDrafixFile, "Save Data ONLY", (txtPatientName.Text), (txtFileNo.Text))

   'If this is a new drawing of a vest then Define the DATA Base
   'fields for the VEST Body and insert the BODYBOX symbol
    PR_PutLine "HANDLE hMPD, hBody;"
   
    PR_UpdateDB

    Close #fNum
   
End Sub


Private Sub PR_GetComboListFromFile(Combo_Name As Control, sFileName As String)
   'General procedure to create the list section of
   'a combo box reading the data from a file
    
    Dim sLine As String, fFileNum As Integer

    fFileNum = FreeFile
    
    If FileLen(sFileName) = 0 Then
        MsgBox sFileName & "Not found", 48, "CAD - Glove Dialogue"
        Exit Sub
    End If

    Open sFileName For Input As fFileNum
    Do While Not EOF(fFileNum)
        Line Input #fFileNum, sLine
        Combo_Name.AddItem sLine
        Loop
    Close #fFileNum

End Sub

Private Sub PR_UpdateDB()
   'Procedure called from
   '    PR_CreateMacro_Save
   'and
   '    PR_CreateMacro_Draw
   '
   'Used to stop duplication on code

    Dim sSymbol As String
    
    sSymbol = "vestbody"

    If txtUidVB.Text = "" Then
       'Define DB Fields
        PR_PutLine "@" & g_PathJOBST & "\VEST\VFIELDS.D;"

       'Find "mainpatientdetails" and get position
        PR_PutLine "XY     xyMPD_Origin, xyMPD_Scale ;"
        PR_PutLine "STRING sMPD_Name;"
        PR_PutLine "ANGLE  aMPD_Angle;"

        PR_PutLine "hMPD = UID (" & QQ & "find" & QC & Val(txtUidMPD.Text) & ");"
        PR_PutLine "if (hMPD)"
        PR_PutLine "  GetGeometry(hMPD, &sMPD_Name, &xyMPD_Origin, &xyMPD_Scale, &aMPD_Angle);"
        PR_PutLine "else"
        PR_PutLine "  Exit(%cancel," & QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & QQ & ");"

       'Insert bodybox
        PR_PutLine "if ( Symbol(" & QQ & "find" & QCQ & sSymbol & QQ & ")){"
        PR_PutLine "  Execute (" & QQ & "menu" & QCQ & "SetLayer" & QC & "Table(" & QQ & "find" & QCQ & "layer" & QCQ & "Data" & QQ & "));"
        PR_PutLine "  hBody = AddEntity(" & QQ & "symbol" & QCQ & sSymbol & QC & "xyMPD_Origin);"
        PR_PutLine "  }"
        PR_PutLine "else"
        PR_PutLine "  Exit(%cancel, " & QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & QQ & ");"
    Else
       'Use existing symbol
        PR_PutLine "hBody = UID (" & QQ & "find" & QC & Val(txtUidVB.Text) & ");"
        PR_PutLine "if (!hBody) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");"

    End If

   'Update the BODY Box symbol
'    PR_PutLine "SetDBData( hBody" & CQ & "LtSCir" & QCQ & txtCir(0).Text & QQ & ");"
'    PR_PutLine "SetDBData( hBody" & CQ & "RtSCir" & QCQ & txtCir(1).Text & QQ & ");"
'    PR_PutLine "SetDBData( hBody" & CQ & "NeckCir" & QCQ & txtCir(2).Text & QQ & ");"
'    PR_PutLine "SetDBData( hBody" & CQ & "SWidth" & QCQ & txtCir(3).Text & QQ & ");"
'    PR_PutLine "SetDBData( hBody" & CQ & "S_Waist" & QCQ & txtCir(4).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "SLgButt" & QCQ & txtCir(4).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "ChestCir" & QCQ & txtCir(5).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "WaistCir" & QCQ & txtCir(6).Text & QQ & ");"
    
'    PR_PutLine "SetDBData( hBody" & CQ & "S_EOS" & QCQ & txtCir(7).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "SFButt" & QCQ & txtCir(7).Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "EOSCir" & QCQ & txtCir(8).Text & QQ & ");"
    
    PR_PutLine "SetDBData( hBody" & CQ & "Closure" & QCQ & cboClosure.Text & QQ & ");"
    PR_PutLine "SetDBData( hBody" & CQ & "Fabric" & QCQ & cboFabric.Text & QQ & ");"
    
    
    PR_PutLine "SetDBData( hBody" & CQ & "fileno" & QCQ & txtFileNo.Text & QQ & ");"

End Sub

Private Sub Tab_Click()
   'Allows the user to use enter as a tab
    
    SendKeys "{TAB}", False

End Sub

Private Sub Timer1_Timer()
    'It is assumed that the link open from Drafix has failed
    'Therefor we "End" here
     End
End Sub

Private Sub txtCir_GotFocus(Index As Integer)
    PR_Select_Text txtCir(Index)
End Sub

Private Sub txtCir_LostFocus(Index As Integer)
    Dim nLen#
    
    nLen = FN_InchesValue(txtCir(Index))
    If nLen > 0 Then
        lblCir(Index).Caption = fnInchesToText(nLen)
    Else
        lblCir(Index).Caption = ""
    End If
    
End Sub

