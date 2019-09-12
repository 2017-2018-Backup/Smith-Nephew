VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form cltxfdia 
   Caption         =   "Patient and Work Order Details"
   ClientHeight    =   6876
   ClientLeft      =   696
   ClientTop       =   1992
   ClientWidth     =   6432
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
   LinkTopic       =   "cltxfdia"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6876
   ScaleWidth      =   6432
   Begin VB.Frame frmPatientDetails 
      Caption         =   "Patient Details"
      Height          =   1575
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   6210
      Begin VB.TextBox txtFileNoPD 
         Height          =   285
         Left            =   1140
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtAgePD 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4980
         TabIndex        =   37
         Top             =   645
         Width           =   435
      End
      Begin VB.ComboBox cboSexPD 
         Height          =   315
         Left            =   4980
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1050
         Width           =   975
      End
      Begin VB.TextBox txtNamePD 
         Height          =   285
         Left            =   1140
         TabIndex        =   35
         Top             =   645
         Width           =   3195
      End
      Begin VB.ComboBox cboDiagnosisPD 
         Height          =   315
         Left            =   1140
         TabIndex        =   34
         Text            =   "cboDiagnosisPD"
         Top             =   1050
         Width           =   3195
      End
      Begin VB.Label labPatientDetails 
         Caption         =   "File No:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   43
         Top             =   285
         Width           =   675
      End
      Begin VB.Label labPatientDetails 
         Caption         =   "Name:"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   42
         Top             =   690
         Width           =   615
      End
      Begin VB.Label labPatientDetails 
         Caption         =   "Diagnosis:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   41
         Top             =   1095
         Width           =   915
      End
      Begin VB.Label labPatientDetails 
         Caption         =   "Age:"
         Height          =   195
         Index           =   3
         Left            =   4500
         TabIndex        =   40
         Top             =   690
         Width           =   495
      End
      Begin VB.Label labPatientDetails 
         Caption         =   "Sex:"
         Height          =   195
         Index           =   4
         Left            =   4500
         TabIndex        =   39
         Top             =   1095
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   5100
      Top             =   6300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      DialogTitle     =   "Open CAD File"
      FontSize        =   0
      MaxFileSize     =   256
      PrinterDefault  =   0   'False
   End
   Begin VB.TextBox txtInitialCADFile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   32
      Text            =   "txtInitialCADFile"
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCurrentCADFile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   31
      Text            =   "txtCurrentCADFile"
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame frmCADFile 
      Caption         =   "Drafix CAD File"
      Height          =   825
      Left            =   120
      TabIndex        =   28
      Top             =   3210
      Width           =   6195
      Begin VB.CommandButton CAD_File 
         Cancel          =   -1  'True
         Caption         =   "CAD File ..."
         Height          =   375
         Left            =   4740
         TabIndex        =   30
         Top             =   330
         Width           =   1275
      End
      Begin VB.Label lblCADFile 
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.TextBox txtInvokedFrom 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   0
      Text            =   "txtInvokedFrom"
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   6240
      Width           =   1275
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   6240
      Width           =   1275
   End
   Begin VB.Frame frmWODetails 
      Caption         =   "Work Order Details"
      Height          =   1320
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   6210
      Begin VB.ComboBox cboDesignerWOD 
         Height          =   315
         Left            =   1140
         TabIndex        =   14
         Top             =   720
         Width           =   1995
      End
      Begin VB.TextBox txtOrderDateWOD 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3100
         MaxLength       =   10
         TabIndex        =   12
         Top             =   315
         Width           =   1200
      End
      Begin VB.ComboBox cboUnitsWOD 
         Height          =   315
         Left            =   4980
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   315
         Width           =   975
      End
      Begin VB.TextBox txtWorkOrderWOD 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1140
         TabIndex        =   11
         Top             =   315
         Width           =   855
      End
      Begin VB.Label labWODetails 
         Caption         =   "Units:"
         Height          =   240
         Index           =   3
         Left            =   4380
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
      Begin VB.Label labWODetails 
         Caption         =   "Temp Date:"
         Height          =   240
         Index           =   2
         Left            =   2100
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label labWODetails 
         Caption         =   "Designer:"
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   765
         Width           =   855
      End
      Begin VB.Label labWODetails 
         Caption         =   "WorkOrder:"
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox txtMessage 
      Height          =   1680
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   4380
      Width           =   6195
   End
   Begin VB.Timer Timer1 
      Left            =   7080
      Top             =   6390
   End
   Begin VB.TextBox txtWO_TXF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   23
      Text            =   "txtWO_TXF"
      Top             =   2055
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtWorkOrder 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   6
      Text            =   "txtWorkOrder"
      Top             =   1695
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtUidMPD 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   17
      Text            =   "txtUidMPD"
      Top             =   6015
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtTemplateEngineer 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   16
      Text            =   "txtTemplateEngineer"
      Top             =   5655
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtOrderDate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   15
      Text            =   "txtOrderDate"
      Top             =   5295
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtMPDwo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   9
      Text            =   "txtMPDwo"
      Top             =   4935
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   8
      Text            =   "txtAge"
      Top             =   3495
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtSex 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   7
      Text            =   "txtSex"
      Top             =   4215
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtDiagnosis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   5
      Text            =   "txtDiagnosis"
      Top             =   4575
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtPatientName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   4
      Text            =   "txtPatientName"
      Top             =   3135
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtUnits 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   3
      Text            =   "txtUnits"
      Top             =   3855
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtFileNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6660
      TabIndex        =   1
      Text            =   "txtFileNo"
      Top             =   2775
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Messages and Errors"
      Height          =   240
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "imageABLE"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   6840
      TabIndex        =   24
      Top             =   1425
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PatientDetails"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   6840
      TabIndex        =   18
      Top             =   2505
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "cltxfdia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   CLTXFDIA.MAK
'Purpose:   Takes data from the Triton eXchange File (TXF).
'           Groups the Work Order details and updates the
'           "Patient Details symbol" and "common symbols" with
'           work order information
'
'Version:   3.02
'Date:      11.Jun.95
'Author:    Gary George
'---------------------------------------------------------------------------------------------
'REVISIONS:
'Date       By      Action
'---------------------------------------------------------------------------------------------
'Dec 98     GG      Ported to VB5
'
'NOTE:-
'

    Option Explicit

Private Sub CAD_File_Click()
   'Find an existing CAD file to update
   '
   'Path to drawings
    Dim lpBuffer As String * 144 'Minimum recommended wrt GetWindowsDirectory()
    Dim nBufferSize As Long
    Dim nSize As Long
    Dim WindowsDir$

    nBufferSize = 143

   'Get the path to the Windows Directory to locate DRAFIX.INI
   '
    nSize = GetWindowsDirectory(lpBuffer, nBufferSize)
    WindowsDir$ = Left$(lpBuffer, nSize)

   'Get the path to the Drawing directory from
   'DRAFIX.INI
   '
    nSize = GetPrivateProfileString("Path", "PathDrawing", "C:\", lpBuffer, nBufferSize, WindowsDir$ & "\DRAFIX.INI")

    CMDialog1.InitDir = Left$(lpBuffer, nSize)
    CMDialog1.Filter = "Drafix CAD (*.cad)|*.cad"
    CMDialog1.filename = ""
    CMDialog1.DefaultExt = "cad"
    CMDialog1.Action = 1
    
    If CMDialog1.filename <> "" Then
        If Dir$(CMDialog1.filename) = "" Then
            PR_Message "CAD File not found." & Chr$(13) & Chr$(10)
            PR_Message "   File: " & CMDialog1.filename & Chr$(13) & Chr$(10)
            MsgBox "CAD File not found.", 48, "Patient Details"
        Else
            lblCADFile.Caption = CMDialog1.filename
        End If
    End If
End Sub

Private Sub Cancel_Click()
    End
End Sub

Private Sub cboUnitsWOD_Click()
    g_sUnits = cboUnitsWOD.List(cboUnitsWOD.ListIndex)
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


Private Function FN_ValidatePDandWOD() As Integer
   'Validates the
   '
   '    Patient Details (PD)
   '    Work Order Details (WOD)
   '
   'Displays a Pop Up Message box and Updates txtMessage
   '
   'NOTES
   '    The PD and WOD can be considered as the working
   '    areas.  The details in this area can come from 3
   '    possible sources:-
   '
   '    1.  The Triton eXchange File (TXF)
   '    2.  The existing drawing     (Dwg)
   '    3.  User input
   '
   '    If the details are correct then the DDE text
   '    boxes can be updated with the valid data
   '
    Dim sMessage$, iMsg%
    Const IDOK = 1
    Const IDCANCEL = 2
    Const IDYES = 6
    Const IDNO = 7

    NL = Chr$(13) & Chr$(10) 'Carriage Return & New Line
    TABBED = Chr$(9)             'Tab

    txtMessage.Text = ""
    
   'Patient Details
   '
    If txtFileNoPD = "" Then PR_Message "Missing File Number."
    
    If txtNamePD = "" Then PR_Message "Missing Patient name."
    
    If Val(txtAgePD) = 0 Then PR_Message "Invalid Patient Age."
  
    If cboSexPD.ListIndex = -1 Then PR_Message "Missing Patient SEX."
    
    If cboDiagnosisPD.Text = "" Then PR_Message "Missing Diagnosis."


   'Work Order Details
   '
    If txtWorkOrderWOD = "" Then PR_Message "Missing Work Order Number."
    
    If txtOrderDateWOD = "" Then PR_Message "Missing Date."
    
    If cboDesignerWOD.Text = "" Then PR_Message "Missing Template Designer."
    
    If cboUnitsWOD.ListIndex = -1 Then PR_Message "Missing Units."

    If cboUnitsWOD.Text <> txtUnits And txtUnits <> "" Then PR_Message "WARNING! Inconsistent units.  The drawing uses " & txtUnits & ", you have selected " & cboUnitsWOD.Text & ".  This may cause problems when re-drawing template patterns."

   'Check Patient, Template Engineer and Diagnosis for double quotes
   'Chr$(34) 'Double quotes (")this is to avoid problems later on
   'in drafix macros.
   'Greatfull thanks to Kathleen Ryan of JOBST for highlighting this!
   '
    If InStr(1, cboDesignerWOD.Text, Chr$(34)) Then
        MsgBox "Designer name includes Double Quotes (""). This is not allowed!", 16
        FN_ValidatePDandWOD = False
        PR_Message "Designer name includes Double Quotes."
        Exit Function
    End If
    If InStr(1, txtNamePD.Text, Chr$(34)) Then
        MsgBox "Patient name includes Double Quotes (""). This is not allowed!", 16
        FN_ValidatePDandWOD = False
        PR_Message "Patient name includes Double Quotes."
        Exit Function
    End If
    If InStr(1, cboDiagnosisPD.Text, Chr$(34)) Then
        MsgBox "Diagnosis includes Double Quotes (""). This is not allowed!", 16
        FN_ValidatePDandWOD = False
        PR_Message "Diagnosis includes Double Quotes."
        Exit Function
    End If

    If Len(txtMessage.Text) > 0 Then
        PR_Message ""
        sMessage = txtMessage.Text + "Do you wish to Continue?"
        iMsg = MsgBox(sMessage, 36, "Incomplete Patient and Order Details")
        If iMsg = IDYES Then FN_ValidatePDandWOD = 1
        If iMsg = IDNO Then FN_ValidatePDandWOD = 0
    Else
        FN_ValidatePDandWOD = 1
    End If

End Function

Private Sub Form_LinkClose()
   'Disable TIMEOUT Timer
    Timer1.Enabled = False
   
   'Based on the data in the txtWorkOrder and the
   'txtWO_TXF field
   '
   '
    If txtInvokedFrom = "imageABLE" Then
       'If data is contained in a TXF file
       'we read the data from the TXF file
        txtWorkOrderWOD = txtWorkOrder
        PR_GetTXFData (txtWO_TXF.Text)

       'If not a new Drawing then Compare Patient data
       'from the existing drawing with TXF file data
        PR_TXFtoDwgComparison

       'Display the file that the drawing is based on
       'Modify caption
        CAD_File.Enabled = False
        If txtInitialCADFile <> "INSPECTION" Then
            frmCADFile.Caption = "Previous Drafix CAD File"
            lblCADFile.Caption = txtInitialCADFile
            If txtInitialCADFile = "NEW" Then CAD_File.Enabled = True
        Else
           'Ensure that Designer is retained on inspection
            frmCADFile.Caption = "Drafix CAD File"
            lblCADFile.Caption = txtCurrentCADFile
            cboDesignerWOD.Text = txtTemplateEngineer.Text
           'Disable the OK button to stop inspectors from
           'being able to save as this is unessesary in the inspection
           'stage
            OK.Enabled = False
        End If

    Else
       'Get the data from the poked DDE text boxes
        PR_GetDrawingData
        frmCADFile.Caption = "Drafix CAD File"
        lblCADFile.Caption = txtCurrentCADFile
        CAD_File.Enabled = False
    End If
    
    Show

End Sub

Private Sub Form_Load()
   'Hide form to allow DDE to happen
   'Show form on LinkClose
    Hide
   
   'The VB programme is initiated either through
   'DRAFIX or imageABLE
    txtInvokedFrom = ""

   'Clear Text Boxes
   'Work Order and TXF file Name
    txtWO_TXF = ""
    txtWorkOrder = ""

   'Patient Details
    txtFileNo = ""
    txtPatientName = ""
    txtAge = ""
    txtUnits = ""
    txtSex = ""
    txtDiagnosis = ""
    txtMPDwo = ""
    txtOrderDate = ""
    txtTemplateEngineer = ""
    txtUidMPD = ""

   'CAD filr details
    txtInitialCADFile = ""
    txtCurrentCADFile = ""
   
   'Position to center of screen
    Left = (Screen.Width - cltxfdia.Width) / 2
    Top = (Screen.Height - cltxfdia.Height) / 2
    
    Set MainForm = cltxfdia

   'Get the path to the JOBST system installation directory
    g_sPathJOBST = fnPathJOBST()
    
   'Load combo box options
    cboSexPD.AddItem "Male"
    cboSexPD.AddItem "Female"

    cboUnitsWOD.AddItem "cm"
    cboUnitsWOD.AddItem "inches"

    PR_GetComboListFromFile cboDiagnosisPD, g_sPathJOBST & "\DEFAULTS\DIAGNOSI.DAT"
    PR_GetComboListFromFile cboDesignerWOD, g_sPathJOBST & "\DEFAULTS\DESIGNER.DAT"

   'Set workorder date to curent date
    txtOrderDateWOD.Text = Format$(Now, "dd/mm/yyyy")

   'Ensure that we timeout after approx 6 seconds
   'The timer is disabled on link close
    Timer1.Interval = 6000
    Timer1.Enabled = False

End Sub

Private Sub OK_Click()
   'Check that with the user that data is correct
   'If OK the update drawing
    Dim sDrafixTask As String

    If FN_ValidatePDandWOD() Then

        PR_Update_DDE_TexTBoxes

       'If the user has selected a file then use it
       'We need to update it
       'NB This can only happen if the txtInitialCADFile was equal
       '   to NEW.  For all others we disable the CAD_File button.
       '
        If Trim$(lblCADFile.Caption) <> Trim$(txtInitialCADFile.Text) And txtInitialCADFile <> "INSPECTION" Then
            PR_UpdateExisting_Macro "C:\JOBST\DRAW.D"
        Else
            PR_DRAFIX_Macro "C:\JOBST\DRAW.D"
        End If

        sDrafixTask = fnGetDrafixWindowTitleText()
        
       'Start the required task by using the appropriate
       'drafix macro
        If sDrafixTask <> "" Then
            AppActivate sDrafixTask
            SendKeys "@C:\JOBST\DRAW.D{enter}", True
            End
        Else
            MsgBox "Unable to find a running copy of DRAFIX!", 16, "Error - Updating Drawing"
        End If
    
    End If

End Sub

Private Sub PR_DRAFIX_Macro(sDrafixFile As String)
   'Create a DRAFIX macro file
   'NOTES
   '    This uses the same DDE fileds that are poked from DRAFIX
   '    It assumes that all data is correct and that all problems
   '    have been reconciled befor this is called.
   '    All layers and DB fields exist.
   '
    Dim sLine$
    
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

   'Define DRAFIX variables
    Print #fNum, "HANDLE hMPD;"
    Print #fNum, "XY     xyO, xyScale;"
    Print #fNum, "ANGLE  aAngle;"
    Print #fNum, "STRING sName, sTXF;"

   'Clear user selections etc
    Print #fNum, "UserSelection ("; QQ; "clear"; QQ; ");"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetStyle"; QC; "Table("; QQ; "find"; QCQ; "style"; QCQ; "bylayer"; QQ; "));"
    Print #fNum, "Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "Table("; QQ; "find"; QCQ; "layer"; QCQ; "bylayer"; QQ; "));"

   'Display Hour Glass symbol
    Print #fNum, "Display ("; QQ; "cursor"; QCQ; "wait"; QCQ; "Updating Drawing"; QQ; ");"
   
   'Set symbol library
    sLine = FN_EscapeSlashesInString(g_sPathJOBST) & "\\JOBST.SLB"
    Print #fNum, "SetSymbolLibrary("; QQ; sLine; QQ; ");"

   '                    ---------------
   '                    PATIENT DETAILS
   '                    ---------------
   '
   'Writes the patient details to the DRAFIX macro file
   'Checks the field txtUidMPD.
   '    If this field is EMPTY then
   '        Inserts a new Patient Details Symbol.
   '        get Entity handle
   '    Else
   '        get Entity handle for given UID from txtUidMPD
   '
   '    Update data base fields


    If txtUidMPD = "" Then
       'Insert a mainpatientdetails symbol
        Print #fNum, "if ( Symbol("; QQ; "find"; QCQ; "mainpatientdetails"; QQ; ")){"
        Print #fNum, "  Execute ("; QQ; "menu"; QCQ; "SetLayer"; QC; "Table("; QQ; "find"; QCQ; "layer"; QCQ; "titlebox"; QQ; "));"
        Print #fNum, "  GetData("; QQ; "SheetSize"; QC; "&xyO);"
        Print #fNum, "  xyO.x  = xyO.x - 5.5;"
        Print #fNum, "  xyO.y  = xyO.y - 6.125; "
        Print #fNum, "  hMPD = AddEntity("; QQ; "symbol"; QCQ; "mainpatientdetails"; QC; "xyO);"
        Print #fNum, "  }"
        Print #fNum, "else"
        Print #fNum, "  Exit(%cancel, "; QQ; "Can't find > mainpatientdetails < symbol to insert\nCheck your installation, that JOBST.SLB exists"; QQ; ");"
    Else
       'Get Handle and Origin
        Print #fNum, "hMPD = UID ("; QQ; "find"; QC; Val(txtUidMPD.Text); ");"
        Print #fNum, "if (hMPD)"
        Print #fNum, "  GetGeometry(hMPD, &sName, &xyO, &xyScale, &aAngle);"
        Print #fNum, "else"
        Print #fNum, "  Exit(%cancel,"; QQ; "Can't find > mainpatientdetails < symbol, Insert Patient Data"; QQ; ");"
    End If

   'Update Database Fields
    Print #fNum, "SetDBData( hMPD"; CQ; "fileno"; QCQ; txtFileNo; QQ; ");"
    Print #fNum, "SetDBData( hMPD"; CQ; "patient"; QCQ; txtPatientName; QQ; ");"
    Print #fNum, "SetDBData( hMPD"; CQ; "age"; QCQ; txtAge; QQ; ");"
    Print #fNum, "SetDBData( hMPD"; CQ; "units"; QCQ; txtUnits; QQ; ");"
    Print #fNum, "SetDBData( hMPD"; CQ; "sex"; QCQ; txtSex; QQ; ");"
    Print #fNum, "SetDBData( hMPD"; CQ; "Diagnosis"; QCQ; txtDiagnosis; QQ; ");"
    Print #fNum, "SetDBData( hMPD"; CQ; "WorkOrder"; QCQ; txtMPDwo; QQ; ");"
    Print #fNum, "SetDBData( hMPD"; CQ; "orderdate"; QCQ; txtOrderDate; QQ; ");"
    Print #fNum, "SetDBData( hMPD"; CQ; "TemplateEngineer"; QCQ; txtTemplateEngineer; QQ; ");"
    
   'Close the Macro File
    Print #fNum, ""
    Print #fNum, "// -- End of MACRO --"
    Close #fNum
   
End Sub

Private Sub PR_GetComboListFromFile(Combo_Name As Control, sFileName As String)
   'General procedure to create the list section of
   'a combo box reading the data from a file
    
    Dim sLine As String, fFileNum As Integer

    fFileNum = FreeFile
    
    If FileLen(sFileName) = 0 Then
        MsgBox sFileName & "Not found", 48
        Exit Sub
    End If

    Open sFileName For Input As fFileNum
    Do While Not EOF(fFileNum)
        Line Input #fFileNum, sLine
        Combo_Name.AddItem sLine
        Loop
    Close #fFileNum

End Sub


Private Sub PR_GetDrawingData()
   'Procedure to get the data from the DDE text boxes.
   'IE
   '    The data that has been loaded from the drawing.
   '
   'It will be used only if no TRITON TXF file exists.
   '
   'INPUT
   '    DDE text boxes
   '
   'OUTPUT
   '    frmPatientDetails
   '                - All of the text boxes in this will be
   '                  updated
   '
   '    frmWorkOrderDetails
   '                - All of the text boxes in this will be
   '                  updated
   '
   '    g_LineItem  - Global variable which is an array
   '                  of type Item, Declared as follows
   '
   '                  Type Item
   '                      PosNo As Integer
   '                      CatNo As String
   '                      Desc  As String
   '                      Group As String
   '                  End Type
   '
   '                  Global Const MAX_LINEITEMS = 200
   '                  Global g_LineItem(1 To MAX_LINEITEMS) As Item
   '
   'NOTES
   '    The line item position number is used as the index into
   '    the global array g_LineItem.  This is not particulary
   '    efficient on storage, but it means no sort routine is required.
   '    A worse effect is that it limits the Work Order to
   '    MAX_LINEITEMS / 10 Major items
   '
    Dim sString As String
    Dim sWorkOrder As String, ii%

   'Update patient details from DDE text boxes
   'Patient file no
    txtFileNoPD = txtFileNo

   'Patient Name
    txtNamePD = txtPatientName

   'Patient age
    txtAgePD = txtAge

   'Patient Sex
    sString = Trim$(txtSex.Text)
    If Left$(sString, 1) = "M" Then
        cboSexPD.ListIndex = 0
    ElseIf Left$(sString, 1) = "F" Then
        cboSexPD.ListIndex = 1
    Else
        cboSexPD.ListIndex = -1
    End If
    
   'Diagnosis Default to Burns
    If Trim$(txtDiagnosis.Text) <> "" Then
        cboDiagnosisPD.Text = Trim$(txtDiagnosis.Text)
    Else
        cboDiagnosisPD.Text = "Burns"
    End If


   'Other Work order details from DDE text boxes
   'Current Work Order number
    txtWorkOrderWOD = txtMPDwo

   'Current Work Order date
   'Only if date given else stick with date from FormLoad event
    If txtOrderDate <> "" Then txtOrderDateWOD = txtOrderDate

   'Current Work Order units
    sString = Trim$(txtUnits.Text)
    If Left$(sString, 1) = "c" Then
        cboUnitsWOD.ListIndex = 0
    ElseIf Left$(sString, 1) = "i" Then
        cboUnitsWOD.ListIndex = 1
    Else
        cboUnitsWOD.ListIndex = 0   'Default to "cm"
    End If

   'Template designer
    cboDesignerWOD.Text = txtTemplateEngineer.Text


End Sub

Private Sub PR_GetTXFData(ByVal sFileName As String)
   'Procedure to read the Triton eXchange File.
   '
   'Places data into the relevent patient details
   'and work order text boxes.
   '
   'Places Line Item data into the LineItem array of Item
   '
   'INPUT
   '    sFileName
   '                - Filename specified above.
   '                  The file contains the following
   '
   '                     <File #>
   '                     <Name>
   '                     <Customer #>       (Not used)
   '                     <Sex>
   '                     <Age>
   '                     <Date Of Birth>    (Not used)
   '                     <Diagnosis>
   '                     <Position #1>      (Not used)
   '                     <Item Code #1>         .
   '                         .                  .
   '                         .                  .
   '                         .                  .
   '                         .                  .
   '                     <Position #n>          .
   '                     <Item Code #n>         .
   '
   '     MAX_LINEITEMS
   '                - Constant
   '
   'OUTPUT
   '    frmPatientDetails
   '                - All of the text boxes in this will be
   '                  updated
   '
   '    frmWorkOrderDetails
   '                - All of the text boxes in this will be
   '                  updated
   '
   '    g_LineItem  - Global variable which is an array
   '                  of type Item, Declared as follows
   '
   '                  Type Item
   '                      PosNo As Integer
   '                      CatNo As String
   '                      Desc  As String
   '                      Group As String
   '                  End Type
   '
   '                  Global Const MAX_LINEITEMS = 200
   '                  Global g_LineItem(1 To MAX_LINEITEMS) As Item
   '
   'NOTES
   '    The line item position number is used as the index into
   '    the global array g_LineItem.  This is not particulary
   '    efficient on storage, but it means no sort routine is required.
   '    A worse effect is that it limits the Work Order to
   '    MAX_LINEITEMS / 10 Major items

    Dim sLine As String, fFileNum As Integer
    Dim sDiagnosis As String, sCatNo As String, sCatDesc As String
    Dim sCatGroup As String
    Dim ii As Integer, iPosNo As Integer, nPosNo As Double
    Dim nLength As Long

    NL = Chr$(13) & Chr$(10) 'New Line  (this is global)
    TABBED = Chr$(9)

    PR_Message "Reading TXF file." & NL
    fFileNum = FreeFile

    If FileLen(sFileName) = 0 Then
        MsgBox sFileName & "Not found", 48
        Exit Sub
    End If

    Open sFileName For Input As fFileNum

   'Patient File No
    PR_ReadLine fFileNum, sLine
    txtFileNoPD.Text = Trim$(sLine)

   'Patient Name
    PR_ReadLine fFileNum, sLine
    txtNamePD.Text = Trim$(sLine)

   'Customer Number (NOT USED)
    PR_ReadLine fFileNum, sLine

   'Patient Sex
    PR_ReadLine fFileNum, sLine
    sLine = Trim$(sLine)
    If UCase$(Left$(sLine, 1)) = "M" Then
        cboSexPD.ListIndex = 0
    ElseIf UCase$(Left$(sLine, 1)) = "F" Then
        cboSexPD.ListIndex = 1
    Else
        cboSexPD.ListIndex = -1
    End If
    g_sSexTXF = cboSexPD.List(cboSexPD.ListIndex)
    
   'Patient Age
    PR_ReadLine fFileNum, sLine
    txtAgePD.Text = Trim$(sLine)

   'Date Of Birth (NOT USED)
    PR_ReadLine fFileNum, sLine

   'Diagnosis
    PR_ReadLine fFileNum, sLine
    g_sDiagnosisTXF = Trim$(sLine)
    cboDiagnosisPD.Text = Trim$(g_sDiagnosisTXF)

    Close #fFileNum

   'Units
    Dim sString$
    sString = Trim$(txtUnits.Text)
    If Left$(sString, 1) = "c" Then
        cboUnitsWOD.ListIndex = 0
    ElseIf Left$(sString, 1) = "i" Then
        cboUnitsWOD.ListIndex = 1
    Else
        cboUnitsWOD.ListIndex = 0   'Default to "cm"
    End If

End Sub

Private Sub PR_Message(sText As String)
   'Small procedure to place the given TEXT string
   'into the txtMessage Text Box
   'Used just to simplify code in PR_TXFtoDwgComparison
    txtMessage.Text = txtMessage.Text & sText & NL

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

Private Sub PR_TXFtoDwgComparison()
   'The purpose of this is to compare the Patient Petails
   'taken from the TXF file with the Patient Details in the
   'Drawing that were poked to the DDE text Boxes.
   '
   'Messages are displayed in the txtMessage Text Box
   '
    NL = Chr$(13) & Chr$(10) 'New Line  (this is global)
    TABBED = Chr$(9)

   'Check if data is available from an existing drawing.
   'This is implied by txtUidMPD which contains the DRAFIX
   'UID for the "mainpatientdetails" symbol.  If this is empty
   'then there is no symbol, hence it is a NEW drawing and
   'the check does not need to be done

    If txtUidMPD = "" Then Exit Sub

    PR_Message "Comparing TXF file details to Drawing." & NL
        
    If Trim$(UCase$(txtFileNoPD.Text)) <> Trim$(UCase$(txtFileNo.Text)) Then
        PR_Message "File Numbers are different."
        PR_Message TABBED & "TXF file no :>" & Trim$(UCase$(txtFileNoPD.Text)) & "<"
        PR_Message TABBED & "Dwg file no :>" & Trim$(UCase$(txtFileNo.Text)) & "<" & NL
        Beep
    End If

    If g_sSexTXF <> txtSex.Text Then
        PR_Message "Patient Sex is different."
        PR_Message TABBED & "TXF Sex is : " & g_sSexTXF
        PR_Message TABBED & "Dwg Sex was: " & txtSex.Text & NL
        Beep
    End If

    If Trim$(UCase$(cboDiagnosisPD.Text)) <> Trim$(UCase$(txtDiagnosis.Text)) Then
        PR_Message "The Diagnoisis has changed."
        PR_Message TABBED & "TXF Diagnosis is :>" & Trim$(UCase$(cboDiagnosisPD.Text)) & "<"
        PR_Message TABBED & "Dwg Diagnosis was:>" & Trim$(UCase$(txtDiagnosis.Text)) & "<" & NL
        Beep
    End If

    If Val(txtAgePD.Text) <> Val(txtAge.Text) Then
        PR_Message "Patient Age has changed."
        PR_Message TABBED & "TXF Age is : " & txtAgePD.Text
        PR_Message TABBED & "Dwg Age was: " & txtAge.Text & NL
        Beep
    End If
    
    
End Sub

Private Sub PR_Update_DDE_TexTBoxes()
   '
   'Takes the data from
   '
   '    Patient Details (PD)        Text Boxes
   '    Work Order Details (WOD)    Text Boxes
   '
   'Formats and inserts it into the DDE text boxes.
   'Data from these boxes is used to create a DRAFIX
   'macro that will update the drawing.
   '
   'The use of DDE text boxes is therefor a misnomer
   'in this case but it is used to be consistent with
   'the transfer of data in the DRAFIX to VB direction
   '
   'NOTES
   '    The PD and WOD can be considered as the working
   '    areas.  The details in this area can come from 3
   '    possible sources:-
   '
   '    1.  The Triton eXchange File (TXF)
   '    2.  The existing drawing     (Dwg)
   '    3.  User input
   '
   '    If the details are correct then the DDE text
   '    boxes can be updated with the valid data
   '
   '
   'ASSUMPTIONS
   '    The data will be valid.  Therefor no error checking
   '
    
   'Patient Details
   '
    txtFileNo = txtFileNoPD
    
    txtPatientName = txtNamePD
    
    txtAge = txtAgePD
  
    txtSex = cboSexPD.Text
    
    txtDiagnosis = cboDiagnosisPD.Text


   'Work Order Details
   '
    txtWorkOrder = txtWorkOrderWOD

    txtMPDwo = txtWorkOrderWOD
    
    txtOrderDate = txtOrderDateWOD
    
    txtTemplateEngineer.Text = cboDesignerWOD.Text
    
    txtUnits = cboUnitsWOD.Text

End Sub

Private Sub PR_UpdateExisting_Macro(sDrafixFile As String)
   'Create a DRAFIX macro file
   'This is a hybrid system where the UPDATING macro has been
   'created as a drafix macro and this macro creating procedure is
   'simply used to provide the data to the macro
   '    CADLINK\CL_UPDTE.D
   '
   'NOTES
   '    This uses the same DDE fields that are poked from DRAFIX
   '    It assumes that all data is correct and that all problems
   '    have been reconciled before this is called.
   '    All layers and DB fields exist.
   '
    Dim sLine$
    
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

   'Define DRAFIX variables
    Print #fNum, "STRING sName, sFileNo, sPatient, sAge, sUnits, sSex, sWorkOrder, sTemplateEngineer, sOrderDate;"
    Print #fNum, "STRING sCADFile, sDiagnosis;"
   
   'Set Existing file name
    sLine = FN_EscapeSlashesInString((lblCADFile.Caption))
    Print #fNum, "sCADFile= "; QQ; sLine; QQ; ";"


   'Set strings for use in update macro
    Print #fNum, "sFileNo = "; QQ; txtFileNo; QQ; ";"
    Print #fNum, "sPatient = "; QQ; txtPatientName; QQ; ";"
    Print #fNum, "sAge = "; QQ; txtAge; QQ; ";"
    Print #fNum, "sUnits ="; QQ; txtUnits; QQ; ";"
    Print #fNum, "sSex = "; QQ; txtSex; QQ; ";"
    Print #fNum, "sDiagnosis ="; QQ; txtDiagnosis; QQ; ";"
    Print #fNum, "sWorkOrder ="; QQ; txtMPDwo; QQ; ";"
    Print #fNum, "sOrderDate ="; QQ; txtOrderDate; QQ; ";"
    Print #fNum, "sTemplateEngineer ="; QQ; txtTemplateEngineer; QQ; ";"
    
   'Call the update macro
    Print #fNum, "@"; g_sPathJOBST; "\CADLINK\CL_UPDTE.D;"

   'Close the Macro File
    Close #fNum

End Sub

Private Sub Timer1_Timer()
   'In case there is no Link_Close event
   'The programme will time out after approx 5 secs
    End
End Sub

