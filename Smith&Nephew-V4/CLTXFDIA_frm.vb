Option Strict Off
Option Explicit On
Imports Autodesk
Imports Autodesk.AutoCAD.ApplicationServices
Imports VB = Microsoft.VisualBasic
Imports Autodesk.AutoCAD.DatabaseServices
Imports BlockCreation
'Imports System.Windows.Forms
Friend Class cltxfdia
    Inherits System.Windows.Forms.Form
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


    'Private Sub CAD_File_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    '    'Find an existing CAD file to update
    '    '
    '    'CAD_File.DialogResult = DialogResult.None
    '    'Path to drawings
    '    Dim lpBuffer As New VB6.FixedLengthString(144) 'Minimum recommended wrt GetWindowsDirectory()
    '    Dim nBufferSize As Integer
    '    Dim nSize As Integer
    '    Dim WindowsDir As String

    '    nBufferSize = 143

    '    'Get the path to the Windows Directory to locate DRAFIX.INI
    '    '
    '    nSize = GetWindowsDirectory(lpBuffer.Value, nBufferSize)
    '    WindowsDir = VB.Left(lpBuffer.Value, nSize)

    '    'Get the path to the Drawing directory from
    '    'DRAFIX.INI
    '    '
    '    nSize = GetPrivateProfileString("Path", "PathDrawing", "C:\", lpBuffer.Value, nBufferSize, WindowsDir & "\DRAFIX.INI")

    '    CMDialog1Open.InitialDirectory = VB.Left(lpBuffer.Value, nSize)
    '    'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '    CMDialog1Open.Filter = "Drafix CAD (*.dwg)|*.dwg"
    '    CMDialog1Open.FileName = ""
    '    CMDialog1Open.CheckFileExists = True
    '    CMDialog1Open.CheckPathExists = True
    '    CMDialog1Open.DefaultExt = "dwg"
    '    Dim res As DialogResult = CMDialog1Open.ShowDialog()
    '    If res = DialogResult.OK Or res = DialogResult.Yes Then
    '        If CMDialog1Open.FileName <> "" Then
    '            'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '            'If Dir(CMDialog1Open.FileName) = "" Then
    '            '    PR_Message("CAD File not found." & Chr(13) & Chr(10))
    '            '    PR_Message("   File: " & CMDialog1Open.FileName & Chr(13) & Chr(10))
    '            '    MsgBox("CAD File not found.", 48, "Patient Details")
    '            'Else
    '            'lblCADFile.Text = CMDialog1Open.FileName
    '            'End If
    '        End If

    '    End If


    'End Sub

    Private Sub Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cancel.Click
        saveInfoToDWG()
        Me.Close()
    End Sub

    'UPGRADE_WARNING: Event cboUnitsWOD.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboUnitsWOD_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboUnitsWOD.SelectedIndexChanged
        CLTXFDIA1.g_sUnits = VB6.GetItemString(cboUnitsWOD, cboUnitsWOD.SelectedIndex)
    End Sub

    Private Function FN_EscapeSlashesInString(ByRef sAssignedString As Object) As String
        'Search through the string looking for " (double quote characater)
        'If found use \ (Backslash) to escape it
        '
        Dim ii As Short
        'UPGRADE_NOTE: Char was upgraded to Char_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim Char_Renamed As String
        Dim sEscapedString As String

        FN_EscapeSlashesInString = ""

        For ii = 1 To Len(sAssignedString)
            'UPGRADE_WARNING: Couldn't resolve default property of object sAssignedString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Char_Renamed = Mid(sAssignedString, ii, 1)
            If Char_Renamed = "\" Then
                sEscapedString = sEscapedString & "\" & Char_Renamed
            Else
                sEscapedString = sEscapedString & Char_Renamed
            End If
        Next ii

        FN_EscapeSlashesInString = sEscapedString

    End Function


    Private Function FN_ValidatePDandWOD() As Short
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
        Dim sMessage As String
        Dim iMsg As Short
        Const IDOK As Short = 1
        Const IDCANCEL As Short = 2
        Const IDYES As Short = 6
        Const IDNO As Short = 7

        'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.NL = Chr(13) & Chr(10) 'Carriage Return & New Line
        TABBED = Chr(9) 'Tab

        ''txtMessage.Text = ""

        'Patient Details
        '
        If txtFileNoPD.Text = "" Then PR_Message("Missing File Number.")

        If txtNamePD.Text = "" Then PR_Message("Missing Patient name.")

        If Val(txtAgePD.Text) = 0 Then PR_Message("Invalid Patient Age.")

        If cboSexPD.SelectedIndex = -1 Then PR_Message("Missing Patient SEX.")

        If cboDiagnosisPD.Text = "" Then PR_Message("Missing Diagnosis.")


        'Work Order Details
        '
        If txtWorkOrderWOD.Text = "" Then PR_Message("Missing Work Order Number.")

        If txtOrderDateWOD.Text = "" Then PR_Message("Missing Date.")

        If cboDesignerWOD.Text = "" Then PR_Message("Missing Template Designer.")

        If cboUnitsWOD.SelectedIndex = -1 Then PR_Message("Missing Units.")

        If cboUnitsWOD.Text <> txtUnits.Text And txtUnits.Text <> "" Then PR_Message("WARNING! Inconsistent units.  The drawing uses " & txtUnits.Text & ", you have selected " & cboUnitsWOD.Text & ".  This may cause problems when re-drawing template patterns.")

        'Check Patient, Template Engineer and Diagnosis for double quotes
        'Chr$(34) 'Double quotes (")this is to avoid problems later on
        'in drafix macros.
        'Greatfull thanks to Kathleen Ryan of JOBST for highlighting this!
        '
        If InStr(1, cboDesignerWOD.Text, Chr(34)) Then
            MsgBox("Designer name includes Double Quotes (""). This is not allowed!", 16)
            FN_ValidatePDandWOD = False
            PR_Message("Designer name includes Double Quotes.")
            Exit Function
        End If
        If InStr(1, txtNamePD.Text, Chr(34)) Then
            MsgBox("Patient name includes Double Quotes (""). This is not allowed!", 16)
            FN_ValidatePDandWOD = False
            PR_Message("Patient name includes Double Quotes.")
            Exit Function
        End If
        If InStr(1, cboDiagnosisPD.Text, Chr(34)) Then
            MsgBox("Diagnosis includes Double Quotes (""). This is not allowed!", 16)
            FN_ValidatePDandWOD = False
            PR_Message("Diagnosis includes Double Quotes.")
            Exit Function
        End If

        'If Len(txtMessage.Text) > 0 Then
        '    PR_Message("")
        '    sMessage = txtMessage.Text & "Do you wish to Continue?"
        '    iMsg = MsgBox(sMessage, 36, "Incomplete Patient and Order Details")
        '    If iMsg = IDYES Then FN_ValidatePDandWOD = 1
        '    If iMsg = IDNO Then FN_ValidatePDandWOD = 0
        'Else
        FN_ValidatePDandWOD = 1
        'End If

    End Function

    'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
    Private Sub Form_LinkClose()
        'Disable TIMEOUT Timer
        Timer1.Enabled = False

        'Based on the data in the txtWorkOrder and the
        'txtWO_TXF field
        '
        '
        If txtInvokedFrom.Text = "imageABLE" Then
            'If data is contained in a TXF file
            'we read the data from the TXF file
            txtWorkOrderWOD.Text = txtWorkOrder.Text
            PR_GetTXFData((txtWO_TXF.Text))

            'If not a new Drawing then Compare Patient data
            'from the existing drawing with TXF file data
            PR_TXFtoDwgComparison()

            'Display the file that the drawing is based on
            'Modify caption
            'CAD_File.Enabled = False
            If txtInitialCADFile.Text <> "INSPECTION" Then
                'frmCADFile.Text = "Previous Drafix CAD File"
                'lblCADFile.Text = txtInitialCADFile.Text
                'If txtInitialCADFile.Text = "NEW" Then CAD_File.Enabled = True
            Else
                'Ensure that Designer is retained on inspection
                'frmCADFile.Text = "Drafix CAD File"
                'lblCADFile.Text = txtCurrentCADFile.Text
                cboDesignerWOD.Text = txtTemplateEngineer.Text
                'Disable the OK button to stop inspectors from
                'being able to save as this is unessesary in the inspection
                'stage
                OK.Enabled = False
            End If

        Else
            'Get the data from the poked DDE text boxes
            PR_GetDrawingData()
            'frmCADFile.Text = "Drafix CAD File"
            'lblCADFile.Text = txtCurrentCADFile.Text
            'CAD_File.Enabled = False
        End If

        Show()

    End Sub

    Private Sub cltxfdia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            'Hide form to allow DDE to happen
            'Show form on LinkClose
            Hide()

            'The VB programme is initiated either through
            'DRAFIX or imageABLE
            txtInvokedFrom.Text = ""

            'Clear Text Boxes
            'Work Order and TXF file Name
            txtWO_TXF.Text = ""
            txtWorkOrder.Text = ""

            'Patient Details
            txtFileNo.Text = ""
            txtPatientName.Text = ""
            txtAge.Text = ""
            txtUnits.Text = ""
            txtSex.Text = ""
            txtDiagnosis.Text = ""
            txtMPDwo.Text = ""
            txtOrderDate.Text = ""
            txtTemplateEngineer.Text = ""
            txtUidMPD.Text = ""

            'CAD filr details
            txtInitialCADFile.Text = ""
            txtCurrentCADFile.Text = ""


            'Position to center of screen
            Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
            Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)

            CLTXFDIA1.MainForm = Me

            'Get the path to the JOBST system installation directory
            CLTXFDIA1.g_sPathJOBST = fnPathJOBST()

            'Load combo box options
            cboSexPD.Items.Add("Male")
            cboSexPD.Items.Add("Female")
            cboSexPD.SelectedIndex = 1

            cboUnitsWOD.Items.Add("cm")
            cboUnitsWOD.Items.Add("inches")
            cboUnitsWOD.SelectedIndex = 0

            ''------------PR_GetComboListFromFile(cboDiagnosisPD, CLTXFDIA1.g_sPathJOBST & "\DEFAULTS\DIAGNOSI.DAT")
            ''---------PR_GetComboListFromFile(cboDesignerWOD, CLTXFDIA1.g_sPathJOBST & "\DEFAULTS\DESIGNER.DAT")
            Dim sSettingsPath As String = fnGetSettingsPath("LookupTables")
            PR_GetComboListFromFile(cboDiagnosisPD, sSettingsPath & "\DIAGNOSI.DAT")
            PR_GetComboListFromFile(cboDesignerWOD, sSettingsPath & "\DESIGNER.DAT")

            'Set workorder date to curent date
            ''txtOrderDateWOD.Text = System.DateTime.Now.ToString("dd/MM/yyyy") ' VB6.Format(Now, "dd/mm/yyyy")
            txtOrderDateWOD.Value = System.DateTime.Now
            DateTimePicker1.Value = System.DateTime.Now

            readDWGInfo()
            'Ensure that we timeout after approx 6 seconds
            'The timer is disabled on link close
            Timer1.Interval = 6000
            Timer1.Enabled = False
            Dim ii As Short
            Dim strDesigner As String = fnGetSettingsPath("Designer_Name")
            For ii = 0 To (cboDesignerWOD.Items.Count - 1)
                If VB6.GetItemString(cboDesignerWOD, ii) = strDesigner Then
                    cboDesignerWOD.SelectedIndex = ii
                End If
            Next ii
            'cboDesignerWOD
            Show()
            txtFileNoPD.Focus()
        Catch ex As Exception
            Me.Close()
        End Try
    End Sub

    'Sub PR_GetComboListFromFile(ByRef Combo_Name As System.Windows.Forms.Control, ByRef sFileName As String)
    '    'General procedure to create the list section of
    '    'a combo box reading the data from a file

    '    Dim sLine As String
    '    Dim fFileNum As Short

    '    fFileNum = FreeFile()

    '    If FileLen(sFileName) = 0 Then
    '        MsgBox(sFileName & "Not found", 48)
    '        Exit Sub
    '    End If

    '    FileOpen(fFileNum, sFileName, OpenMode.Input)
    '    Do While Not EOF(fFileNum)
    '        sLine = LineInput(fFileNum)
    '        'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        Combo_Name.AddItem(sLine)
    '    Loop
    '    FileClose(fFileNum)

    'End Sub

    Private Sub OK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OK.Click
        ''Check that with the user that data is correct
        ''If OK the update drawing
        Dim sError As String = ""
        If txtFileNoPD.Text = "" Then
            sError = "Missing File Number"
        End If
        If txtNamePD.Text = "" Then
            sError = sError & Environment.NewLine & "Missing Patient Name"
        End If
        If txtAgePD.Text = "" Then
            sError = sError & Environment.NewLine & "Invalid Patient Age"
        End If
        If cboDiagnosisPD.Text = "" Then
            sError = sError & Environment.NewLine & "Missing Patient Diagnosis"
        End If
        If txtWorkOrderWOD.Text = "" Then
            sError = sError & Environment.NewLine & "Missing Work Order Number"
        End If
        If cboDesignerWOD.Text = "" Then
            sError = sError & Environment.NewLine & "Missing Template Designer"
        End If
        If sError <> "" Then
            sError = sError & Environment.NewLine & Environment.NewLine & "This data is mandatory" & Environment.NewLine & "The Titleblock should not be created when data is missing"
            MsgBox(sError, 32, "New Patient Details")
            Exit Sub
        End If
        saveInfoToDWG()

        'Dim obj As New BlockCreation.BlockCreation
        'obj.sLayout = fnGetSettingsPath("DRAW_IN_LAYOUT")
        'obj.sUnits = fnGetSettingsPath("UNITS")
        'obj.InsertBlockInAutoCAD(txtFileNoPD.Text, txtNamePD.Text, cboDiagnosisPD.Text, txtAgePD.Text, cboSexPD.Text, txtWorkOrderWOD.Text, txtOrderDateWOD.Text, cboDesignerWOD.Text, cboUnitsWOD.Text)

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            If acBlkTbl.Has("MAINPATIENTDETAILS") Then
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.PaperSpace), OpenMode.ForWrite)
                For Each objID As ObjectId In acBlkTblRec
                    Dim dbObj As DBObject = acTrans.GetObject(objID, OpenMode.ForWrite)
                    If TypeOf dbObj Is BlockReference Then
                        Dim blkRef As BlockReference = dbObj
                        If blkRef.Name = "MAINPATIENTDETAILS" Then
                            For Each attributeID As ObjectId In blkRef.AttributeCollection
                                Dim attRefObj As DBObject = acTrans.GetObject(attributeID, OpenMode.ForWrite)
                                If TypeOf attRefObj Is AttributeReference Then
                                    Dim acAttRef As AttributeReference = attRefObj
                                    If acAttRef.Tag.ToUpper().Equals("FILENO", StringComparison.InvariantCultureIgnoreCase) Then
                                        acAttRef.TextString = txtFileNoPD.Text
                                    ElseIf acAttRef.Tag.ToUpper().Equals("PATIENT", StringComparison.InvariantCultureIgnoreCase) Then
                                        acAttRef.TextString = txtNamePD.Text
                                    ElseIf acAttRef.Tag.ToUpper().Equals("DIAGNOSIS", StringComparison.InvariantCultureIgnoreCase) Then
                                        acAttRef.TextString = cboDiagnosisPD.Text
                                    ElseIf acAttRef.Tag.ToUpper().Equals("AGE", StringComparison.InvariantCultureIgnoreCase) Then
                                        acAttRef.TextString = txtAgePD.Text
                                    ElseIf acAttRef.Tag.ToUpper().Equals("SEX", StringComparison.InvariantCultureIgnoreCase) Then
                                        acAttRef.TextString = cboSexPD.Text
                                    ElseIf acAttRef.Tag.ToUpper().Equals("WORKORDER", StringComparison.InvariantCultureIgnoreCase) Then
                                        acAttRef.TextString = txtWorkOrderWOD.Text
                                    ElseIf acAttRef.Tag.ToUpper().Equals("ORDERDATE", StringComparison.InvariantCultureIgnoreCase) Then
                                        acAttRef.TextString = txtOrderDateWOD.Text
                                    ElseIf acAttRef.Tag.ToUpper().Equals("TEMPLATEENGINEER", StringComparison.InvariantCultureIgnoreCase) Then
                                        acAttRef.TextString = cboDesignerWOD.Text
                                    ElseIf acAttRef.Tag.ToUpper().Equals("UNITS", StringComparison.InvariantCultureIgnoreCase) Then
                                        acAttRef.TextString = cboUnitsWOD.Text
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
            acTrans.Commit()
        End Using
        fnWriteValueInSettingsFile("Designer_Name", cboDesignerWOD.Text)
        Close()

    End Sub

    Private Sub PR_DRAFIX_Macro(ByRef sDrafixFile As String)
        'Create a DRAFIX macro file
        'NOTES
        '    This uses the same DDE fileds that are poked from DRAFIX
        '    It assumes that all data is correct and that all problems
        '    have been reconciled befor this is called.
        '    All layers and DB fields exist.
        '
        Dim sLine As String

        'Open file
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.fNum = FreeFile()
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileOpen(CLTXFDIA1.fNum, sDrafixFile, VB.OpenMode.Output)

        'Initialise String globals
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.CC = Chr(44) 'The comma (,)
        'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.NL = Chr(10) 'The new line character
        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.QQ = Chr(34) 'Double quotes (")
        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.QCQ = CLTXFDIA1.QQ & CLTXFDIA1.CC & CLTXFDIA1.QQ
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.QC = CLTXFDIA1.QQ & CLTXFDIA1.CC
        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.CQ = CLTXFDIA1.CC & CLTXFDIA1.QQ

        'Write header information etc. to the DRAFIX macro file
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "//Patient    - " & txtPatientName.Text & CLTXFDIA1.CC & " " & txtFileNo.Text & CLTXFDIA1.CC)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "//Work Order - " & txtWorkOrder.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "//by Visual Basic")

        'Define DRAFIX variables
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "HANDLE hMPD;")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "XY     xyO, xyScale;")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "ANGLE  aAngle;")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "STRING sName, sTXF;")

        'Clear user selections etc
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "UserSelection (" & CLTXFDIA1.QQ & "clear" & CLTXFDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "Execute (" & CLTXFDIA1.QQ & "menu" & CLTXFDIA1.QCQ & "SetStyle" & CLTXFDIA1.QC & "Table(" & CLTXFDIA1.QQ & "find" & CLTXFDIA1.QCQ & "style" & CLTXFDIA1.QCQ & "bylayer" & CLTXFDIA1.QQ & "));")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "Execute (" & CLTXFDIA1.QQ & "menu" & CLTXFDIA1.QCQ & "SetLayer" & CLTXFDIA1.QC & "Table(" & CLTXFDIA1.QQ & "find" & CLTXFDIA1.QCQ & "layer" & CLTXFDIA1.QCQ & "bylayer" & CLTXFDIA1.QQ & "));")

        'Display Hour Glass symbol
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "Display (" & CLTXFDIA1.QQ & "cursor" & CLTXFDIA1.QCQ & "wait" & CLTXFDIA1.QCQ & "Updating Drawing" & CLTXFDIA1.QQ & ");")

        'Set symbol library
        sLine = FN_EscapeSlashesInString(CLTXFDIA1.g_sPathJOBST) & "\\JOBST.SLB"
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "SetSymbolLibrary(" & CLTXFDIA1.QQ & sLine & CLTXFDIA1.QQ & ");")

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


        If txtUidMPD.Text = "" Then
            'Insert a mainpatientdetails symbol
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "if ( Symbol(" & CLTXFDIA1.QQ & "find" & CLTXFDIA1.QCQ & "mainpatientdetails" & CLTXFDIA1.QQ & ")){")
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "  Execute (" & CLTXFDIA1.QQ & "menu" & CLTXFDIA1.QCQ & "SetLayer" & CLTXFDIA1.QC & "Table(" & CLTXFDIA1.QQ & "find" & CLTXFDIA1.QCQ & "layer" & CLTXFDIA1.QCQ & "titlebox" & CLTXFDIA1.QQ & "));")
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "  GetData(" & CLTXFDIA1.QQ & "SheetSize" & CLTXFDIA1.QC & "&xyO);")
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "  xyO.x  = xyO.x - 5.5;")
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "  xyO.y  = xyO.y - 6.125; ")
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "  hMPD = AddEntity(" & CLTXFDIA1.QQ & "symbol" & CLTXFDIA1.QCQ & "mainpatientdetails" & CLTXFDIA1.QC & "xyO);")
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "  }")
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "else")
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "  Exit(%cancel, " & CLTXFDIA1.QQ & "Can't find > mainpatientdetails < symbol to insert\nCheck your installation, that JOBST.SLB exists" & CLTXFDIA1.QQ & ");")
        Else
            'Get Handle and Origin
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "hMPD = UID (" & CLTXFDIA1.QQ & "find" & CLTXFDIA1.QC & Val(txtUidMPD.Text) & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "if (hMPD)")
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "  GetGeometry(hMPD, &sName, &xyO, &xyScale, &aAngle);")
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "else")
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CLTXFDIA1.fNum, "  Exit(%cancel," & CLTXFDIA1.QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & CLTXFDIA1.QQ & ");")
        End If

        'Update Database Fields
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "SetDBData( hMPD" & CLTXFDIA1.CQ & "fileno" & CLTXFDIA1.QCQ & txtFileNo.Text & CLTXFDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "SetDBData( hMPD" & CLTXFDIA1.CQ & "patient" & CLTXFDIA1.QCQ & txtPatientName.Text & CLTXFDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "SetDBData( hMPD" & CLTXFDIA1.CQ & "age" & CLTXFDIA1.QCQ & txtAge.Text & CLTXFDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "SetDBData( hMPD" & CLTXFDIA1.CQ & "units" & CLTXFDIA1.QCQ & txtUnits.Text & CLTXFDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "SetDBData( hMPD" & CLTXFDIA1.CQ & "sex" & CLTXFDIA1.QCQ & txtSex.Text & CLTXFDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "SetDBData( hMPD" & CLTXFDIA1.CQ & "Diagnosis" & CLTXFDIA1.QCQ & txtDiagnosis.Text & CLTXFDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "SetDBData( hMPD" & CLTXFDIA1.CQ & "WorkOrder" & CLTXFDIA1.QCQ & txtMPDwo.Text & CLTXFDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "SetDBData( hMPD" & CLTXFDIA1.CQ & "orderdate" & CLTXFDIA1.QCQ & txtOrderDate.Text & CLTXFDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "SetDBData( hMPD" & CLTXFDIA1.CQ & "TemplateEngineer" & CLTXFDIA1.QCQ & txtTemplateEngineer.Text & CLTXFDIA1.QQ & ");")

        'Close the Macro File
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "// -- End of MACRO --")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileClose(CLTXFDIA1.fNum)

    End Sub

    Private Sub PR_GetComboListFromFile(ByRef Combo_Name As System.Windows.Forms.ComboBox, ByRef sFileName As String)
        'General procedure to create the list section of
        'a combo box reading the data from a file

        Dim sLine As String
        Dim fFileNum As Short

        fFileNum = FreeFile()

        If FileLen(sFileName) = 0 Then
            MsgBox(sFileName & "Not found", 48)
            Exit Sub
        End If

        FileOpen(fFileNum, sFileName, VB.OpenMode.Input)
        Do While Not EOF(fFileNum)
            sLine = LineInput(fFileNum)
            'UPGRADE_WARNING: Couldn't resolve default property of object Combo_Name.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Combo_Name.Items.Add(sLine) ' .AddItem(sLine)
        Loop
        FileClose(fFileNum)

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
        Dim sWorkOrder As String
        Dim ii As Short

        'Update patient details from DDE text boxes
        'Patient file no
        txtFileNoPD.Text = txtFileNo.Text

        'Patient Name
        txtNamePD.Text = txtPatientName.Text

        'Patient age
        txtAgePD.Text = txtAge.Text

        'Patient Sex
        sString = Trim(txtSex.Text)
        If VB.Left(sString, 1) = "M" Then
            cboSexPD.SelectedIndex = 0
        ElseIf VB.Left(sString, 1) = "F" Then
            cboSexPD.SelectedIndex = 1
        Else
            cboSexPD.SelectedIndex = -1
        End If

        'Diagnosis Default to Burns
        If Trim(txtDiagnosis.Text) <> "" Then
            cboDiagnosisPD.Text = Trim(txtDiagnosis.Text)
        Else
            cboDiagnosisPD.Text = "Burns"
        End If


        'Other Work order details from DDE text boxes
        'Current Work Order number
        txtWorkOrderWOD.Text = txtMPDwo.Text

        'Current Work Order date
        'Only if date given else stick with date from FormLoad event
        If txtOrderDate.Text <> "" Then txtOrderDateWOD.Text = txtOrderDate.Text

        'Current Work Order units
        sString = Trim(txtUnits.Text)
        If VB.Left(sString, 1) = "c" Then
            cboUnitsWOD.SelectedIndex = 0
        ElseIf VB.Left(sString, 1) = "i" Then
            cboUnitsWOD.SelectedIndex = 1
        Else
            cboUnitsWOD.SelectedIndex = 0 'Default to "cm"
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

        Dim sLine As String
        Dim fFileNum As Short
        Dim sCatNo, sDiagnosis, sCatDesc As String
        Dim sCatGroup As String
        Dim ii, iPosNo As Short
        Dim nPosNo As Double
        Dim nLength As Integer

        'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.NL = Chr(13) & Chr(10) 'New Line  (this is global)
        TABBED = Chr(9)

        'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_Message("Reading TXF file." & CLTXFDIA1.NL)
        fFileNum = FreeFile()

        If FileLen(sFileName) = 0 Then
            MsgBox(sFileName & "Not found", 48)
            Exit Sub
        End If

        FileOpen(fFileNum, sFileName, VB.OpenMode.Input)

        'Patient File No
        PR_ReadLine(fFileNum, sLine)
        txtFileNoPD.Text = Trim(sLine)

        'Patient Name
        PR_ReadLine(fFileNum, sLine)
        txtNamePD.Text = Trim(sLine)

        'Customer Number (NOT USED)
        PR_ReadLine(fFileNum, sLine)

        'Patient Sex
        PR_ReadLine(fFileNum, sLine)
        sLine = Trim(sLine)
        If UCase(VB.Left(sLine, 1)) = "M" Then
            cboSexPD.SelectedIndex = 0
        ElseIf UCase(VB.Left(sLine, 1)) = "F" Then
            cboSexPD.SelectedIndex = 1
        Else
            cboSexPD.SelectedIndex = -1
        End If
        g_sSexTXF = VB6.GetItemString(cboSexPD, cboSexPD.SelectedIndex)

        'Patient Age
        PR_ReadLine(fFileNum, sLine)
        txtAgePD.Text = Trim(sLine)

        'Date Of Birth (NOT USED)
        PR_ReadLine(fFileNum, sLine)

        'Diagnosis
        PR_ReadLine(fFileNum, sLine)
        g_sDiagnosisTXF = Trim(sLine)
        cboDiagnosisPD.Text = Trim(g_sDiagnosisTXF)

        FileClose(fFileNum)

        'Units
        Dim sString As String
        sString = Trim(txtUnits.Text)
        If VB.Left(sString, 1) = "c" Then
            cboUnitsWOD.SelectedIndex = 0
        ElseIf VB.Left(sString, 1) = "i" Then
            cboUnitsWOD.SelectedIndex = 1
        Else
            cboUnitsWOD.SelectedIndex = 0 'Default to "cm"
        End If

    End Sub

    Private Sub PR_Message(ByRef sText As String)
        'Small procedure to place the given TEXT string
        'into the txtMessage Text Box
        'Used just to simplify code in PR_TXFtoDwgComparison
        'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''txtMessage.Text = txtMessage.Text & sText & CLTXFDIA1.NL

    End Sub

    Private Sub PR_ReadLine(ByVal fFile As Short, ByRef sLine As String)
        'Read a line, character at a time up to either the
        '   NL character
        'or
        '   CRNL characters
        '
        'UPGRADE_NOTE: Char was upgraded to Char_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim NewLine, CR As String
        Dim Char_Renamed As New VB6.FixedLengthString(1)
        NewLine = Chr(10)
        CR = Chr(13)
        sLine = ""

        Do Until EOF(fFile)
            Char_Renamed.Value = InputString(fFile, 1)
            If Char_Renamed.Value = NewLine Then Exit Do
            If Char_Renamed.Value <> CR Then sLine = sLine & Char_Renamed.Value
        Loop

    End Sub

    Private Sub PR_TXFtoDwgComparison()
        'The purpose of this is to compare the Patient Petails
        'taken from the TXF file with the Patient Details in the
        'Drawing that were poked to the DDE text Boxes.
        '
        'Messages are displayed in the txtMessage Text Box
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.NL = Chr(13) & Chr(10) 'New Line  (this is global)
        TABBED = Chr(9)

        'Check if data is available from an existing drawing.
        'This is implied by txtUidMPD which contains the DRAFIX
        'UID for the "mainpatientdetails" symbol.  If this is empty
        'then there is no symbol, hence it is a NEW drawing and
        'the check does not need to be done

        If txtUidMPD.Text = "" Then Exit Sub

        'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PR_Message("Comparing TXF file details to Drawing." & CLTXFDIA1.NL)

        If Trim(UCase(txtFileNoPD.Text)) <> Trim(UCase(txtFileNo.Text)) Then
            PR_Message("File Numbers are different.")
            PR_Message(TABBED & "TXF file no :>" & Trim(UCase(txtFileNoPD.Text)) & "<")
            'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_Message(TABBED & "Dwg file no :>" & Trim(UCase(txtFileNo.Text)) & "<" & CLTXFDIA1.NL)
            Beep()
        End If

        If g_sSexTXF <> txtSex.Text Then
            PR_Message("Patient Sex is different.")
            PR_Message(TABBED & "TXF Sex is : " & g_sSexTXF)
            'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_Message(TABBED & "Dwg Sex was: " & txtSex.Text & CLTXFDIA1.NL)
            Beep()
        End If

        If Trim(UCase(cboDiagnosisPD.Text)) <> Trim(UCase(txtDiagnosis.Text)) Then
            PR_Message("The Diagnoisis has changed.")
            PR_Message(TABBED & "TXF Diagnosis is :>" & Trim(UCase(cboDiagnosisPD.Text)) & "<")
            'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_Message(TABBED & "Dwg Diagnosis was:>" & Trim(UCase(txtDiagnosis.Text)) & "<" & CLTXFDIA1.NL)
            Beep()
        End If

        If Val(txtAgePD.Text) <> Val(txtAge.Text) Then
            PR_Message("Patient Age has changed.")
            PR_Message(TABBED & "TXF Age is : " & txtAgePD.Text)
            'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_Message(TABBED & "Dwg Age was: " & txtAge.Text & CLTXFDIA1.NL)
            Beep()
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
        txtFileNo.Text = txtFileNoPD.Text

        txtPatientName.Text = txtNamePD.Text

        txtAge.Text = txtAgePD.Text

        txtSex.Text = cboSexPD.Text

        txtDiagnosis.Text = cboDiagnosisPD.Text


        'Work Order Details
        '
        txtWorkOrder.Text = txtWorkOrderWOD.Text

        txtMPDwo.Text = txtWorkOrderWOD.Text

        txtOrderDate.Text = txtOrderDateWOD.Text

        txtTemplateEngineer.Text = cboDesignerWOD.Text

        txtUnits.Text = cboUnitsWOD.Text

    End Sub

    Private Sub PR_UpdateExisting_Macro(ByRef sDrafixFile As String)
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
        Dim sLine As String

        'Open file
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.fNum = FreeFile()
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileOpen(CLTXFDIA1.fNum, sDrafixFile, VB.OpenMode.Output)

        'Initialise String globals
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.CC = Chr(44) 'The comma (,)
        'UPGRADE_WARNING: Couldn't resolve default property of object NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.NL = Chr(10) 'The new line character
        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.QQ = Chr(34) 'Double quotes (")
        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.QCQ = CLTXFDIA1.QQ & CLTXFDIA1.CC & CLTXFDIA1.QQ
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.QC = CLTXFDIA1.QQ & CLTXFDIA1.CC
        'UPGRADE_WARNING: Couldn't resolve default property of object QQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CLTXFDIA1.CQ = CLTXFDIA1.CC & CLTXFDIA1.QQ

        'Write header information etc. to the DRAFIX macro file
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "//DRAFIX Macro created - " & DateString & "  " & TimeString)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "//Patient    - " & txtPatientName.Text & CLTXFDIA1.CC & " " & txtFileNo.Text & CLTXFDIA1.CC)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "//Work Order - " & txtWorkOrder.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "//by Visual Basic")

        'Define DRAFIX variables
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "STRING sName, sFileNo, sPatient, sAge, sUnits, sSex, sWorkOrder, sTemplateEngineer, sOrderDate;")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "STRING sCADFile, sDiagnosis;")

        'Set Existing file name
        'sLine = FN_EscapeSlashesInString((lblCADFile.Text))
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "sCADFile= " & CLTXFDIA1.QQ & sLine & CLTXFDIA1.QQ & ";")


        'Set strings for use in update macro
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "sFileNo = " & CLTXFDIA1.QQ & txtFileNo.Text & CLTXFDIA1.QQ & ";")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "sPatient = " & CLTXFDIA1.QQ & txtPatientName.Text & CLTXFDIA1.QQ & ";")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "sAge = " & CLTXFDIA1.QQ & txtAge.Text & CLTXFDIA1.QQ & ";")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "sUnits =" & CLTXFDIA1.QQ & txtUnits.Text & CLTXFDIA1.QQ & ";")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "sSex = " & CLTXFDIA1.QQ & txtSex.Text & CLTXFDIA1.QQ & ";")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "sDiagnosis =" & CLTXFDIA1.QQ & txtDiagnosis.Text & CLTXFDIA1.QQ & ";")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "sWorkOrder =" & CLTXFDIA1.QQ & txtMPDwo.Text & CLTXFDIA1.QQ & ";")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "sOrderDate =" & CLTXFDIA1.QQ & txtOrderDate.Text & CLTXFDIA1.QQ & ";")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "sTemplateEngineer =" & CLTXFDIA1.QQ & txtTemplateEngineer.Text & CLTXFDIA1.QQ & ";")

        'Call the update macro
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CLTXFDIA1.fNum, "@" & CLTXFDIA1.g_sPathJOBST & "\CADLINK\CL_UPDTE.D;")

        'Close the Macro File
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FileClose(CLTXFDIA1.fNum)

    End Sub

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        'In case there is no Link_Close event
        'The programme will time out after approx 5 secs
        'End
    End Sub
    Private Sub saveInfoToDWG()
        Try
            Dim _sClass As New SurroundingClass()
            If (_sClass.GetXrecord("NewPatient", "NEWPATIENTDIC") IsNot Nothing) Then
                _sClass.RemoveXrecord("NewPatient", "NEWPATIENTDIC")
            End If

            Dim resbuf As New ResultBuffer
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtFileNoPD.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtNamePD.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtAgePD.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtWorkOrderWOD.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboSexPD.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboUnitsWOD.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboDesignerWOD.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboDiagnosisPD.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), DateTimePicker1.Text))

            _sClass.SetXrecord(resbuf, "NewPatient", "NEWPATIENTDIC")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub readDWGInfo()
        Try

            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
            resbuf = _sClass.GetXrecord("NewPatient", "NEWPATIENTDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                txtFileNoPD.Text = arr(0).Value
                txtNamePD.Text = arr(1).Value
                txtAgePD.Text = arr(2).Value
                txtWorkOrderWOD.Text = arr(3).Value
                cboSexPD.Text = arr(4).Value
                cboUnitsWOD.Text = arr(5).Value
                cboDesignerWOD.Text = arr(6).Value
                cboDiagnosisPD.Text = arr(7).Value
                DateTimePicker1.Text = arr(8).Value
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        ''CurrentYear - BirthDate  
        Dim startTime As DateTime = Convert.ToDateTime(DateTimePicker1.Value)
        Dim endTime As DateTime = DateTime.Today
        Dim span As TimeSpan = endTime.Subtract(startTime)
        Dim totalDays As Double = span.TotalDays
        Dim totalYears As Double = Math.Truncate(totalDays / 365)
        txtAgePD.Text = totalYears.ToString()
    End Sub
    Public Sub SetXrecord(ByVal resbuf As ResultBuffer, Optional ByVal dictName As String = "SNInfo", Optional ByVal key As String = "SNDIC")
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim NOD As DBDictionary = CType(tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead), DBDictionary)
            Dim dict As DBDictionary

            If NOD.Contains(dictName) Then
                dict = CType(tr.GetObject(NOD.GetAt(dictName), OpenMode.ForWrite), DBDictionary)
            Else
                dict = New DBDictionary()
                NOD.UpgradeOpen()
                NOD.SetAt(dictName, dict)
                tr.AddNewlyCreatedDBObject(dict, True)
            End If

            Dim xRec As Xrecord = New Xrecord()
            xRec.Data = resbuf
            dict.SetAt(key, xRec)
            tr.AddNewlyCreatedDBObject(xRec, True)
            tr.Commit()
        End Using
    End Sub

    Public Function GetXrecord(Optional ByVal dictName As String = "SNInfo", Optional ByVal key As String = "SNDIC") As ResultBuffer
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim NOD As DBDictionary = CType(tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead), DBDictionary)
            If Not NOD.Contains(dictName) Then Return Nothing
            Dim dict As DBDictionary = TryCast(tr.GetObject(NOD.GetAt(dictName), OpenMode.ForRead), DBDictionary)
            If dict Is Nothing OrElse Not dict.Contains(key) Then Return Nothing
            Dim xRec As Xrecord = TryCast(tr.GetObject(dict.GetAt(key), OpenMode.ForRead), Xrecord)
            If xRec Is Nothing Then Return Nothing
            Return xRec.Data
        End Using
    End Function
    Public Sub RemoveXrecord(Optional ByVal dictName As String = "SNInfo", Optional ByVal key As String = "SNDIC")
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim dl As DocumentLock = doc.LockDocument(DocumentLockMode.ProtectedAutoWrite, Nothing, Nothing, True)

        Using dl

            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim NOD As DBDictionary = CType(tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead), DBDictionary)
                If Not NOD.Contains(dictName) Then Return
                Dim dict As DBDictionary = TryCast(tr.GetObject(NOD.GetAt(dictName), OpenMode.ForWrite), DBDictionary)
                If dict Is Nothing OrElse Not dict.Contains(key) Then Return
                dict.Remove(key)
            End Using
        End Using
    End Sub

    Private Sub cltxfdia_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        saveInfoToDWG()
    End Sub
End Class