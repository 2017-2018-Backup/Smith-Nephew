Option Strict Off
Option Explicit On
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports VB = Microsoft.VisualBasic
Friend Class HeadNeck
    Inherits System.Windows.Forms.Form
    'Project:   HeadNeck
    'File:
    'Purpose:   To draw face masks and chin straps
    '
    'Version:   3.00
    'Date:      95
    'Author:    Cieran McCavanagh
    '           Gary George
    '
    '-------------------------------------------------------
    'REVISIONS:
    'Date       By      Action
    '-------------------------------------------------------
    'Dec 98     GG      Ported to VB5
    '
    'Notes:-


    ' Declare arrays
    'UPGRADE_WARNING: Lower bound of array FaceMaskChartA was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim FaceMaskChartA(3, 25) As Double
    'UPGRADE_WARNING: Lower bound of array FaceMaskChart2 was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim FaceMaskChart2(3, 16) As Double
    'UPGRADE_WARNING: Lower bound of array FaceMaskChart3 was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim FaceMaskChart3(3, 47) As Double

    'UPGRADE_WARNING: Lower bound of array TopArcAdj was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim TopArcAdj(19) As Double
    'UPGRADE_WARNING: Lower bound of array TopArcOpp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim TopArcOpp(19) As Double
    'UPGRADE_WARNING: Lower bound of array TopArcHyp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim TopArcHyp(19) As Double

    'UPGRADE_WARNING: Lower bound of array TopRightArcAdjIn was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim TopRightArcAdjIn(19) As Double
    'UPGRADE_WARNING: Lower bound of array TopRightArcAdjOut was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim TopRightArcAdjOut(19) As Double
    'UPGRADE_WARNING: Lower bound of array TopRightArcOppIn was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim TopRightArcOppIn(19) As Double
    'UPGRADE_WARNING: Lower bound of array TopRightArcOppOut was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim TopRightArcOppOut(19) As Double

    'UPGRADE_WARNING: Lower bound of array RightArcAdjUp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim RightArcAdjUp(20) As Double
    'UPGRADE_WARNING: Lower bound of array RightArcAdjDown was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim RightArcAdjDown(20) As Double
    'UPGRADE_WARNING: Lower bound of array RightArcOppUp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim RightArcOppUp(20) As Double
    'UPGRADE_WARNING: Lower bound of array RightArcOppDown was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim RightArcOppDown(20) As Double
    'UPGRADE_WARNING: Lower bound of array RightArcOpenAdjUp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim RightArcOpenAdjUp(20) As Double

    'UPGRADE_WARNING: Lower bound of array LeftArcAdjIn was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim LeftArcAdjIn(19) As Double
    'UPGRADE_WARNING: Lower bound of array LeftArcAdjOut was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim LeftArcAdjOut(19) As Double
    'UPGRADE_WARNING: Lower bound of array LeftArcOppIn was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim LeftArcOppIn(19) As Double
    'UPGRADE_WARNING: Lower bound of array LeftArcOppOut was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim LeftArcOppOut(19) As Double
    'UPGRADE_WARNING: Lower bound of array LeftArcHyp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim LeftArcHyp(19) As Double

    'UPGRADE_WARNING: Lower bound of array BotFaceAdj was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim BotFaceAdj(19) As Double
    'UPGRADE_WARNING: Lower bound of array BotFaceOpp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim BotFaceOpp(19) As Double

    'UPGRADE_WARNING: Lower bound of array NeckRightAdj was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim NeckRightAdj(19) As Double
    'UPGRADE_WARNING: Lower bound of array NeckRightOpp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim NeckRightOpp(19) As Double

    'UPGRADE_WARNING: Lower bound of array ChinProfileX was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim ChinProfileX(15) As Double
    'UPGRADE_WARNING: Lower bound of array ChinProfileY was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim ChinProfileY(15) As Double

    'UPGRADE_WARNING: Lower bound of array LipProfileX was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim LipProfileX(30) As Double
    'UPGRADE_WARNING: Lower bound of array LipProfileY was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim LipProfileY(30) As Double

    'UPGRADE_WARNING: Lower bound of array OpenFaceProfileX was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim OpenFaceProfileX(12) As Double
    'UPGRADE_WARNING: Lower bound of array OpenFaceProfileY was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim OpenFaceProfileY(12) As Double

    'UPGRADE_WARNING: Lower bound of array NoseTipLenAdj was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim NoseTipLenAdj(50) As Double
    'UPGRADE_WARNING: Lower bound of array NoseTipLenOpp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim NoseTipLenOpp(50) As Double
    'UPGRADE_WARNING: Lower bound of array NoseTipWidthAdj was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim NoseTipWidthAdj(50) As Double
    'UPGRADE_WARNING: Lower bound of array NoseTipWidthOpp was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim NoseTipWidthOpp(50) As Double

    'UPGRADE_WARNING: Lower bound of array ChinStrapFrontX was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim ChinStrapFrontX(8) As Double
    'UPGRADE_WARNING: Lower bound of array ChinStrapFrontY was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim ChinStrapFrontY(8) As Double
    'UPGRADE_WARNING: Lower bound of array ChinStrapBackX was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim ChinStrapBackX(19) As Double
    'UPGRADE_WARNING: Lower bound of array ChinStrapBackY was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim ChinStrapBackY(19) As Double
    'UPGRADE_WARNING: Lower bound of array ChinStrapChinX was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim ChinStrapChinX(13) As Double
    'UPGRADE_WARNING: Lower bound of array ChinStrapChinY was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim ChinStrapChinY(13) As Double
    'UPGRADE_WARNING: Lower bound of array ChinStrapLipX was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim ChinStrapLipX(25) As Double
    'UPGRADE_WARNING: Lower bound of array ChinStrapLipY was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim ChinStrapLipY(25) As Double
    Dim xyChinTopCen As HEADNECK1.xy
    Dim xyChinBotCen As HEADNECK1.xy
    Dim g_nChinBotRadius As Double
    Dim g_nChinTopRadius As Double
    Dim xyNoseDiagTop, xyNoseDiagBot As HEADNECK1.xy 'Used DrawNose & DrawEyeFlap
    Dim xyInsertion As HEADNECK1.xy 'Insertion point get from user
    Dim g_bIsClose As Boolean

    '    '* Windows API Functions Declarations
    '    Private Declare Function GetWindow Lib "User" (ByVal hwnd%, ByVal wCmd%) As Integer
    '    Private Declare Function GetWindowText Lib "User" (ByVal hwnd%, ByVal lpString$, ByVal aint%) As Integer
    '    Private Declare Function GetWindowTextLength Lib "User" (ByVal hwnd%) As Integer
    '    Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
    '    Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)

    'Constanst used by GetWindow
    '    Const GW_CHILD = 5
    '    Const GW_HWNDFIRST = 0
    '    Const GW_HWNDLAST = 1
    '    Const GW_HWNDNEXT = 2
    '    Const GW_HWNDPREV = 3
    '    Const GW_OWNER = 4

    Private Function Arccos(ByRef X As Double) As Double
        Arccos = System.Math.Atan(-X / System.Math.Sqrt(-X * X + 1)) + 1.5708
    End Function

    'UPGRADE_WARNING: Event chkLeftEarClosed.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkLeftEarClosed_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLeftEarClosed.CheckStateChanged

        ' if the left ear is closed, then
        ' the Left Ear Flap is disabled
        ' Else it is enabled

        If chkLeftEarClosed.CheckState = 1 Then
            chkLeftEarFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLeftEarFlap.Enabled = False
        ElseIf chkLeftEarClosed.CheckState = 0 Then
            chkLeftEarFlap.Enabled = True
        End If

    End Sub

    'UPGRADE_WARNING: Event chkLeftEarFlap.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkLeftEarFlap_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLeftEarFlap.CheckStateChanged

        ' if the Left Ear Flap is included, then
        ' the Left Ear Closed is disabled
        ' Else it is enabled

        If chkLeftEarFlap.CheckState = 1 Then
            chkLeftEarClosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLeftEarClosed.Enabled = False
        ElseIf chkLeftEarFlap.CheckState = 0 Then
            chkLeftEarClosed.Enabled = True
        End If

    End Sub

    'UPGRADE_WARNING: Event chkLining.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkLining_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLining.CheckStateChanged

        ' if Lining is Chosen, then the
        ' Open Head Mask check box is disabled
        ' else it is ensbled

        If chkLining.CheckState = 1 Then
            chkOpenHeadMask.Enabled = False
            chkOpenHeadMask.CheckState = System.Windows.Forms.CheckState.Unchecked
        ElseIf chkLining.CheckState = 0 Then
            chkOpenHeadMask.Enabled = True
        End If


    End Sub

    'UPGRADE_WARNING: Event chkLipStrap.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkLipStrap_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLipStrap.CheckStateChanged

        ' The Lip Covering is only available if there is a Lip Strap with the Following:
        ' an Open Face Mask, a Chin Strap, or a Modified Chin Strap
        ' If there is no lipstrap then the Lip Covering is disabled

        'GG-19.Sep.96
        '    If (optOpenFaceMask.Value = True Or optChinStrap.Value = True Or optModifiedChinStrap.Value = True) And chkLipStrap.Value = 1 Then
        '        chkLipCovering.Enabled = True
        '    ElseIf (optOpenFaceMask.Value = True Or optChinStrap.Value = True Or optModifiedChinStrap.Value = True) And chkLipStrap.Value = 0 Then
        '        chkLipCovering.Enabled = False
        '        chkLipCovering.Value = 0
        '    End If

    End Sub

    'UPGRADE_WARNING: Event chkOpenHeadMask.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkOpenHeadMask_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkOpenHeadMask.CheckStateChanged

        ' if an Open Head Mask is Chosen, then the
        ' Lining check box is disabled
        ' else it is ensbled

        If chkOpenHeadMask.CheckState = 1 Then
            chkLining.Enabled = False
            chkLining.CheckState = System.Windows.Forms.CheckState.Unchecked
        ElseIf chkOpenHeadMask.CheckState = 0 Then
            chkLining.Enabled = True
        End If

    End Sub

    'UPGRADE_WARNING: Event chkRightEarClosed.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkRightEarClosed_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkRightEarClosed.CheckStateChanged

        ' if the Right ear is closed, then
        ' the Right Ear Flap is disabled
        ' Else it is enabled

        If chkRightEarClosed.CheckState = 1 Then
            chkRightEarFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEarFlap.Enabled = False
        ElseIf chkRightEarClosed.CheckState = 0 Then
            chkRightEarFlap.Enabled = True
        End If

    End Sub

    'UPGRADE_WARNING: Event chkRightEarFlap.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkRightEarFlap_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkRightEarFlap.CheckStateChanged

        ' if the Right Ear Flap is included, then
        ' the Right Ear Closed is disabled
        ' Else it is enabled

        If chkRightEarFlap.CheckState = 1 Then
            chkRightEarClosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEarClosed.Enabled = False
        ElseIf chkRightEarFlap.CheckState = 0 Then
            chkRightEarClosed.Enabled = True
        End If

    End Sub

    'UPGRADE_WARNING: Event chkVelcro.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkVelcro_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkVelcro.CheckStateChanged

        If chkVelcro.CheckState = 1 Then
            chkZipper.Enabled = False
            chkZipper.CheckState = System.Windows.Forms.CheckState.Unchecked
        ElseIf chkVelcro.CheckState = 0 Then
            chkZipper.Enabled = True
        End If

    End Sub

    'UPGRADE_WARNING: Event chkZipper.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkZipper_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkZipper.CheckStateChanged

        If chkZipper.CheckState = 1 Then
            chkVelcro.Enabled = False
            chkVelcro.CheckState = System.Windows.Forms.CheckState.Unchecked
        ElseIf chkZipper.CheckState = 0 Then
            chkVelcro.Enabled = True
        End If

    End Sub

    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click

        ' This Procedure Compares the Values on the Form with
        ' the Values last saved.  If changes have been made, the
        ' option is given to save the changes before closing

        'System.Windows.Forms.Form.ActiveForm.Close()
        'Exit Sub

        Dim Response As Short
        Dim DataString, MeasurementString As String
        Dim OptionChoice, Options As String

        ' Check For any Changes
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetMeasurementString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        MeasurementString = FN_SetMeasurementString()

        'UPGRADE_WARNING: Couldn't resolve default property of object FN_ModificationString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        DataString = FN_ModificationString()
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_ChooseDesign(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Options = FN_ChooseDesign()
        DataString = Options & DataString
        If (DataString = txtData.Text And MeasurementString = txtMeasurements.Text) Or (txtData.Text = "txtData" And txtMeasurements.Text = "txtMeasurements") Then
            ' no changes
            'End
            ''Return
            g_bIsClose = True
            Me.Close()
        Else
            ' changes have been made, find out what to do?
            Response = MsgBox("Changes have been made, Save changes before closing", 35, "Head & Neck Details")

            ' Determine response of message box
            If Response = 6 Then
                ' Yes button selected
                txtData.Text = DataString
                txtMeasurements.Text = MeasurementString
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_OpenSave(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
                txtfNum.Text = FN_OpenSave(sDrawFile, txtPatientName.Text, txtFileNo.Text)
                PR_SaveDetails()
                FileClose(CInt(txtfNum.Text))
                'AppActivate(fnGetDrafixWindowTitleText())
                'System.Windows.Forms.SendKeys.SendWait("@" & sDrawFile & "{enter}")
                'End
                saveInfoToDWG()
                g_bIsClose = True
                Me.Close()
            ElseIf Response = 7 Then
                ' No button selected
                g_bIsClose = True
                Me.Close()
                'End
            ElseIf Response = 2 Then
                ' Cancel button selected
                Exit Sub
            End If

        End If

    End Sub

    Private Sub cmdDraw_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDraw.Click

        ' This Procedure Checks that the required values have been
        ' entered, It produces code for the required drawing, draws the
        ' design, and saves the results to the Head and Neck symbol

        Dim DataString, OptionChoice As String
        Dim HeadBand, NeckDepth As Double
        Dim CatalogueNo As String


        ' Check Correct Values have been entered for choices

        ' Check for Fabric type
        If cboFabric.Text = "" Then
            MsgBox("Fabric Type Has Not Been Entered", 48, "Head & Neck Details")
            cboFabric.Focus()
            Exit Sub
        End If

        ' Check for Extension Collar Values
        If optContouredChinCollar.Checked = True Or optChinCollar.Checked = True Then
            If Val(txtCircOfNeck.Text) <= 0 Then
                MsgBox("Invalid or No Measurement entered for CIRC of Neck", 48, "Head & Neck Details")
                txtCircOfNeck.Focus()
                Exit Sub
            End If
        End If

        ' Check Neck Depth
        Dim Response As Short
        If txtThroatToSternal.Text <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            NeckDepth = FN_CmToInches(txtUnits, txtThroatToSternal)
            If NeckDepth < 1.125 Then
                Response = MsgBox("Neck Depth is less than the specification minimum", 1, "Head & Neck Details")
                If Response = 2 Then
                    txtThroatToSternal.Focus()
                    Exit Sub
                End If
            End If
        End If

        ' Check For Face Mask Values
        If optFaceMask.Checked = True Or optOpenFaceMask.Checked = True Or optChinStrap.Checked = True Or optModifiedChinStrap.Checked = True Then
            If Val(txtChinToMouth.Text) <= 0 Then
                MsgBox("Invalid or No Measurement entered for Chin To Mouth", 48, "Head & Neck Details")
                txtChinToMouth.Focus()
                Exit Sub
            End If
            If Val(txtCircEyeBrow.Text) <= 0 Then
                MsgBox("Invalid or No Measurement entered for CIRC above Eyebrow", 48, "Head & Neck Details")
                txtCircEyeBrow.Focus()
                Exit Sub
            End If
            If Val(txtCircChinAngle.Text) <= 0 Then
                MsgBox("Invalid or No Measurement entered for CIRC of Head at Chin Angle", 48, "Head & Neck Details")
                txtCircChinAngle.Focus()
                Exit Sub
            End If
            If Val(txtCircOfNeck.Text) <= 0 Then
                MsgBox("Invalid or No Measurement entered for CIRC of Neck", 48, "Head & Neck Details")
                txtCircOfNeck.Focus()
                Exit Sub
            End If
            If chkNoseCovering.CheckState = 1 Then
                If Val(txtTipOfNose.Text) <= 0 Then
                    MsgBox("Invalid or No Measurement entered for Across Tip of Nose", 48, "Head & Neck Details")
                    txtTipOfNose.Focus()
                    Exit Sub
                End If
                If Val(txtLengthOfNose.Text) <= 0 Then
                    MsgBox("Invalid or No Measurement entered for Length of Nose", 48, "Head & Neck Details")
                    txtLengthOfNose.Focus()
                    Exit Sub
                End If
            End If
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_ChooseDesign(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtOptionChoice.Text = FN_ChooseDesign()

        ' Set Data String, the data string contains the design choice and modifications
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_ModificationString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        DataString = FN_ModificationString()
        txtData.Text = txtOptionChoice.Text & DataString

        ' Set Measurements String, the Measurements string contains the dimensions of the patients head
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetMeasurementString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtMeasurements.Text = FN_SetMeasurementString()

        ' Select the required template radius for patient
        If optContouredChinCollar.Checked = False And optChinCollar.Checked = False Then
            PR_SelectTemplateRadius()
        End If

        ' Check For HeadBand Values
        If optHeadBand.Checked = True Then
            If txtHeadBandDepth.Text = "" Then
                MsgBox("Invalid or No Measurement entered for Head Band Depth", 48, "Head & Neck Details")
                txtHeadBandDepth.Focus()
                Exit Sub
            End If
            If Val(txtCircEyeBrow.Text) <= 0 Then
                MsgBox("Invalid or No Measurement entered for CIRC above Eyebrow", 48, "Head & Neck Details")
                txtCircEyeBrow.Focus()
                Exit Sub
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            HeadBand = FN_CmToInches(txtUnits, txtHeadBandDepth)
            If HeadBand > TopArcAdj(CInt(txtRadiusNo.Text)) Then
                MsgBox("Head Band depth is too Large.", 48, "Head Band")
                Exit Sub
            End If
        End If


        If chkNoseCovering.CheckState = 1 Then
            ' Check Values have Been Entered
            If (Val(txtTipOfNose.Text) <= 0) Then
                MsgBox("No value has been entered for distance Across Tip of Nose.", 48, "Nose Covering")
                txtTipOfNose.Focus()
            End If
            If (Val(txtLengthOfNose.Text) <= 0) Then
                MsgBox("No value has been entered for the Length of Nose.", 48, "Nose Covering")
                txtLengthOfNose.Focus()
            End If
        End If


        'System.Windows.Forms.Form.ActiveForm.Close()
        Me.Hide()
        PR_GetInsertionPoint()
        ' Open file for drawing code
        Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
        txtfNum.Text = CStr(FN_Open(sDrawFile, txtPatientName, txtFileNo))
        PR_SetLayer("TemplateLeft")

        ' Draw relevant design option
        PR_FaceMask()
        PR_OpenFaceMask()
        PR_DrawHeadBand()
        PR_DrawChinStrap()
        PR_DrawChinCollar()

        ' Save details to HeadNeck symbol
        PR_SaveDetails()

        ' Close drawing code file
        FileClose(CInt(txtfNum.Text))
        saveInfoToDWG()
        g_bIsClose = True
        Me.Close()
        'Modified for AutoCAD 2018

        ' Activate drafix file and draw design
        'AppActivate(fnGetDrafixWindowTitleText())
        'System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
        'End
    End Sub

    Private Function FN_CalcAngle(ByRef xyStart As HEADNECK1.xy, ByRef xyEnd As HEADNECK1.xy) As Double

        'Function to return the angle between two points in degrees
        'in the range 0 - 360
        'Zero is always 0 and is never 360

        Dim X, Y As Object
        Dim rAngle As Double
        Dim PI As Double
        PI = 3.141592654

        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X = xyEnd.X - xyStart.X
        'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Y = xyEnd.Y - xyStart.Y

        'Horizomtal
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If X = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Y > 0 Then
                FN_CalcAngle = 90 * (PI / 180)
            Else
                FN_CalcAngle = 270 * (PI / 180)
            End If 
            Exit Function
        End If

        'Vertical (avoid divide by zero later)
        'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Y = 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If X > 0 Then
                FN_CalcAngle = 0 * (PI / 180)
            Else
                FN_CalcAngle = 180 * (PI / 180)
            End If
            Exit Function
        End If

        'All other cases
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        rAngle = System.Math.Atan(Y / X) * (180 / PI) 'Convert to degrees

        If rAngle < 0 Then rAngle = rAngle + 180 'rAngle range is -PI/2 & +PI/2

        'UPGRADE_WARNING: Couldn't resolve default property of object Y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Y > 0 Then
            FN_CalcAngle = rAngle
        Else
            FN_CalcAngle = rAngle + 180
        End If
        FN_CalcAngle = FN_CalcAngle * (PI / 180)

    End Function

    Private Function FN_CalcCirCirInt(ByRef xyCen1 As HEADNECK1.xy, ByRef nRad1 As Double, ByRef xyCen2 As HEADNECK1.xy, ByRef nRad2 As Double, ByRef xyInt1 As HEADNECK1.xy, ByRef xyInt2 As HEADNECK1.xy) As Short
        'Function that will return
        '    TRUE    if two circles intersect
        '    FALSE   if two circles don't intersect
        '
        'The intersection points are returned in the values
        '
        '    xyInt1 & xyInt2
        '
        'with intersection with lowest X value as xyInt1
        '
        Dim aTheta, nLength, aAngle, nCosTheta As Double
        Dim xyTmp As HEADNECK1.xy

        FN_CalcCirCirInt = False

        'Check that the circles can intersect
        'It is a theorem of plane geometry that no three real numbers
        'a, b and c can be the lenghts of the sides of a triangle unless
        'the sum of any two is greater than the third.
        'We use this as our main test of possible intersection, we also check for silly
        'data.

        'Test for silly data
        If xyCen1.X = xyCen2.X And xyCen1.Y = xyCen2.Y Then Exit Function
        If nRad1 <= 0 Or nRad2 <= 0 Then Exit Function

        'Test for intersection
        nLength = FN_CalcLength(xyCen1, xyCen2)

        If (nLength + nRad1 < nRad2) Or (nLength + nRad2 < nRad1) Or (nRad1 + nRad2 < nLength) Then
            Exit Function
        Else
            FN_CalcCirCirInt = True
        End If

        'Calculate intesection points
        '
        'Special case where circles touch (ie Intersect at one point only)


        'Angle between centers
        'Note: Length between centers from above
        aAngle = FN_CalcAngle(xyCen1, xyCen2) * (180 / HEADNECK1.PI)

        'Get angle w.r.t line between centers to the intersection point
        'use cosine rule
        '
        nCosTheta = -((nRad2 ^ 2 - (nLength ^ 2 + nRad1 ^ 2)) / (2 * nLength * nRad1))

        aTheta = System.Math.Atan(-nCosTheta / System.Math.Sqrt(-(nCosTheta ^ 2) + 1)) + 1.5708

        aTheta = aTheta * (180 / HEADNECK1.PI) 'convert to degrees

        aAngle = aAngle - aTheta
        PR_CalcPolar(xyCen1, aAngle, nRad1, xyInt1)

        aAngle = aAngle + aTheta
        PR_CalcPolar(xyCen1, aAngle, nRad1, xyInt2)

        If xyInt2.X < xyInt1.X Then
            'UPGRADE_WARNING: Couldn't resolve default property of object xyTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xyTmp = xyInt1
            'UPGRADE_WARNING: Couldn't resolve default property of object xyInt1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xyInt1 = xyInt2
            'UPGRADE_WARNING: Couldn't resolve default property of object xyInt2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xyInt2 = xyTmp
        End If


    End Function

    Private Function FN_CalcLength(ByRef xyStart As HEADNECK1.xy, ByRef xyEnd As HEADNECK1.xy) As Double
        'Fuction to return the length between two points
        'Greatfull thanks to Pythagorus

        FN_CalcLength = System.Math.Sqrt((xyEnd.X - xyStart.X) ^ 2 + (xyEnd.Y - xyStart.Y) ^ 2)

    End Function

    Private Function FN_ChooseDesign() As Object

        ' this function returns the code word for the chosen design

        If optFaceMask.Checked = True Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_ChooseDesign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_ChooseDesign = "RFM"
        ElseIf optHeadBand.Checked = True Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_ChooseDesign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_ChooseDesign = "RHB"
        ElseIf optOpenFaceMask.Checked = True Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_ChooseDesign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_ChooseDesign = "OFM"
        ElseIf optChinStrap.Checked = True Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_ChooseDesign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_ChooseDesign = "RCS"
        ElseIf optModifiedChinStrap.Checked = True Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_ChooseDesign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_ChooseDesign = "MCS"
        ElseIf optChinCollar.Checked = True Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_ChooseDesign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_ChooseDesign = "RCC"
        ElseIf optContouredChinCollar.Checked = True Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_ChooseDesign. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_ChooseDesign = "CCC"
        End If

    End Function

    Function FN_CirLinInt(ByRef xyStart As HEADNECK1.xy, ByRef xyEnd As HEADNECK1.xy, ByRef xyCen As HEADNECK1.xy, ByRef nRad As Double, ByRef xyInt As HEADNECK1.xy) As Short
        'Function to calculate the intersection between
        'a line and a circle.
        'Note:-
        '    Returns true if intersection found.
        '    The first intersection (only) is found.
        '    Ported from DRAFIX CAD DLG version.
        '

        Static nM, nC, nA, nSlope, nB, nK, nCalcTmp As Object
        Static nRoot As Double
        Static nSign As Short

        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Dim PI As Double = 3.141592654
        nSlope = FN_CalcAngle(xyStart, xyEnd) * (180 / PI)

        'Horizontal Line
        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nSlope = 0 Or nSlope = 180 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nSlope = -1
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nC = nRad ^ 2 - (xyStart.Y - xyCen.Y) ^ 2
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If nC < 0 Then
                FN_CirLinInt = False 'no roots
                Exit Function
            End If
            nSign = 1 'test each root
            While nSign > -2
                'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nRoot = xyCen.X + System.Math.Sqrt(nC) * nSign
                'UPGRADE_WARNING: Couldn't resolve default property of object max(xyStart.x, xyEnd.x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object min(xyStart.x, xyEnd.x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                    xyInt.X = nRoot
                    xyInt.Y = xyStart.Y
                    FN_CirLinInt = True
                    Exit Function
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False
            Exit Function
        End If

        'Vertical Line
        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nSlope = 90 Or nSlope = 270 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nSlope = -1
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nC = nRad ^ 2 - (xyStart.X - xyCen.X) ^ 2
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If nC < 0 Then
                FN_CirLinInt = False 'no roots
                Exit Function
            End If
            nSign = 1 'test each root
            While nSign > -2
                'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nRoot = xyCen.Y + System.Math.Sqrt(nC) * nSign
                'UPGRADE_WARNING: Couldn't resolve default property of object max(xyStart.y, xyEnd.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object min(xyStart.y, xyEnd.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If nRoot >= min(xyStart.Y, xyEnd.Y) And nRoot <= max(xyStart.Y, xyEnd.Y) Then
                    xyInt.Y = nRoot
                    xyInt.X = xyStart.X
                    FN_CirLinInt = True
                    Exit Function
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False
            Exit Function
        End If

        'Non-othogonal line
        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nSlope > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nM = (xyEnd.Y - xyStart.Y) / (xyEnd.X - xyStart.X) 'Slope
            'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nK = xyStart.Y - nM * xyStart.X 'Y-Axis intercept
            'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nA = (1 + nM ^ 2)
            'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nB = 2 * (-xyCen.X + (nM * nK) - (xyCen.Y * nM))
            'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nC = (xyCen.X ^ 2) + (nK ^ 2) + (xyCen.Y ^ 2) - (2 * xyCen.Y * nK) - (nRad ^ 2)
            'UPGRADE_WARNING: Couldn't resolve default property of object nA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nCalcTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nCalcTmp = (nB ^ 2) - (4 * nC * nA)

            'UPGRADE_WARNING: Couldn't resolve default property of object nCalcTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (nCalcTmp < 0) Then
                FN_CirLinInt = False 'No Roots
                Exit Function
            End If
            nSign = 1
            While nSign > -2
                'UPGRADE_WARNING: Couldn't resolve default property of object nA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object nCalcTmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object nB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nRoot = (-nB + (System.Math.Sqrt(nCalcTmp) / nSign)) / (2 * nA)
                'UPGRADE_WARNING: Couldn't resolve default property of object max(xyStart.x, xyEnd.x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object min(xyStart.x, xyEnd.x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If nRoot >= min(xyStart.X, xyEnd.X) And nRoot <= max(xyStart.X, xyEnd.X) Then
                    xyInt.X = nRoot
                    'UPGRADE_WARNING: Couldn't resolve default property of object nK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object nM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    xyInt.Y = nM * nRoot + nK
                    FN_CirLinInt = True
                    Exit Function 'Return first root found
                End If
                nSign = nSign - 2
            End While
            FN_CirLinInt = False 'Should never get to here
        End If
        FN_CirLinInt = False

    End Function

    Private Function FN_CmToInches(ByRef flag As System.Windows.Forms.Control, ByRef cm As System.Windows.Forms.Control) As Object

        ' This function converts cm to inches

        Dim temprem, remainder, inch, temp, rnder As Double
        Dim Value, realrem As Double

        'UPGRADE_WARNING: Couldn't resolve default property of object flag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If flag.Text = "cm" Then ' value is in cm
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inch = Val(cm.Text) / 2.54
            temp = Int(inch)
            remainder = (inch - temp)
            remainder = remainder / 0.125
            realrem = Int(remainder)
            temprem = remainder - Int(remainder)
            If temprem >= 0.5 Then
                realrem = realrem + 1
            End If
            remainder = realrem * 0.125
            inch = temp + remainder
            FN_CmToInches = System.Math.Abs(inch)
            'UPGRADE_WARNING: Couldn't resolve default property of object flag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf flag.Text = "inches" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Value = cm.Text
            If Value >= 1 Then
                temp = Int(Value)
                remainder = Value - temp
                If remainder > 0.75 Then
                    MsgBox("Invalid Measurements, Inches and Eights only", 48, "Head & Neck Details")
                    'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    FN_CmToInches = "-1"
                    Exit Function
                Else
                    remainder = remainder * 10
                    remainder = remainder * 0.125
                    FN_CmToInches = System.Math.Abs(temp + remainder)
                End If
            Else
                temp = Int(Value)
                remainder = Value - temp
                If remainder > 0.75 Then
                    MsgBox("Invalid Measurements, Inches and Eights only", 48, "Head & Neck Details")
                    'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    FN_CmToInches = "-1"
                    Exit Function
                Else
                    remainder = remainder * 10
                    remainder = remainder * 0.125
                    FN_CmToInches = System.Math.Abs(remainder)
                End If
            End If
        End If

    End Function


    Private Function FN_InchesToCm(ByRef flag As Object, ByRef inch As Object) As Object

        ' This function converts inches to cm

        Dim temprem, remainder, cm, temp, rnder As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object flag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If flag = "cm" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object inch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            cm = Val(inch) * 2.54
            'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            temp = Int(CDbl(cm))
            'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object remainder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            remainder = (cm - temp)
            'UPGRADE_WARNING: Couldn't resolve default property of object remainder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object temprem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            temprem = remainder * 10
            'UPGRADE_WARNING: Couldn't resolve default property of object temprem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            temprem = Int(CDbl(temprem))
            'UPGRADE_WARNING: Couldn't resolve default property of object temprem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object remainder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object rnder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            rnder = (remainder * 10) - temprem
            If rnder > 0.5 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object temprem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                temprem = temprem + 1
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object temprem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            temprem = temprem / 10
            'UPGRADE_WARNING: Couldn't resolve default property of object temprem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            cm = temp + temprem
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_InchesToCm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_InchesToCm = cm
            'UPGRADE_WARNING: Couldn't resolve default property of object flag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf flag = "inches" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            temp = Int(CDbl(inch))
            'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object inch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object remainder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            remainder = inch - temp
            'UPGRADE_WARNING: Couldn't resolve default property of object remainder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object rnder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            rnder = remainder / 0.125
            'UPGRADE_WARNING: Couldn't resolve default property of object rnder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            rnder = Int(CDbl(rnder))
            'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object rnder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            cm = (rnder / 10) + temp
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_InchesToCm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_InchesToCm = cm
        End If

    End Function

    Private Function FN_InchesToText(ByRef nInches As Double) As String

        'Function returns a decimal value in inches as a string
        '
        Dim nPrecision, nDec As Double
        Dim iInt, iEighths As Short
        Dim sString As String
        nPrecision = 0.125

        'Split into decimal parts
        iInt = Int(nInches)
        nDec = nInches - iInt
        If nDec <> 0 Then 'Avoid overflow
            iEighths = Int(nDec / nPrecision)
        Else
            iEighths = 0
        End If

        'Format string
        If iInt <> 0 Then
            sString = LTrim(Str(iInt))
        Else
            sString = "  "
        End If
        If iEighths <> 0 Then
            Select Case iEighths
                Case 2, 6
                    sString = sString & " " & LTrim(Str(iEighths / 2)) & "/4"
                Case 4
                    sString = sString & " " & "1/2"
                Case Else
                    sString = sString & " " & LTrim(Str(iEighths)) & "/8"
            End Select
        Else
            sString = sString
        End If

        'Return formatted string
        FN_InchesToText = sString

    End Function

    Function FN_LinLinInt(ByRef xyLine1Start As HEADNECK1.xy, ByRef xyLine1End As HEADNECK1.xy, ByRef xyLine2Start As HEADNECK1.xy, ByRef xyLine2End As HEADNECK1.xy, ByRef xyInt As HEADNECK1.xy) As Short
        'Function:
        '       BOOLEAN = FN_LinLinInt( xyLine1Start, xyLine1End, xyLine2Start, xyLine2End, xyInt);
        'Parameters:
        '       xyLine1Start = xyLine1Start.X, xyLine1Start.Y
        '       xyLine1End = xyLine1End.X, xyLine1End.Y
        '       xyLine2Start = xyLine2Start.X, xyLine2Start.Y
        '       xyLine2End = xyLine2End.X, xyLine2End.Y
        '
        'Returns:
        '       True if intersection found and lies on the line
        '       False if no intesection
        '       xyInt =  intersection
        '
        Dim nY, nSlope1, nK2, nK1, nM1, nCase, nX As Double
        Dim nM2, nSlope2 As Object

        'Initialy false
        FN_LinLinInt = False

        'Calculate slope of lines
        nCase = 0
        nSlope1 = FN_CalcAngle(xyLine1Start, xyLine1End)
        If nSlope1 = 0 Or nSlope1 = 180 Then nCase = nCase + 1
        If nSlope1 = 90 Or nSlope1 = 270 Then nCase = nCase + 2

        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nSlope2 = FN_CalcAngle(xyLine2Start, xyLine2End)
        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nSlope2 = 0 Or nSlope2 = 180 Then nCase = nCase + 4
        'UPGRADE_WARNING: Couldn't resolve default property of object nSlope2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nSlope2 = 90 Or nSlope2 = 270 Then nCase = nCase + 8

        Select Case nCase

            Case 0
                'Both lines are Non-Orthogonal Lines
                nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
                'UPGRADE_WARNING: Couldn't resolve default property of object nM2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nM2 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
                'UPGRADE_WARNING: Couldn't resolve default property of object nM2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If (nM1 = nM2) Then Exit Function 'Parallel lines
                nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
                'UPGRADE_WARNING: Couldn't resolve default property of object nM2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nK2 = xyLine2Start.Y - (nM2 * xyLine2Start.X) 'Y-Axis intercept
                If (nK1 = nK2) Then Exit Function
                'Find X
                'UPGRADE_WARNING: Couldn't resolve default property of object nM2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                nX = (nK2 - nK1) / (nM1 - nM2)
                'Find Y
                nY = (nM1 * nX) + nK1

            Case 1
                'Line 1 is Horizontal or Line 2 is horizontal
                nM1 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
                nK1 = xyLine2Start.Y - (nM1 * xyLine2Start.X) 'Y-Axis intercept
                nY = xyLine1Start.Y
                'Solve for X at the given Y value
                nX = (nY - nK1) / nM1

            Case 2
                'Line 1 is Vertical or Line 2 is Vertical
                nM1 = (xyLine2End.Y - xyLine2Start.Y) / (xyLine2End.X - xyLine2Start.X) 'Slope
                nK1 = xyLine2Start.Y - (nM1 * xyLine2Start.X) 'Y-Axis intercept
                nX = xyLine1Start.X
                'Solve for Y at the given X value
                nY = (nM1 * nX) + nK1

            Case 4
                'Line 1 is Horizontal or Line 2 is horizontal
                nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
                nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
                nY = xyLine2Start.Y

                'Solve for X at the given Y value
                nX = (nY - nK1) / nM1

            Case 5
                'Parallel orthogonal lines, no intersection possible
                Exit Function

            Case 6
                'Line1 is Vertical and the Line2 is Horizontal
                nX = xyLine1Start.X
                nY = xyLine2Start.Y

            Case 8
                'Line 1 is Vertical or Line 2 is Vertical
                nM1 = (xyLine1End.Y - xyLine1Start.Y) / (xyLine1End.X - xyLine1Start.X) 'Slope
                nK1 = xyLine1Start.Y - (nM1 * xyLine1Start.X) 'Y-Axis intercept
                nX = xyLine2Start.X
                'Solve for Y at the given X value
                nY = (nM1 * nX) + nK1

            Case 9
                'Line1 is Horizontal and the Line2 is Vertical
                nX = xyLine2Start.X
                nY = xyLine1Start.Y

            Case 10
                'Parallel orthogonal lines, no intersection possible
                Exit Function

            Case Else
                Exit Function

        End Select

        'Ensure that the points X and Y are on the lines
        xyInt.X = nX
        xyInt.Y = nY

        'Line 1
        'UPGRADE_WARNING: Couldn't resolve default property of object max(xyLine1Start.x, xyLine1End.x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object min(xyLine1Start.x, xyLine1End.x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (nX < min(xyLine1Start.X, xyLine1End.X) Or nX > max(xyLine1Start.X, xyLine1End.X)) Then Exit Function
        'UPGRADE_WARNING: Couldn't resolve default property of object max(xyLine1Start.y, xyLine1End.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object min(xyLine1Start.y, xyLine1End.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (nY < min(xyLine1Start.Y, xyLine1End.Y) Or nY > max(xyLine1Start.Y, xyLine1End.Y)) Then Exit Function

        'Line 2
        'UPGRADE_WARNING: Couldn't resolve default property of object max(xyLine2Start.x, xyLine2End.x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object min(xyLine2Start.x, xyLine2End.x). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (nX < min(xyLine2Start.X, xyLine2End.X) Or nX > max(xyLine2Start.X, xyLine2End.X)) Then Exit Function
        'UPGRADE_WARNING: Couldn't resolve default property of object max(xyLine2Start.y, xyLine2End.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object min(xyLine2Start.y, xyLine2End.y). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (nY < min(xyLine2Start.Y, xyLine2End.Y) Or nY > max(xyLine2Start.Y, xyLine2End.Y)) Then Exit Function

        FN_LinLinInt = True

    End Function

    Private Function FN_ModificationString() As Object

        ' This Function sets a string containing the chosen modifications

        Dim Modifications As Object 'As String

        ' Set Data String
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = chkLeftEarFlap.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkRightEarFlap.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkNeckElastic.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkLeftEyeFlap.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkRightEyeFlap.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkOpenHeadMask.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkLipStrap.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkLipCovering.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkNoseCovering.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkZipper.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkLining.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkEarSize.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkRightEarClosed.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkLeftEarClosed.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkEyes.CheckState
        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Modifications = Modifications & chkVelcro.CheckState

        'UPGRADE_WARNING: Couldn't resolve default property of object Modifications. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_ModificationString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FN_ModificationString = Modifications

    End Function

    Private Function FN_Open(ByRef sDrafixFile As String, ByRef sName As System.Windows.Forms.Control, ByRef sPatientFile As System.Windows.Forms.Control) As Short

        'Initialise String globals
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HEADNECK1.CC = Chr(44) 'The comma (,)
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HEADNECK1.NL = Chr(10) 'The new line character
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HEADNECK1.QQ = Chr(34) 'Double quotes (")
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HEADNECK1.QCQ = HEADNECK1.QQ & HEADNECK1.CC & HEADNECK1.QQ
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HEADNECK1.QC = HEADNECK1.QQ & HEADNECK1.CC
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HEADNECK1.CQ = HEADNECK1.CC & HEADNECK1.QQ
        '
        'Open the DRAFIX macro file
        'Initialise Global variables

        'Open file
        txtfNum.Text = CStr(FreeFile())
        FileOpen(CInt(txtfNum.Text), sDrafixFile, VB.OpenMode.Output)
        FN_Open = CShort(txtfNum.Text)

        'Initialise patient globals
        'UPGRADE_WARNING: Couldn't resolve default property of object sPatientFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtFileNo.Text = sPatientFile.Text
        'UPGRADE_WARNING: Couldn't resolve default property of object sName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtPatientName.Text = sName.Text

        'Globals to reduced drafix code written to file
        txtCurrentLayer.Text = ""
        txtCurrTextHt.Text = CStr(0.125)
        txtCurrTextAspect.Text = CStr(0.6)
        txtCurrTextHorizJust.Text = CStr(1) 'Left
        txtCurrTextVertJust.Text = CStr(32) 'Bottom
        txtCurrTextFont.Text = CStr(0) 'BLOCK

        'Write header information etc. to the DRAFIX macro file

        PrintLine(CInt(txtfNum.Text), "//DRAFIX Macro created - " & DateString & "  " & TimeString)
        PrintLine(CInt(txtfNum.Text), "//Patient - " & txtPatientName.Text & HEADNECK1.CC & " " & txtFileNo.Text & HEADNECK1.CC & " Head & Neck ")
        PrintLine(CInt(txtfNum.Text), "//by Visual Basic")

        'Define DRAFIX variables
        PrintLine(CInt(txtfNum.Text), "HANDLE hLayer, hFacemask, hTitle, hChan, hEnt;")
        PrintLine(CInt(txtfNum.Text), "HEADNECK1.XY     xyTitleBoxOrigin,xyTitleOrigin, xyStart,TitleOrigin, xyTitleScale, xyOrigin;")
        PrintLine(CInt(txtfNum.Text), "ANGLE  aTitleAngle;")
        PrintLine(CInt(txtfNum.Text), "STRING sTitleName, sFileNo;")

        'Set up data base fields
        PrintLine(CInt(txtfNum.Text), "Table(" & HEADNECK1.QQ & "add" & HEADNECK1.QCQ & "field" & HEADNECK1.QCQ & "ID" & HEADNECK1.QCQ & "string" & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "Table(" & HEADNECK1.QQ & "add" & HEADNECK1.QCQ & "field" & HEADNECK1.QCQ & "HeadNeck" & HEADNECK1.QCQ & "string" & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "Table(" & HEADNECK1.QQ & "add" & HEADNECK1.QCQ & "field" & HEADNECK1.QCQ & "units" & HEADNECK1.QCQ & "string" & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "Table(" & HEADNECK1.QQ & "add" & HEADNECK1.QCQ & "field" & HEADNECK1.QCQ & "Data" & HEADNECK1.QCQ & "string" & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "Table(" & HEADNECK1.QQ & "add" & HEADNECK1.QCQ & "field" & HEADNECK1.QCQ & "WorkOrder" & HEADNECK1.QCQ & "string" & HEADNECK1.QQ & ");")

        'Text data
        PrintLine(CInt(txtfNum.Text), "SetData(" & HEADNECK1.QQ & "TextHorzJust" & HEADNECK1.QC & txtCurrTextHorizJust.Text & ");")
        PrintLine(CInt(txtfNum.Text), "SetData(" & HEADNECK1.QQ & "TextVertJust" & HEADNECK1.QC & txtCurrTextVertJust.Text & ");")
        PrintLine(CInt(txtfNum.Text), "SetData(" & HEADNECK1.QQ & "TextHeight" & HEADNECK1.QC & txtCurrTextHt.Text & ");")
        PrintLine(CInt(txtfNum.Text), "SetData(" & HEADNECK1.QQ & "TextAspect" & HEADNECK1.QC & txtCurrTextAspect.Text & ");")
        PrintLine(CInt(txtfNum.Text), "SetData(" & HEADNECK1.QQ & "TextFont" & HEADNECK1.QC & txtCurrTextFont.Text & ");")

        'Get Start point
        PrintLine(CInt(txtfNum.Text), "GetUser (" & HEADNECK1.QQ & "xy" & HEADNECK1.QCQ & "Indicate Start Point" & HEADNECK1.QC & "&xyStart);")

        'Display Hour Glass symbol
        PrintLine(CInt(txtfNum.Text), "Display (" & HEADNECK1.QQ & "cursor" & HEADNECK1.QCQ & "wait" & HEADNECK1.QCQ & "Drawing" & HEADNECK1.QQ & ");")

        'Clear user selections etc
        PrintLine(CInt(txtfNum.Text), "UserSelection (" & HEADNECK1.QQ & "clear" & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "Execute (" & HEADNECK1.QQ & "menu" & HEADNECK1.QCQ & "SetStyle" & HEADNECK1.QC & "Table(" & HEADNECK1.QQ & "find" & HEADNECK1.QCQ & "style" & HEADNECK1.QCQ & "bylayer" & HEADNECK1.QQ & "));")

        'Place a marker at the start point later use
        PR_SetLayer("Construct")
        PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "marker" & HEADNECK1.QCQ & "cross" & HEADNECK1.QC & "xyStart" & HEADNECK1.CC & "0.125);")
        PR_AddEntityID("OriginMarker")
        ' Update data base fields
        PrintLine(CInt(txtfNum.Text), "SetDBData(hEnt," & HEADNECK1.QQ & "Fabric" & HEADNECK1.QCQ & cboFabric.Text & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "SetDBData(hEnt," & HEADNECK1.QQ & "WorkOrder" & HEADNECK1.QCQ & txtWorkOrder.Text & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "SetDBData(hEnt," & HEADNECK1.QQ & "Data" & HEADNECK1.QCQ & txtData.Text & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "SetDBData(hEnt," & HEADNECK1.QQ & "units" & HEADNECK1.QCQ & txtUnits.Text & HEADNECK1.QQ & ");")

        'Set values for use futher on by other macros
        PrintLine(CInt(txtfNum.Text), "xyOrigin = xyStart" & ";")
        PrintLine(CInt(txtfNum.Text), "sFileNo = " & HEADNECK1.QQ & txtFileNo.Text & HEADNECK1.QQ & ";")

    End Function


    Private Function FN_OpenSave(ByRef sDrafixFile As String, ByRef sName As Object, ByRef sPatientFile As Object) As Object

        'Initialise String globals
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HEADNECK1.CC = Chr(44) 'The comma (,)
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.NL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HEADNECK1.NL = Chr(10) 'The new line character
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HEADNECK1.QQ = Chr(34) 'Double quotes (")
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HEADNECK1.QCQ = HEADNECK1.QQ & HEADNECK1.CC & HEADNECK1.QQ
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HEADNECK1.QC = HEADNECK1.QQ & HEADNECK1.CC
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HEADNECK1.CQ = HEADNECK1.CC & HEADNECK1.QQ
        '
        'Open the DRAFIX macro file
        'Initialise Global variables

        'Open file
        txtfNum.Text = CStr(FreeFile())
        FileOpen(CInt(txtfNum.Text), sDrafixFile, VB.OpenMode.Output)
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_OpenSave. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FN_OpenSave = txtfNum.Text

        'Initialise patient globals
        'UPGRADE_WARNING: Couldn't resolve default property of object sPatientFile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtFileNo.Text = sPatientFile
        'UPGRADE_WARNING: Couldn't resolve default property of object sName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtPatientName.Text = sName

        'Globals to reduced drafix code written to file
        txtCurrentLayer.Text = ""
        txtCurrTextHt.Text = CStr(0.125)
        txtCurrTextAspect.Text = CStr(0.6)
        txtCurrTextHorizJust.Text = CStr(1) 'Left
        txtCurrTextVertJust.Text = CStr(32) 'Bottom
        txtCurrTextFont.Text = CStr(0) 'BLOCK

        'Write header information etc. to the DRAFIX macro file

        PrintLine(CInt(txtfNum.Text), "//DRAFIX Macro created - " & DateString & "  " & TimeString)
        PrintLine(CInt(txtfNum.Text), "//Patient - " & txtPatientName.Text & HEADNECK1.CC & " " & txtFileNo.Text & HEADNECK1.CC & " Head & Neck ")
        PrintLine(CInt(txtfNum.Text), "//by Visual Basic")

        'Define DRAFIX variables
        PrintLine(CInt(txtfNum.Text), "HANDLE hLayer, hFacemask, hTitle, hChan, hEnt;")
        PrintLine(CInt(txtfNum.Text), "HEADNECK1.XY     xyTitleBoxOrigin,xyTitleOrigin, xyStart,TitleOrigin, xyTitleScale, xyOrigin;")
        PrintLine(CInt(txtfNum.Text), "ANGLE  aTitleAngle;")
        PrintLine(CInt(txtfNum.Text), "STRING sTitleName, sFileNo;")

        'Set up ID  data base field
        ' Print #txtfNum, "Table("; HEADNECK1.qq; "add"; HEADNECK1.QCQ; "field"; HEADNECK1.QCQ; "ID"; HEADNECK1.QCQ; "string"; HEADNECK1.qq; ");"

        'Text data
        PrintLine(CInt(txtfNum.Text), "SetData(" & HEADNECK1.QQ & "TextHorzJust" & HEADNECK1.QC & txtCurrTextHorizJust.Text & ");")
        PrintLine(CInt(txtfNum.Text), "SetData(" & HEADNECK1.QQ & "TextVertJust" & HEADNECK1.QC & txtCurrTextVertJust.Text & ");")
        PrintLine(CInt(txtfNum.Text), "SetData(" & HEADNECK1.QQ & "TextHeight" & HEADNECK1.QC & txtCurrTextHt.Text & ");")
        PrintLine(CInt(txtfNum.Text), "SetData(" & HEADNECK1.QQ & "TextAspect" & HEADNECK1.QC & txtCurrTextAspect.Text & ");")
        PrintLine(CInt(txtfNum.Text), "SetData(" & HEADNECK1.QQ & "TextFont" & HEADNECK1.QC & txtCurrTextFont.Text & ");")



        'Clear user selections etc
        PrintLine(CInt(txtfNum.Text), "UserSelection (" & HEADNECK1.QQ & "clear" & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "UserSelection (" & HEADNECK1.QQ & "update" & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "Execute (" & HEADNECK1.QQ & "menu" & HEADNECK1.QCQ & "SetStyle" & HEADNECK1.QC & "Table(" & HEADNECK1.QQ & "find" & HEADNECK1.QCQ & "style" & HEADNECK1.QCQ & "bylayer" & HEADNECK1.QQ & "));")


        'Set values for use futher on by other macros
        PrintLine(CInt(txtfNum.Text), "sFileNo = " & HEADNECK1.QQ & txtFileNo.Text & HEADNECK1.QQ & ";")


    End Function


    Private Function FN_SetDetailField(ByRef Field As Object) As Object

        ' This function sets the detail field saved
        ' with the drawing

        Dim bit, pos, dig, Result As Object

        Dim tempval As Double
        If Len(Field) = 4 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Field. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            pos = InStr(Field, ".")
            'UPGRADE_WARNING: Couldn't resolve default property of object pos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If pos <> 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dig. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dig = Int(CDbl(Field))
                'UPGRADE_WARNING: Couldn't resolve default property of object Field. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object bit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                bit = VB.Right(Field, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object bit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dig. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Result = dig & bit
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_SetDetailField = Result
        ElseIf Len(Field) = 3 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Field. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            pos = InStr(Field, ".")
            'UPGRADE_WARNING: Couldn't resolve default property of object pos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If pos <> 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dig. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dig = Int(CDbl(Field))
                'UPGRADE_WARNING: Couldn't resolve default property of object Field. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object bit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                bit = VB.Right(Field, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object bit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dig. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Result = "0" & dig & bit
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_SetDetailField = Result
        ElseIf Len(Field) = 2 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Field. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            tempval = Val(Field)
            If tempval > 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Field. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Result = Field
                'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FN_SetDetailField = Result & "0"
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object Field. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Result = VB.Right(Field, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FN_SetDetailField = "00" & Result
            End If
        ElseIf Len(Field) = 1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Field. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Result = Field
            'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_SetDetailField = "0" & Result & "0"
            'UPGRADE_WARNING: Couldn't resolve default property of object Field. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf Field = "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_SetDetailField = "000"
        End If

    End Function

    Private Function FN_SetMeasurementString() As Object

        ' This Function Constructs a String to represent the Measurements

        Dim Rec4, Rec2, Rec1, Rec3, Rec5 As Object
        Dim Rec10, Rec8, Rec6, Rec7, Rec9, Rec11 As Object

        ' Set Measurements of the Face Mask
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rec1 = FN_SetDetailField(txtChinToMouth.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rec2 = FN_SetDetailField(txtCircEyeBrow.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rec3 = FN_SetDetailField(txtCircChinAngle.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rec4 = FN_SetDetailField(txtCircOfNeck.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rec5 = FN_SetDetailField(txtThroatToSternal.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec6. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rec6 = FN_SetDetailField(txtTipOfNose.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec7. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rec7 = FN_SetDetailField(txtLengthOfNose.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec8. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rec8 = FN_SetDetailField(txtHeadBandDepth.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec9. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rec9 = FN_SetDetailField(txtLeftEarLength.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec10. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rec10 = FN_SetDetailField(txtRightEarLength.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetDetailField(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec11. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Rec11 = FN_SetDetailField(txtChinCollarMin.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec11. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec10. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec9. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec8. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec7. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec6. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Rec1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_SetMeasurementString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        FN_SetMeasurementString = Rec1 & Rec2 & Rec3 & Rec4 & Rec5 & Rec6 & Rec7 & Rec8 & Rec9 & Rec10 & Rec11

    End Function

    Private Function fnRoundInches(ByVal nNumber As Double) As Double
        'Function to return the rounded value in decimal inches
        'returns to the nearest eighth (0.125)
        'E.G.
        '    5.67         = 5 inches and 0.67 inches
        '                   0.67 / 0.125 = 5.36 eighths
        '                   5.36 eighths = 5 eighths (rounded to nearest eighth)
        '    5.67         = 5 inches and 5 eighths
        '    5.67         = 5 + ( 5 * 0.125)
        '    5.67         = 6.625 inches
        '

        Dim iInt, iSign As Short
        Dim nPrecision, nDec As Double

        'Return 0 if input is Zero
        If nNumber = 0 Then
            fnRoundInches = 0
            Exit Function
        End If

        'Set precision
        nPrecision = 0.125

        'Break input into components
        iSign = System.Math.Sign(nNumber)
        nNumber = System.Math.Abs(nNumber)
        iInt = Int(nNumber)
        nDec = nNumber - iInt

        'Get decimal part in precision units
        If nDec <> 0 Then
            nDec = nDec / nPrecision 'Avoid overflow
        End If
        nDec = round(nDec)

        'Return value
        fnRoundInches = (iInt + (nDec * nPrecision)) * iSign

    End Function

    'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
    Private Sub Form_LinkClose()
        'Disable timer
        'Stop the timer used to ensure that the Dialogue dies
        'if the DRAFIX macro fails to establish a DDE Link
        Timer1.Enabled = False

        Show()
        ' Check for previous instance of HeadNeck, if thereis Unload
        'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'

        '/////-----------------------Check Later-=-------------------------------------------------------------
        'If App.PrevInstance Then
        '    Beep()
        '    MsgBox("The drawing routine for the Head & Neck is already running.", 48, "Head & Neck")
        '    Me.Close()
        '    'End
        '    Return
        'End If

        ' Check for Patient details
        If txtUidMPD.Text = "" Then
            MsgBox("Patient Details have not been entered.", 16, "Head & Neck")
            g_bIsClose = True
            Me.Close()
            'End
            Return
        End If

        cboFabric.Text = txtFabric.Text

        ' get measurements from txtMeasurements, and place in relevant text boxes
        PR_SeperateMeasurements()

        ' get Modifications from txtData, and set check boxes
        PR_SeperateModifications()

        Show()

        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default ' set mouse pointer to arrow

    End Sub

    Private Sub HeadNeck_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            ' this procedure loads all the required information
            ' to display the form HEADNECK

            ' Maintain while loading DDE data
            Hide()
            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' change pointer to hourglass
            ' Reset in form_linkclose

            ' Position to center of screen
            Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
            Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)

            HEADNECK1.MainForm = Me

            Dim textline As String
            Dim filenum As Short
            g_bIsClose = False

            ' Load Template Arrays, these arrays hold mappings of the
            ' 51508 R1 template, with measurements from the centre
            PR_LoadTopArcArrays()
            PR_LoadTopRightArrays()
            PR_LoadRightArcArrays()
            PR_LoadLeftArcArrays()
            PR_LoadBotFaceArrays()
            PR_LoadNeckArrays()

            ' Load Face Mask Chart Arrays, taken from specifications
            PR_LoadFaceMaskChart1()
            PR_LoadFaceMaskChart2()
            PR_LoadFaceMaskChart3()

            txtUidMPD.Text = ""
            txtUidHN.Text = ""

            HEADNECK1.g_sPathJOBST = fnPathJOBST()
            ' Load Fabric File
            filenum = FreeFile()
            Dim sSettingsPath As String = fnGetSettingsPath("LookupTables")
            FileOpen(filenum, sSettingsPath & "\fabric.dat", VB.OpenMode.Input)
            Do While Not EOF(filenum)
                textline = LineInput(filenum)
                cboFabric.Items.Add(textline)
            Loop
            FileClose()

            'Start a timer
            'The Timer is disabled in LinkClose
            'If after 10 seconds the timer event will "End" the programme
            'This ensures that the dialogue dies in event of a failure
            'on the drafix macro side
            Timer1.Interval = 10000 'Approx 10 Seconds
            Timer1.Enabled = True
            Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
            Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
            Dim blkId As ObjectId = New ObjectId()
            Dim obj As New BlockCreation.BlockCreation
            blkId = obj.LoadBlockInstance()
            If (blkId.IsNull()) Then
                MsgBox("Patient Details have not been entered", 48, "Head & Neck Details")
                g_bIsClose = True
                Me.Close()
                Exit Sub
            End If
            obj.BindAttributes(blkId, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units)
            If fileNo = "" Then
                MsgBox("Please enter Patient Details", 48, "Head & Neck Details Dialog")
                g_bIsClose = True
                Me.Close()
                Exit Sub
            End If
            txtPatientName.Text = patient
            txtFileNo.Text = fileNo
            txtDiagnosis.Text = diagnosis
            txtSex.Text = sex
            txtAge.Text = age
            txtUnits.Text = units
            txtWorkOrder.Text = workOrder

            txtWorkOrder1.Text = workOrder
            txtDesigner.Text = tempEng
            txtTempDate.Text = tempDate
            DateTimePicker1.Value = System.DateTime.Now
            readDWGInfo()
            txtMeasurements.Text = FN_SetMeasurementString()
            txtOptionChoice.Text = FN_ChooseDesign()
            Dim DataString As String = FN_ModificationString()
            txtData.Text = txtOptionChoice.Text & DataString
        Catch ex As Exception
            g_bIsClose = True
            Me.Close()
        End Try
    End Sub

    Private Function max(ByRef nFirst As Object, ByRef nSecond As Object) As Object
        ' Returns the maximum of two numbers
        'UPGRADE_WARNING: Couldn't resolve default property of object nSecond. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nFirst >= nSecond Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object max. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            max = nFirst
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object nSecond. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object max. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            max = nSecond
        End If
    End Function

    Private Function min(ByRef nFirst As Object, ByRef nSecond As Object) As Object
        ' Returns the minimum of two numbers
        'UPGRADE_WARNING: Couldn't resolve default property of object nSecond. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If nFirst <= nSecond Then
            'UPGRADE_WARNING: Couldn't resolve default property of object nFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            min = nFirst
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object nSecond. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object min. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            min = nSecond
        End If
    End Function

    'UPGRADE_WARNING: Event optChinCollar.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optChinCollar_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optChinCollar.CheckedChanged
        If eventSender.Checked Then

            ' Initalise available Modifications for Chin Collar

            chkLeftEarFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEarFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkNeckElastic.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLeftEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLeftEarClosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEarClosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkOpenHeadMask.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLipStrap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLipCovering.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkNoseCovering.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkZipper.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLining.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEarSize.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEyes.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkVelcro.CheckState = System.Windows.Forms.CheckState.Unchecked

            chkEyes.Enabled = False
            chkLeftEarFlap.Enabled = False
            chkRightEarFlap.Enabled = False
            chkNeckElastic.Enabled = False
            chkLeftEyeFlap.Enabled = False
            chkRightEyeFlap.Enabled = False
            chkLeftEarClosed.Enabled = False
            chkRightEarClosed.Enabled = False
            chkOpenHeadMask.Enabled = False
            chkLipStrap.Enabled = False
            chkLipCovering.Enabled = False
            chkNoseCovering.Enabled = False
            chkZipper.Enabled = False
            chkLining.Enabled = False
            chkEarSize.Enabled = False
            chkVelcro.Enabled = False

        End If
    End Sub

    'UPGRADE_WARNING: Event optChinStrap.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optChinStrap_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optChinStrap.CheckedChanged
        If eventSender.Checked Then

            ' Initalise available Modifications for Chin Strap

            chkLeftEarFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEarFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLeftEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkOpenHeadMask.CheckState = System.Windows.Forms.CheckState.Unchecked
            '    If chkLipStrap.Value = 0 Then
            chkLipCovering.CheckState = System.Windows.Forms.CheckState.Unchecked
            '    End If
            chkNoseCovering.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkZipper.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLining.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEarSize.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLeftEarClosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEarClosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEyes.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkVelcro.CheckState = System.Windows.Forms.CheckState.Unchecked

            chkEyes.Enabled = False
            chkLeftEarFlap.Enabled = False
            chkRightEarFlap.Enabled = False
            chkNeckElastic.Enabled = True
            chkLeftEyeFlap.Enabled = False
            chkRightEyeFlap.Enabled = False
            chkLeftEarClosed.Enabled = False
            chkRightEarClosed.Enabled = False
            chkOpenHeadMask.Enabled = False
            chkLipStrap.Enabled = True
            '   If chkLipStrap.Value = 0 Then
            '       chkLipCovering.Enabled = False
            '   Else
            chkLipCovering.Enabled = True
            '   End If
            chkNoseCovering.Enabled = False
            chkZipper.Enabled = False
            chkLining.Enabled = False
            chkEarSize.Enabled = False
            chkVelcro.Enabled = False

        End If
    End Sub

    'UPGRADE_WARNING: Event optContouredChinCollar.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optContouredChinCollar_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optContouredChinCollar.CheckedChanged
        If eventSender.Checked Then

            ' Initalise available Modifications for Contoured Chin Collar

            chkLeftEarFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEarFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkNeckElastic.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLeftEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLeftEarClosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEarClosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkOpenHeadMask.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLipStrap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLipCovering.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkNoseCovering.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkZipper.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLining.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEarSize.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEyes.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkVelcro.CheckState = System.Windows.Forms.CheckState.Unchecked

            chkEyes.Enabled = False
            chkLeftEarFlap.Enabled = False
            chkRightEarFlap.Enabled = False
            chkNeckElastic.Enabled = False
            chkLeftEyeFlap.Enabled = False
            chkRightEyeFlap.Enabled = False
            chkLeftEarClosed.Enabled = False
            chkRightEarClosed.Enabled = False
            chkOpenHeadMask.Enabled = False
            chkLipStrap.Enabled = False
            chkLipCovering.Enabled = False
            chkNoseCovering.Enabled = False
            chkZipper.Enabled = False
            chkLining.Enabled = False
            chkEarSize.Enabled = False
            chkVelcro.Enabled = False

        End If
    End Sub

    'UPGRADE_WARNING: Event optFaceMask.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optFaceMask_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optFaceMask.CheckedChanged
        If eventSender.Checked Then

            ' Initalise available Modifications for Face Mask

            chkLipStrap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEyes.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLipCovering.CheckState = System.Windows.Forms.CheckState.Unchecked

            chkEyes.Enabled = False
            chkLeftEarFlap.Enabled = True
            chkRightEarFlap.Enabled = True
            chkNeckElastic.Enabled = True
            chkLeftEyeFlap.Enabled = True
            chkRightEyeFlap.Enabled = True
            chkLeftEarClosed.Enabled = True
            chkRightEarClosed.Enabled = True
            chkOpenHeadMask.Enabled = True
            chkLipStrap.Enabled = False
            chkLipCovering.Enabled = True
            chkNoseCovering.Enabled = True
            chkZipper.Enabled = True
            chkLining.Enabled = True
            chkEarSize.Enabled = True
            chkVelcro.Enabled = True

        End If
    End Sub

    'UPGRADE_WARNING: Event optHeadBand.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optHeadBand_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optHeadBand.CheckedChanged
        If eventSender.Checked Then

            ' Initalise available Modifications for Head Band

            chkLeftEarFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEarFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkNeckElastic.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLeftEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLeftEarClosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEarClosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkOpenHeadMask.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLipStrap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLipCovering.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkNoseCovering.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkZipper.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLining.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEarSize.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEyes.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkVelcro.CheckState = System.Windows.Forms.CheckState.Unchecked

            chkEyes.Enabled = False
            chkLeftEarFlap.Enabled = False
            chkRightEarFlap.Enabled = False
            chkNeckElastic.Enabled = False
            chkLeftEyeFlap.Enabled = False
            chkRightEyeFlap.Enabled = False
            chkLeftEarClosed.Enabled = False
            chkRightEarClosed.Enabled = False
            chkOpenHeadMask.Enabled = False
            chkLipStrap.Enabled = False
            chkLipCovering.Enabled = False
            chkNoseCovering.Enabled = False
            chkZipper.Enabled = False
            chkLining.Enabled = False
            chkEarSize.Enabled = False
            chkVelcro.Enabled = False

        End If
    End Sub

    'UPGRADE_WARNING: Event optModifiedChinStrap.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optModifiedChinStrap_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optModifiedChinStrap.CheckedChanged
        If eventSender.Checked Then

            ' Initalise available Modifications for Modified Chin Strap

            chkLeftEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkOpenHeadMask.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkNoseCovering.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkZipper.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLining.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEyes.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkVelcro.CheckState = System.Windows.Forms.CheckState.Unchecked

            chkEyes.Enabled = False
            chkNeckElastic.Enabled = True
            chkLeftEyeFlap.Enabled = False
            chkRightEyeFlap.Enabled = False
            chkLeftEarFlap.Enabled = True
            chkRightEarFlap.Enabled = True
            chkLeftEarClosed.Enabled = True
            chkRightEarClosed.Enabled = True
            chkOpenHeadMask.Enabled = False
            chkLipStrap.Enabled = True
            '    If chkLipStrap.Value = 0 Then
            '        chkLipCovering.Enabled = False
            '    Else
            chkLipCovering.Enabled = True
            '    End If
            chkNoseCovering.Enabled = False
            chkZipper.Enabled = False
            chkLining.Enabled = False
            chkEarSize.Enabled = True
            chkVelcro.Enabled = False

        End If
    End Sub

    'UPGRADE_WARNING: Event optOpenFaceMask.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optOpenFaceMask_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optOpenFaceMask.CheckedChanged
        If eventSender.Checked Then

            ' Initalise available Modifications for Open Face Mask

            chkLeftEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkRightEyeFlap.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkLipCovering.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkNoseCovering.CheckState = System.Windows.Forms.CheckState.Unchecked

            chkEyes.Enabled = True
            chkLeftEarFlap.Enabled = True
            chkRightEarFlap.Enabled = True
            chkNeckElastic.Enabled = True
            chkLeftEyeFlap.Enabled = False
            chkRightEyeFlap.Enabled = False
            chkLeftEarClosed.Enabled = True
            chkRightEarClosed.Enabled = True
            chkOpenHeadMask.Enabled = True
            chkLipStrap.Enabled = True
            chkLipCovering.Enabled = False
            chkNoseCovering.Enabled = False
            chkZipper.Enabled = True
            chkLining.Enabled = True
            chkEarSize.Enabled = True
            chkVelcro.Enabled = True

        End If
    End Sub

    Private Sub PR_AddEntityID(ByRef sType As String)

        'To the DRAFIX macro file (given by the global fNum)
        'write the syntax to add to an ENTITY the database information
        'in the DB variable "ID" that will allow the identity of an entity
        'to be retrieved, by other parts of the system.
        '
        'For this to work it assumes that the following DRAFIX variables
        'are defined.
        '    HANDLE  hEnt
        '
        'Note:-
        '    fNum, CC, HEADNECK1.qq, HEADNECK1.NL are globals initialised by FN_Open
        '
        Dim sID, sData As String

        sID = txtOptionChoice.Text & txtFileNo.Text & sType
        sData = txtData.Text

        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CInt(txtfNum.Text), "if (hEnt) SetDBData( hEnt," & HEADNECK1.QQ & "ID" & HEADNECK1.QQ & HEADNECK1.CC & HEADNECK1.QQ & sID & HEADNECK1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(CInt(txtfNum.Text), "if (hEnt) SetDBData( hEnt," & HEADNECK1.QQ & "Data" & HEADNECK1.QQ & HEADNECK1.CC & HEADNECK1.QQ & sData & HEADNECK1.QQ & ");")

    End Sub

    Private Sub PR_CalcPolar(ByRef xyStart As HEADNECK1.xy, ByVal nAngle As Double, ByRef nLength As Double, ByRef xyReturn As HEADNECK1.xy)
        'Procedure to return a point at a distance and an angle from a given point
        '
        Dim a, B As Double

        Select Case nAngle
            Case 0
                xyReturn.X = xyStart.X + nLength
                xyReturn.Y = xyStart.Y
            Case 180
                xyReturn.X = xyStart.X - nLength
                xyReturn.Y = xyStart.Y
            Case 90
                xyReturn.X = xyStart.X
                xyReturn.Y = xyStart.Y + nLength
            Case 270
                xyReturn.X = xyStart.X
                xyReturn.Y = xyStart.Y - nLength
            Case Else
                'Convert from degees to radians
                nAngle = nAngle * HEADNECK1.PI / 180
                B = System.Math.Sin(nAngle) * nLength
                a = System.Math.Cos(nAngle) * nLength
                xyReturn.X = xyStart.X + a
                xyReturn.Y = xyStart.Y + B
        End Select
    End Sub

    Private Sub PR_ChinStrapDimensions()

        ' this procedure sets dimensions for the chinstrap

        'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtCSChintoMouth.Text = FN_CmToInches(txtUnits, txtChinToMouth)
        txtCSNeckCircum.Text = CStr((FN_CmToInches(txtUnits, txtCircOfNeck)) * 0.46)
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(txtUnits, txtCircChinAngle). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtCSChinAngle.Text = CStr((FN_CmToInches(txtUnits, txtCircEyeBrow) + FN_CmToInches(txtUnits, txtCircChinAngle)) * 0.19)
        txtCSForeHead.Text = CStr(CDbl(txtCSChinAngle.Text) / 2)
        txtCSMouthWidth.Text = CStr(CDbl(txtCSChinAngle.Text) / 8)
        txtCSEarTopHeight.Text = CStr(CDbl(txtCSForeHead.Text) - 1.125)
        txtCSEarBotHeight.Text = CStr(CDbl(txtCSChintoMouth.Text) * 0.72)

    End Sub

    Private Sub PR_DrawArc(ByRef xyCen As HEADNECK1.xy, ByRef xyArcStart As HEADNECK1.xy, ByRef xyArcEnd As HEADNECK1.xy)

        ' this procedure draws an arc between two points

        Dim nEndAng, nRad, nStartAng, nDeltaAng As Object


        'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nRad = FN_CalcLength(xyCen, xyArcStart)
        'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nStartAng = FN_CalcAngle(xyCen, xyArcStart)
        'UPGRADE_WARNING: Couldn't resolve default property of object nEndAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nEndAng = FN_CalcAngle(xyCen, xyArcEnd)
        'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nEndAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nDeltaAng = nEndAng - nStartAng
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)

            '' Create an arc that is at 6.25,9.125 with a radius of 6, and
            '' starts at 64 degrees and ends at 204 degrees
            Using acArc As Arc = New Arc(New Point3d(xyInsertion.X + xyCen.X, xyInsertion.Y + xyCen.Y, 0),
                                         nRad, nStartAng, nEndAng)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acArc.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If

                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acArc)
                acTrans.AddNewlyCreatedDBObject(acArc, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
            End Using

            'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "arc" & HEADNECK1.QC & "xyStart.x +" & Str(xyCen.X) & HEADNECK1.CC & "xyStart.y +" & Str(xyCen.Y) & HEADNECK1.CC & Str(nRad) & HEADNECK1.CC & Str(nStartAng) & HEADNECK1.CC & Str(nDeltaAng) & ");")

    End Sub

    Private Sub PR_DrawChin()

        ' This Procedure Forms The Mouth, Jobst Procedure No.9

        Dim xyMouthLeft, xyMouthRight As HEADNECK1.xy
        Dim xy4, xy2, xy1, xy3, xyTmp As HEADNECK1.xy
        Dim xyTopCen, xyBotCen, xyChinStart As HEADNECK1.xy
        Dim n2, n1, n3 As Double
        Dim nBotRadius, rAngle, aAngle, aInc, nValue, nTopRadius As Double
        Dim DefaultNeckDepth, GivenNeckDepth, nNeckOffset As Double

        Dim ChinStep2, TempY, Tempx, ChinDist, ChinStep1 As Double
        Dim PI, MouthWidth, Beta As Double
        Dim BotCurve, TopCurve, xyChinArc As HEADNECK1.xy
        Dim DiagLen, XA, YA, DiagStep As Double
        Dim TopRadius, BotRadius As Double
        Dim ii As Short
        Dim DiagStart, DiagEnd As HEADNECK1.xy
        Dim Delta As Double
        Dim v As Short
        Dim ChinHt, Theta As Double
        Dim i, j As Short

        TopRadius = 1.8 / 2.54
        BotRadius = 0.9 / 2.54
        PI = 3.141592654

        ' Check chart2 for mouth width
        For i = 1 To 16
            If FaceMaskChart2(1, i) = Val(txtCircumferenceTotal.Text) Then
                MouthWidth = FaceMaskChart2(2, i)
            End If
        Next i

        ' Set Cordinates and draw Mouth
        If (optOpenFaceMask.Checked = True) And chkLipStrap.CheckState = 0 Then
            Tempx = -(LeftArcAdjOut(CInt(txtRadiusNo.Text))) + 0.375
        Else
            Tempx = -(LeftArcAdjOut(CInt(txtRadiusNo.Text)))
        End If
        TempY = Val(txtMouthHeight.Text)

        PR_MakeXY(xyMouthLeft, Tempx, TempY)

        Tempx = -(System.Math.Abs(LeftArcAdjOut(CInt(txtRadiusNo.Text))) - MouthWidth)
        PR_MakeXY(xyMouthRight, Tempx, TempY)
        txtMouthRightX.Text = CStr(xyMouthRight.X)
        txtMouthRightY.Text = CStr(xyMouthRight.Y)
        If (optOpenFaceMask.Checked = True) And chkEyes.CheckState = 1 Then
            PR_MakeXY(xyMouthRight, Tempx - 0.125, TempY)
        End If

        PR_DrawLine(xyMouthLeft, xyMouthRight)
        PR_AddEntityID("Mouth")


        'Bug fix code added G George 2.Jul.96
        'Construction points
        nNeckOffset = 1
        If Val(txtAge.Text) > 10 Then
            DefaultNeckDepth = 2
        Else
            DefaultNeckDepth = 1.5
        End If

        If txtThroatToSternal.Text <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            GivenNeckDepth = FN_CmToInches(txtUnits, txtThroatToSternal)
            If GivenNeckDepth > DefaultNeckDepth Then
                nNeckOffset = nNeckOffset + System.Math.Abs(GivenNeckDepth - DefaultNeckDepth)
            End If
        End If

        PR_MakeXY(xyChinStart, Val(txtChinLeftBotX.Text), Val(txtChinLeftBotY.Text))
        PR_MakeXY(xy1, xyChinStart.X, xyChinStart.Y + (nNeckOffset - 1))
        PR_MakeXY(xy2, xy1.X, xyChinStart.Y + nNeckOffset)
        PR_MakeXY(xy4, -System.Math.Abs(LeftArcAdjOut(CInt(txtRadiusNo.Text))) + 0.375, Val(txtMouthHeight.Text))
        PR_MakeXY(xy3, xy4.X, xy4.Y - 1)

        'Fit curve to three lines defined by the 4 points above
        n1 = FN_CalcLength(xy1, xy2)
        n2 = FN_CalcLength(xy2, xy3)
        n3 = FN_CalcLength(xy3, xy4)
        nValue = FN_CalcLength(xy1, xy3)
        nValue = ((n1 ^ 2 + n2 ^ 2) - nValue ^ 2) / (2 * n1 * n2)
        rAngle = Arccos(nValue) / 2
        nBotRadius = System.Math.Tan(rAngle) * (2 * (n1 / 3))
        nTopRadius = System.Math.Tan(rAngle) * (n3 / 2)

        ChinProfileX(1) = xy1.X
        ChinProfileY(1) = xy1.Y

        ChinProfileX(2) = xy1.X
        ChinProfileY(2) = xy1.Y + (n1 / 3)

        PR_CalcPolar(xy2, (FN_CalcAngle(xy2, xy3) * (180 / PI)), (2 * (n1 / 3)), xyTmp)
        ChinProfileX(5) = xyTmp.X
        ChinProfileY(5) = xyTmp.Y

        PR_MakeXY(xyTmp, ChinProfileX(2), ChinProfileY(2))
        PR_CalcPolar(xyTmp, 180, nBotRadius, xyBotCen)
        PR_MakeXY(xyTmp, ChinProfileX(5), ChinProfileY(5))
        aInc = (FN_CalcAngle(xyBotCen, xyTmp) * (180 / PI)) / 3

        For ii = 3 To 4
            PR_CalcPolar(xyBotCen, aInc, nBotRadius, xyTmp)
            aInc = aInc + aInc
            ChinProfileX(ii) = xyTmp.X
            ChinProfileY(ii) = xyTmp.Y
        Next ii

        PR_CalcPolar(xy3, (FN_CalcAngle(xy3, xy2) * (180 / PI)), (n3 / 2), xyTmp)
        ChinProfileX(6) = xyTmp.X
        ChinProfileY(6) = xyTmp.Y

        PR_CalcPolar(xyTmp, (FN_CalcAngle(xy3, xy2) * (180 / PI)) + 90, nTopRadius, xyTopCen)
        aAngle = FN_CalcAngle(xyTopCen, xyTmp) * (180 / PI)
        aInc = (aAngle - 180) / 3

        For ii = 7 To 9
            aAngle = aAngle - aInc
            PR_CalcPolar(xyTopCen, aAngle, nTopRadius, xyTmp)
            ChinProfileX(ii) = xyTmp.X
            ChinProfileY(ii) = xyTmp.Y
        Next ii

        ChinProfileX(10) = xy4.X
        ChinProfileY(10) = xy4.Y

        PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "fitted" & HEADNECK1.QQ)
        PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyChinStart.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyChinStart.Y))
        For ii = 1 To 10
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(ChinProfileX(ii)) & HEADNECK1.CC & "xyStart.y+" & Str(ChinProfileY(ii)))
        Next ii
        PrintLine(CInt(txtfNum.Text), ");")
        ''To Draw Spline
        Dim ptColl As Point3dCollection = New Point3dCollection()
        ptColl.Add(New Point3d(xyInsertion.X + xyChinStart.X, xyInsertion.Y + xyChinStart.Y, 0))
        For ii = 1 To 10
            ptColl.Add(New Point3d(xyInsertion.X + ChinProfileX(ii), xyInsertion.Y + ChinProfileY(ii), 0))
        Next
        'ptColl.Add(New Point3d(xyInsertion.X + ChinProfileX(10), xyInsertion.Y + ChinProfileY(10), 0))
        PR_DrawSpline(ptColl)
        PR_AddEntityID("Chin")

        'Set center point for subsequent use in lip covering
        'UPGRADE_WARNING: Couldn't resolve default property of object xyChinBotCen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xyChinBotCen = xyBotCen
        'UPGRADE_WARNING: Couldn't resolve default property of object xyChinTopCen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xyChinTopCen = xyTopCen
        g_nChinBotRadius = nBotRadius
        g_nChinTopRadius = nTopRadius

    End Sub


    Private Sub PR_DrawChinCollar()

        ' this Procedure calculates and draws the Chin Extension Collar

        Dim xyTextPoint As HEADNECK1.xy
        Dim xyBotInnerLeft, xyBotInnerRight As HEADNECK1.xy
        Dim xyTopInnerLeft, xyTopInnerRight As HEADNECK1.xy
        Dim xyBotOuterLeft, xyBotOuterRight As HEADNECK1.xy
        Dim xyTopOuterLeft, xyTopOuterRight As HEADNECK1.xy
        Dim xyTopInnerContour, xyBotInnerContour As HEADNECK1.xy
        Dim xyTopOuterContour, xyBotOuterContour As HEADNECK1.xy
        Dim NeckCircum, NeckDepth, ContourMin As Double
        Dim OuterXDist, InnerXDist, YDist As Double
        Dim YContourStart, InnerXContourDist, OuterXContourDist, YContourEnd As Double
        Dim IntVal As Short
        Dim DecVal As Double

        If optChinCollar.Checked = True Or optContouredChinCollar.Checked = True Then

            'Round Up to nearest 1/2"
            If txtCircOfNeck.Text <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                NeckCircum = FN_CmToInches(txtUnits, txtCircOfNeck)
                IntVal = Int(NeckCircum)
                DecVal = NeckCircum - IntVal
                If DecVal > 0 And DecVal <= 0.5 Then
                    NeckCircum = IntVal + 0.5
                ElseIf DecVal > 0.5 Then
                    NeckCircum = IntVal + 1
                End If
            End If

            'Round Up to nearest 1/2"
            If txtThroatToSternal.Text <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                NeckDepth = FN_CmToInches(txtUnits, txtThroatToSternal)
                IntVal = Int(NeckDepth)
                DecVal = NeckDepth - IntVal
                If DecVal > 0 And DecVal <= 0.5 Then
                    NeckDepth = IntVal + 0.5
                ElseIf DecVal > 0.5 Then
                    NeckDepth = IntVal + 1
                End If
            End If

            'Round Up to nearest 1/2"
            If txtChinCollarMin.Text <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ContourMin = FN_CmToInches(txtUnits, txtChinCollarMin)
                IntVal = Int(ContourMin)
                DecVal = ContourMin - IntVal
                If DecVal > 0 And DecVal <= 0.5 Then
                    ContourMin = IntVal + 0.5
                ElseIf DecVal > 0.5 Then
                    ContourMin = IntVal + 1
                End If
                '            If ContourMin < 2.5 Then ContourMin = 2.5
                If ContourMin < 1 Then ContourMin = 1
            Else
                ContourMin = NeckDepth - 0.5
                If ContourMin < 2.5 Then ContourMin = 2.5
            End If

            ' Set Horizontal Distances
            If NeckCircum <= 8 Then
                InnerXDist = 6.1875
                OuterXDist = 7
            Else
                InnerXDist = (((NeckCircum - 8) * 9) / 14) + 6.1875
                OuterXDist = (((NeckCircum - 8) * 3) / 4) + 7
            End If

            ' Set Contour Horizontal Distances
            If NeckCircum <= 8 Then
                InnerXContourDist = 2.0625
                OuterXContourDist = 2.3125
            Else
                InnerXContourDist = ((NeckCircum - 8) / 4.618) + 2.0625
                OuterXContourDist = ((NeckCircum - 8) / 4) + 2.3125
            End If

            ' Set Vertical Distances
            If NeckDepth <= 0 Then
                If optContouredChinCollar.Checked = True Then
                    YDist = 3
                Else
                    YDist = 3.5
                End If
                '        ElseIf NeckDepth > 0 And NeckDepth <= 2.5 Then
                '            YDist = 3     'chg
            Else
                YDist = NeckDepth + 0.5
            End If

            If ContourMin > 0 Then
                ContourMin = ContourMin + 0.5
                YContourStart = (YDist - ContourMin) / 2
                YContourEnd = YDist - ((YDist - ContourMin) / 2)
            End If

            ' Set Chin Collar Points
            PR_MakeXY(xyBotInnerLeft, 0, 0)
            PR_MakeXY(xyTopInnerLeft, 0, YDist)
            PR_MakeXY(xyTopOuterLeft, 0, 0)
            PR_MakeXY(xyBotOuterLeft, 0, -YDist)

            If optContouredChinCollar.Checked = True Then
                PR_MakeXY(xyBotInnerContour, InnerXContourDist, 0)
                PR_MakeXY(xyBotInnerRight, InnerXDist, YContourStart)
                PR_MakeXY(xyTopInnerRight, InnerXDist, YContourEnd)
                PR_MakeXY(xyTopInnerContour, InnerXContourDist, YDist)
                PR_MakeXY(xyTopOuterContour, OuterXContourDist, 0)
                PR_MakeXY(xyTopOuterRight, OuterXDist, -YContourStart)
                PR_MakeXY(xyBotOuterRight, OuterXDist, -YContourEnd)
                PR_MakeXY(xyBotOuterContour, OuterXContourDist, -YDist)
            Else
                PR_MakeXY(xyBotInnerRight, InnerXDist, 0)
                PR_MakeXY(xyTopInnerRight, InnerXDist, YDist)
                PR_MakeXY(xyTopOuterRight, OuterXDist, 0)
                PR_MakeXY(xyBotOuterRight, OuterXDist, -YDist)
            End If

            ' Draw Inner Collar
            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "polyline" & HEADNECK1.QQ)
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyBotInnerLeft.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyBotInnerLeft.Y))
            If optContouredChinCollar.Checked = True Then
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyBotInnerContour.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyBotInnerContour.Y))
            End If
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyBotInnerRight.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyBotInnerRight.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyTopInnerRight.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyTopInnerRight.Y))
            If optContouredChinCollar.Checked = True Then
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyTopInnerContour.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyTopInnerContour.Y))
            End If
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyTopInnerLeft.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyTopInnerLeft.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyBotInnerLeft.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyBotInnerLeft.Y))
            PrintLine(CInt(txtfNum.Text), ");")
            ''To Draw Polyline
            Dim nCount As Integer = 5
            If optContouredChinCollar.Checked = True Then
                nCount = 7
            End If
            Dim PolyStart(nCount), PolyEnd(nCount) As Double
            PolyStart(1) = xyBotInnerLeft.X
            PolyEnd(1) = xyBotInnerLeft.Y
            Dim nVal As Integer = 2
            If optContouredChinCollar.Checked = True Then
                PolyStart(2) = xyBotInnerContour.X
                PolyEnd(2) = xyBotInnerContour.Y
                nVal = nVal + 1
            End If
            PolyStart(nVal) = xyBotInnerRight.X
            PolyEnd(nVal) = xyBotInnerRight.Y
            nVal = nVal + 1
            PolyStart(nVal) = xyTopInnerRight.X
            PolyEnd(nVal) = xyTopInnerRight.Y
            nVal = nVal + 1
            If optContouredChinCollar.Checked = True Then
                PolyStart(nVal) = xyTopInnerContour.X
                PolyEnd(nVal) = xyTopInnerContour.Y
                nVal = nVal + 1
            End If
            PolyStart(nVal) = xyTopInnerLeft.X
            PolyEnd(nVal) = xyTopInnerLeft.Y
            nVal = nVal + 1
            PolyStart(nVal) = xyBotInnerLeft.X
            PolyEnd(nVal) = xyBotInnerLeft.Y
            PR_DrawPoly(PolyStart, PolyEnd, nCount)
            PR_AddEntityID("InnerChinCollar")

            ' Draw Outer Collar
            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "polyline" & HEADNECK1.QQ)
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyTopOuterLeft.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyTopOuterLeft.Y))
            If optContouredChinCollar.Checked = True Then
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyTopOuterContour.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyTopOuterContour.Y))
            End If
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyTopOuterRight.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyTopOuterRight.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyBotOuterRight.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyBotOuterRight.Y))
            If optContouredChinCollar.Checked = True Then
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyBotOuterContour.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyBotOuterContour.Y))
            End If
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyBotOuterLeft.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyBotOuterLeft.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyTopOuterLeft.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyTopOuterLeft.Y))
            PrintLine(CInt(txtfNum.Text), ");")

            ''To Draw Polyline
            nVal = 2
            PolyStart(1) = xyTopOuterLeft.X
            PolyEnd(1) = xyTopOuterLeft.Y
            If optContouredChinCollar.Checked = True Then
                PolyStart(2) = xyTopOuterContour.X
                PolyEnd(2) = xyTopOuterContour.Y
                nVal = nVal + 1
            End If
            PolyStart(nVal) = xyTopOuterRight.X
            PolyEnd(nVal) = xyTopOuterRight.Y
            nVal = nVal + 1
            PolyStart(nVal) = xyBotOuterRight.X
            PolyEnd(nVal) = xyBotOuterRight.Y
            nVal = nVal + 1
            If optContouredChinCollar.Checked = True Then
                PolyStart(nVal) = xyBotOuterContour.X
                PolyEnd(nVal) = xyBotOuterContour.Y
                nVal = nVal + 1
            End If
            PolyStart(nVal) = xyBotOuterLeft.X
            PolyEnd(nVal) = xyBotOuterLeft.Y
            nVal = nVal + 1
            PolyStart(nVal) = xyTopOuterLeft.X
            PolyEnd(nVal) = xyTopOuterLeft.Y
            PR_DrawPoly(PolyStart, PolyEnd, nCount)

            PR_AddEntityID("OuterChinCollar")

            PR_MakeXY(xyTextPoint, xyTopInnerLeft.X + 1, xyTopInnerLeft.Y - 0.5)
            PR_TemplateDetails(xyTextPoint, "INNER Chin Collar")
            PR_MakeXY(xyTextPoint, xyBotInnerLeft.X + 0.35, xyBotInnerLeft.Y + 1.75)
            PR_SetLayer("Notes")
            PR_DrawText("Fold", xyTextPoint, 0.25, (90 * HEADNECK1.PI / 180))

            PR_MakeXY(xyTextPoint, xyTopOuterLeft.X + 1, xyTopOuterLeft.Y - 0.5)
            PR_TemplateDetails(xyTextPoint, "OUTER Chin Collar")
            PR_MakeXY(xyTextPoint, xyBotOuterLeft.X + 0.35, xyBotOuterLeft.Y + 1.75)
            PR_SetLayer("Notes")
            PR_DrawText("Fold", xyTextPoint, 0.25, (90 * HEADNECK1.PI / 180))
            PR_SetLayer("TemplateLeft")

        End If

    End Sub

    Private Sub PR_DrawChinStrap()

        Dim xyNeckTopFront, xyNeckTopBack As HEADNECK1.xy
        Dim xyNeckBottomFront, xyNeckBottomBack As HEADNECK1.xy
        Dim xyNeckTopVelcro, xyNeckBottomVelcro As HEADNECK1.xy
        Dim xyChinStrapTopFront, xyChinStrapTopBack As HEADNECK1.xy
        Dim xyFacelineTop, xyFacelineBot As HEADNECK1.xy
        Dim xyBacklineTop, xyBacklineBot As HEADNECK1.xy
        Dim xyMouthFront, xyMouthBack As HEADNECK1.xy
        Dim xyLipStrapTop, xyLipStrapBottom As HEADNECK1.xy
        Dim xyStrapText, xyTextPoint As HEADNECK1.xy
        Dim XDistance, YDistance As Double
        Dim NeckDepth, XOffset As Double
        Dim RightArcStart, RightArcEnd As Double
        Dim RightArcStep, HeadBack As Double
        Dim TempLen, StrapLength As Double
        Dim XLen2, Xlen1, VelcroLen As Double
        Dim VelcroText As Object

        If optChinStrap.Checked = True Or optModifiedChinStrap.Checked = True Then
            PR_ChinStrapDimensions()

            ' Get Neck Points
            XDistance = Val(txtCSNeckCircum.Text) / 2
            RightArcStart = 9 / 2.54
            RightArcEnd = 17.05 / 2.54
            RightArcStep = (RightArcEnd - RightArcStart) / 16
            RightArcStep = RightArcStep * (Val(txtRadiusNo.Text) - 1)
            HeadBack = RightArcStart + RightArcStep
            XOffset = 0

            PR_MakeXY(xyNeckTopBack, XDistance, 0)
            If xyNeckTopBack.X > HeadBack Then
                XOffset = -(xyNeckTopBack.X - HeadBack)
                PR_MakeXY(xyNeckTopBack, xyNeckTopBack.X + XOffset, 0)
            End If

            PR_MakeXY(xyNeckTopFront, -XDistance + XOffset, 0)
            txtxyNeckTopFrontX.Text = CStr(xyNeckTopFront.X)
            txtxyNeckTopFrontY.Text = CStr(xyNeckTopFront.Y)

            ' Set Min  Neck depth of 2" for adults and 1.5" Under ten
            If txtThroatToSternal.Text = "" Then
                If Val(txtAge.Text) >= 10 Then
                    NeckDepth = 2
                Else
                    NeckDepth = 1.5
                End If
            ElseIf txtThroatToSternal.Text <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                NeckDepth = FN_CmToInches(txtUnits, txtThroatToSternal)
            End If

            If NeckDepth < 1 Then
                MsgBox("Neck Depth should be Greater than 1'', Consult with Supervisor.", 48, "Chin Strap Details")
            End If

            PR_MakeXY(xyNeckBottomFront, xyNeckTopFront.X, -NeckDepth)
            PR_MakeXY(xyNeckBottomBack, xyNeckTopBack.X, -NeckDepth)
            PR_MakeXY(xyNeckTopVelcro, xyNeckTopBack.X + 1, xyNeckTopBack.Y)
            PR_MakeXY(xyNeckBottomVelcro, xyNeckBottomBack.X + 1 + XOffset, xyNeckBottomBack.Y)

            PR_DrawLine(xyNeckBottomVelcro, xyNeckTopVelcro)
            PR_AddEntityID("VelcroBack")
            PR_SetLayer("Notes")
            PR_DrawLine(xyNeckTopBack, xyNeckBottomBack)
            PR_SetLayer("TemplateLeft")
            PR_AddEntityID("NeckBack")

            PR_SetLayer("Notes")
            PR_MakeXY(xyStrapText, xyNeckTopBack.X, xyNeckTopBack.Y - 0.25)

            VelcroLen = NeckDepth - 0.125
            'UPGRADE_WARNING: Couldn't resolve default property of object VelcroText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                VelcroText = Str(VelcroLen)
            Else
                VelcroText = FN_InchesToText(VelcroLen) + Chr(34)
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object VelcroText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'VelcroText = VelcroText & "\"""
            PR_DrawText(VelcroText, xyStrapText, 0.1, 0)

            PR_MakeXY(xyStrapText, xyNeckTopBack.X + 0.1, xyNeckTopBack.Y - 0.5)
            PR_DrawText(" Velcro", xyStrapText, 0.1, 0)
            PR_SetLayer("TemplateLeft")

            ' Set Top Front Point
            YDistance = Val(txtCSChinAngle.Text)
            PR_MakeXY(xyChinStrapTopFront, 0, YDistance)

            ' Set FaceLine and Back Line
            XDistance = Val(txtCSChinAngle.Text) / 2
            PR_MakeXY(xyFacelineTop, xyChinStrapTopFront.X - XDistance, xyChinStrapTopFront.Y)
            PR_MakeXY(xyBacklineTop, xyChinStrapTopFront.X + XDistance, xyChinStrapTopFront.Y)

            PR_MakeXY(xyFacelineBot, xyFacelineTop.X, 0)
            PR_MakeXY(xyBacklineBot, xyBacklineTop.X, 0)

            ' Set Mouth and lipstrap
            YDistance = Val(txtCSChintoMouth.Text)
            XDistance = Val(txtCSMouthWidth.Text)
            PR_MakeXY(xyMouthFront, xyFacelineBot.X, xyFacelineBot.Y + YDistance)
            PR_MakeXY(xyMouthBack, xyMouthFront.X + XDistance, xyMouthFront.Y)

            PR_MakeXY(xyLipStrapTop, xyMouthFront.X - 0.125, xyMouthFront.Y + 0.75)
            PR_MakeXY(xyLipStrapBottom, xyMouthFront.X - 0.125, xyMouthFront.Y)

            ' Set and draw front curve
            PR_DrawCSFrontCurve(xyLipStrapTop, xyLipStrapBottom, xyMouthFront, xyMouthBack, xyFacelineTop, xyChinStrapTopFront)

            ' Set and draw back curve
            If optChinStrap.Checked = True Then
                PR_DrawCSBackCurve(xyBacklineBot, xyFacelineTop, xyChinStrapTopFront, xyNeckTopVelcro)
            ElseIf optModifiedChinStrap.Checked = True Then
                PR_DrawCSModBackCurve(xyNeckTopVelcro, xyChinStrapTopFront)
            End If

            ' Draw Top
            If optChinStrap.Checked = True Then
                PR_MakeXY(xyChinStrapTopBack, xyChinStrapTopFront.X + 2, xyChinStrapTopFront.Y)
            ElseIf optModifiedChinStrap.Checked = True Then
                PR_MakeXY(xyChinStrapTopBack, xyChinStrapTopFront.X + 2.5, xyChinStrapTopFront.Y)
            End If
            PR_DrawLine(xyChinStrapTopFront, xyChinStrapTopBack)
            PR_AddEntityID("ChinStrapTop")

            ' Draw Velcro Top
            PR_DrawLine(xyNeckTopBack, xyNeckTopVelcro)
            PR_AddEntityID("VelcroTop")

            ' Set and draw chin
            PR_DrawCSChin(xyMouthFront, xyMouthBack, xyNeckBottomFront, xyNeckBottomVelcro)

            ' Calculate Strap Length
            TempLen = fnRoundInches(FN_CalcLength(xyChinStrapTopBack, xyNeckTopBack))
            StrapLength = Int(TempLen)
            TempLen = TempLen - StrapLength
            StrapLength = StrapLength - 1 'Subtract an inch
            'Round down to nearest 1/2"
            If TempLen >= 0.5 Then
                StrapLength = StrapLength + 0.5
            End If

            'PR_DrawMarker(xyChinStrapTopBack)
            'PR_DrawText "xyChinStrapTopBack", xyChinStrapTopBack, .05, 0
            'PR_DrawMarker xyNeckTopBack
            'PR_DrawText "xyNeckTopBack", xyNeckTopBack, .05, 0

            'To Draw Cross Block
            PR_DrawCrossBlock()

            PR_MakeXY(xyStrapText, xyChinStrapTopFront.X + 0.1, xyChinStrapTopBack.Y - 0.5)
            PR_SetLayer("Notes")
            'PR_DrawText("Strap Length: " & FN_InchesToText(StrapLength) & "\""", xyStrapText, 0.1, 0)
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                ''Changed for #161 in issue list
                StrapLength = (StrapLength * 2.54) - 2.5
                Dim nInteger As Integer = Int(StrapLength)
                Dim Dec As Double = (StrapLength - nInteger) * 10
                Dec = Int(Dec) / 10
                StrapLength = nInteger + Dec
                PR_DrawText("Strap Length: " & Str(StrapLength) & " CM", xyStrapText, 0.1, 0)
            Else
                PR_DrawText("Strap Length: " & FN_InchesToText(StrapLength) & Chr(34), xyStrapText, 0.1, 0)
            End If
            PR_SetLayer("TemplateLeft")

            ' Set and draw Lip Covering
            PR_DrawCSLipCovering()

            ' Set and Draw Ear Openings
            PR_DrawEars()
            PR_DrawEarFlaps()


            ' Add Patient Details
            PR_MakeXY(xyTextPoint, -1.5, 0.75)
            PR_TemplateDetails(xyTextPoint, "Patient Details")

        End If

    End Sub

    Private Sub PR_DrawCSBackCurve(ByRef xyBacklineBot As HEADNECK1.xy, ByRef xyFacelineTop As HEADNECK1.xy, ByRef xyChinStrapTopFront As HEADNECK1.xy, ByRef xyNeckTopVelcro As HEADNECK1.xy)

        Dim XDistance, YDistance As Double
        Dim nRadius, aAngle, XStep, YStep, aInc, nOffSet As Double
        Dim ii As Short
        Dim xyArcCen, xyTopEar, xyBotEar, xyTmp As HEADNECK1.xy

        XDistance = System.Math.Abs(xyBacklineBot.X) - 1
        YDistance = xyFacelineTop.Y - CDbl(txtCSEarTopHeight.Text)

        XStep = XDistance / 20
        YStep = YDistance / 5

        PR_MakeXY(xyTopEar, 0, CDbl(txtCSEarTopHeight.Text))
        PR_MakeXY(xyBotEar, 0, CDbl(txtCSEarBotHeight.Text))
        nOffSet = (CDbl(txtCSEarTopHeight.Text) - CDbl(txtCSEarBotHeight.Text)) / 8

        ChinStrapBackX(1) = xyChinStrapTopFront.X + 2
        ChinStrapBackY(1) = xyChinStrapTopFront.Y
        ChinStrapBackX(2) = (xyChinStrapTopFront.X + 2) - (1 * XStep)
        ChinStrapBackY(2) = xyChinStrapTopFront.Y - YStep
        ChinStrapBackX(3) = (xyChinStrapTopFront.X + 2) - (2.5 * XStep)
        ChinStrapBackY(3) = xyChinStrapTopFront.Y - (2 * YStep)
        ChinStrapBackX(4) = (xyChinStrapTopFront.X + 2) - (4.5 * XStep)
        ChinStrapBackY(4) = xyChinStrapTopFront.Y - (3 * YStep)
        ChinStrapBackX(5) = (xyChinStrapTopFront.X + 2) - (7.5 * XStep)
        ChinStrapBackY(5) = xyChinStrapTopFront.Y - (4 * YStep)

        PR_MakeXY(xyTmp, ChinStrapBackX(5), ChinStrapBackY(5))
        aAngle = FN_CalcAngle(xyTopEar, xyTmp) * (180 / HEADNECK1.PI)
        PR_CalcPolar(xyTopEar, aAngle, nOffSet, xyTmp)

        aAngle = (aAngle + 90) / 2
        nRadius = nOffSet * System.Math.Tan(aAngle * (HEADNECK1.PI / 180))
        PR_MakeXY(xyArcCen, xyTopEar.X + nRadius, xyTopEar.Y - nOffSet)
        aAngle = FN_CalcAngle(xyArcCen, xyTmp) * (180 / HEADNECK1.PI)
        aInc = (180 - aAngle) / 3
        For ii = 6 To 9
            PR_CalcPolar(xyArcCen, aAngle, nRadius, xyTmp)
            aAngle = aAngle + aInc
            ChinStrapBackX(ii) = xyTmp.X
            ChinStrapBackY(ii) = xyTmp.Y
        Next ii

        'Original code GG-19.Sep.96
        '    ChinStrapBackX(6) = 0
        '    ChinStrapBackY(6) = txtCSEarTopHeight
        '    ChinStrapBackX(7) = 0
        '    ChinStrapBackY(7) = txtCSEarBotHeight + (5 * ((txtCSEarTopHeight - txtCSEarBotHeight) / 6))
        '    ChinStrapBackX(8) = 0
        '    ChinStrapBackY(8) = txtCSEarBotHeight + (2 * ((txtCSEarTopHeight - txtCSEarBotHeight) / 3))
        '    ChinStrapBackX(9) = 0
        '    ChinStrapBackY(9) = txtCSEarBotHeight + ((txtCSEarTopHeight - txtCSEarBotHeight) / 2)
        '    ChinStrapBackX(10) = 0
        '    ChinStrapBackY(10) = txtCSEarBotHeight + ((txtCSEarTopHeight - txtCSEarBotHeight) / 3)
        '    ChinStrapBackX(11) = 0
        '    ChinStrapBackY(11) = txtCSEarBotHeight + (1 * ((txtCSEarTopHeight - txtCSEarBotHeight) / 6))
        '    ChinStrapBackX(12) = 0
        '    ChinStrapBackY(12) = txtCSEarBotHeight
        '

        ChinStrapBackX(13) = 1.5 * XStep
        ChinStrapBackY(13) = CDbl(txtCSEarBotHeight.Text) / 2

        PR_MakeXY(xyTmp, ChinStrapBackX(13), ChinStrapBackY(13))
        aAngle = FN_CalcAngle(xyBotEar, xyTmp) * (180 / HEADNECK1.PI)
        PR_CalcPolar(xyBotEar, aAngle, nOffSet, xyTmp)

        aAngle = ((360 - aAngle) + 90) / 2
        nRadius = nOffSet * System.Math.Tan(aAngle * (HEADNECK1.PI / 180))
        PR_MakeXY(xyArcCen, xyBotEar.X + nRadius, xyBotEar.Y + nOffSet)
        aAngle = FN_CalcAngle(xyArcCen, xyTmp) * (180 / HEADNECK1.PI)
        aInc = (aAngle - 180) / 3
        aAngle = 180
        For ii = 10 To 12
            PR_CalcPolar(xyArcCen, aAngle, nRadius, xyTmp)
            aAngle = aAngle + aInc
            ChinStrapBackX(ii) = xyTmp.X
            ChinStrapBackY(ii) = xyTmp.Y
        Next ii

        ChinStrapBackX(14) = 4 * XStep
        ChinStrapBackY(14) = 0.125
        ChinStrapBackX(15) = 10 * XStep
        ChinStrapBackY(15) = 0
        ChinStrapBackX(16) = 15 * XStep
        ChinStrapBackY(16) = 0
        ChinStrapBackX(17) = xyNeckTopVelcro.X - 1
        ChinStrapBackY(17) = 0

        PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "fitted" & HEADNECK1.QQ)
        For ii = 1 To 17
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(ChinStrapBackX(ii)) & HEADNECK1.CC & "xyStart.y+" & Str(ChinStrapBackY(ii)))
        Next ii
        PrintLine(CInt(txtfNum.Text), ");")
        PR_AddEntityID("ChinStrapBack")
        ''To Draw Spline
        Dim ptColl As Point3dCollection = New Point3dCollection()
        For ii = 1 To 17
            ptColl.Add(New Point3d(xyInsertion.X + ChinStrapBackX(ii), xyInsertion.Y + ChinStrapBackY(ii), 0))
        Next ii
        PR_DrawSpline(ptColl)
        'For ii = 1 To 17
        '    PR_MakeXY xyTmp, ChinStrapBackX(ii), ChinStrapBackY(ii)
        '    PR_DrawMarker xyTmp
        '    PR_DrawText Trim$(Str$(ii)), xyTmp, .05, 0
        'Next ii

    End Sub

    Private Sub PR_DrawCSChin(ByRef xyMouthFront As HEADNECK1.xy, ByRef xyMouthBack As HEADNECK1.xy, ByRef xyNeckBottomFront As HEADNECK1.xy, ByRef xyNeckBottomVelcro As HEADNECK1.xy)

        Dim xyChinTop, xyChinBottom As HEADNECK1.xy
        Dim xyBotCurve, xyTopCurve As HEADNECK1.xy
        Dim xyDiagEnd, xyDiagStart, xyTmp As HEADNECK1.xy
        Dim ChinStep2, ChinStep1, ChinDist As Double
        Dim xyChinArc As HEADNECK1.xy
        Dim PI, ChinHt As Double
        Dim BotRadius, TopRadius As Double
        Dim Theta, Delta As Double
        Dim XA, YA As Double
        Dim X As Short
        Dim DiagLen, DiagStep As Double
        Dim ii, j, i, v As Short
        Dim OffsetChinX(12) As Double
        Dim OffsetChinY(12) As Double

        TopRadius = 1.8 / 2.54
        BotRadius = 0.9 / 2.54
        PI = 3.141592654

        If optChinStrap.Checked = True Or optModifiedChinStrap.Checked = True Then

            'Set and Draw Chin
            PR_MakeXY(xyChinTop, xyMouthFront.X + 0.25, xyMouthFront.Y)
            PR_MakeXY(xyChinBottom, xyNeckBottomFront.X + 0.25, xyNeckBottomFront.Y)
            txtxyChinTopX.Text = CStr(xyChinTop.X)
            txtxyChinTopY.Text = CStr(xyChinTop.Y)

            xyChinTop.X = xyChinTop.X + 0.125
            xyChinBottom.X = xyChinBottom.X + 0.125

            ChinHt = System.Math.Abs(xyChinTop.Y - xyChinBottom.Y)
            Delta = 30

            If Val(txtThroatToSternal.Text) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ChinDist = FN_CmToInches(txtUnits, txtThroatToSternal)
            Else
                If Val(txtAge.Text) > 10 Then
                    ChinDist = 2
                Else
                    ChinDist = 1.5
                End If
            End If


            If ChinDist >= 2 Then
                ChinStep1 = 2
                ChinStep2 = 0.375
            Else
                ChinStep1 = ChinDist
                '        ChinStep2 = (3 * ChinDist) / 10
                ChinStep2 = 0.375
            End If

            ChinStrapChinX(1) = (xyChinBottom.X)
            ChinStrapChinY(1) = (xyChinBottom.Y)

            ChinStrapChinX(2) = xyChinBottom.X
            ChinStrapChinY(2) = 0 - ChinStep1

            ChinStrapChinX(3) = xyChinBottom.X
            ChinStrapChinY(3) = 0 - ChinStep2

            For i = 1 To 3
                OffsetChinX(i) = ChinStrapChinX(i) - 0.125
                OffsetChinY(i) = ChinStrapChinY(i)
            Next i

            PR_MakeXY(xyBotCurve, ChinStrapChinX(3) - BotRadius, ChinStrapChinY(3))

            For i = 1 To 3
                Theta = (i * (Delta)) * (PI / 180)
                XA = BotRadius * System.Math.Cos(Theta)
                YA = BotRadius * System.Math.Sin(Theta)
                ChinStrapChinX(i + 3) = xyBotCurve.X + XA
                ChinStrapChinY(i + 3) = xyBotCurve.Y + YA
                XA = (BotRadius - 0.125) * System.Math.Cos(Theta)
                YA = (BotRadius - 0.125) * System.Math.Sin(Theta)
                OffsetChinX(i + 3) = xyBotCurve.X + XA
                OffsetChinY(i + 3) = xyBotCurve.Y + YA
            Next i

            ChinStrapChinX(12) = xyChinTop.X
            ChinStrapChinY(12) = xyChinTop.Y

            ChinStrapChinX(10) = xyChinTop.X
            ChinStrapChinY(10) = ChinStrapChinY(7) + TopRadius

            ChinStrapChinY(10) = xyBotCurve.Y + BotRadius + TopRadius

            ChinStrapChinX(11) = xyChinTop.X
            ChinStrapChinY(11) = xyChinTop.Y - ((ChinStrapChinY(12) - ChinStrapChinY(10)) / 2)

            For i = 10 To 12
                OffsetChinX(i) = ChinStrapChinX(i) - 0.125
                OffsetChinY(i) = ChinStrapChinY(i)
            Next i

            PR_MakeXY(xyTopCurve, ChinStrapChinX(10) + TopRadius, ChinStrapChinY(10))

            j = 9
            For i = 1 To 3
                Theta = (i * Delta) * (PI / 180)
                XA = TopRadius * System.Math.Cos(Theta)
                YA = TopRadius * System.Math.Sin(Theta)
                ChinStrapChinX(j) = xyTopCurve.X - XA
                ChinStrapChinY(j) = xyTopCurve.Y - YA
                XA = (TopRadius + 0.125) * System.Math.Cos(Theta)
                YA = (TopRadius + 0.125) * System.Math.Sin(Theta)
                OffsetChinX(j) = xyTopCurve.X - XA
                OffsetChinY(j) = xyTopCurve.Y - YA
                j = j - 1
            Next i

            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "fitted" & HEADNECK1.QQ)
            For ii = 1 To 12
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(OffsetChinX(ii)) & HEADNECK1.CC & "xyStart.y+" & Str(OffsetChinY(ii)))
            Next ii
            PrintLine(CInt(txtfNum.Text), ");")
            ''To Draw Spline
            Dim ptColl As Point3dCollection = New Point3dCollection()
            For ii = 1 To 12
                ptColl.Add(New Point3d(xyInsertion.X + OffsetChinX(ii), xyInsertion.Y + OffsetChinY(ii), 0))
            Next ii
            PR_DrawSpline(ptColl)

            'Revise back
            xyChinTop.X = xyChinTop.X - 0.125
            xyChinBottom.X = xyChinBottom.X - 0.125

            PR_AddEntityID("Chin")
            PR_DrawLine(xyChinTop, xyMouthBack)
            PR_AddEntityID("ChinStrapMouth")
            PR_DrawLine(xyChinBottom, xyNeckBottomVelcro)
            PR_AddEntityID("ChinStrapBottom")

            If chkNeckElastic.CheckState = 1 Then
                PR_MakeXY(xyTmp, xyChinBottom.X + 0.4, xyChinBottom.Y + 0.5)
                PR_SetLayer("Notes")
                PR_DrawText("1" & Chr(34) & " Elastic Sewn Underneath", xyTmp, 0.1, 0)
                PR_SetLayer("TemplateLeft")
            End If


            PR_SetLayer("Construct")
            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "fitted" & HEADNECK1.QQ)
            For ii = 1 To 12
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(ChinStrapChinX(ii)) & HEADNECK1.CC & "xyStart.y+" & Str(ChinStrapChinY(ii)))
            Next ii
            PrintLine(CInt(txtfNum.Text), ");")
            PR_AddEntityID("LipStrapChinConstructLine")
            ''To Draw Spline
            Dim ptCollection As Point3dCollection = New Point3dCollection()
            For ii = 1 To 12
                ptCollection.Add(New Point3d(xyInsertion.X + ChinStrapChinX(ii), xyInsertion.Y + ChinStrapChinY(ii), 0))
            Next ii
            PR_DrawSpline(ptCollection)

            'Set Global Points for chin curve
            'UPGRADE_WARNING: Couldn't resolve default property of object xyChinTopCen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xyChinTopCen = xyTopCurve
            'UPGRADE_WARNING: Couldn't resolve default property of object xyChinBotCen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xyChinBotCen = xyBotCurve

        End If

    End Sub

    Private Sub PR_DrawCSFrontCurve(ByRef xyLipStrapTop As HEADNECK1.xy, ByRef xyLipStrapBottom As HEADNECK1.xy, ByRef xyMouthFront As HEADNECK1.xy, ByRef xyMouthBack As HEADNECK1.xy, ByRef xyFacelineTop As HEADNECK1.xy, ByRef xyChinStrapTopFront As HEADNECK1.xy)

        Dim FrontStartX, FrontStartY As Double
        Dim XDistance, YDistance As Double
        Dim XStep, YStep As Double
        Dim ii As Short

        ' Set Front Curve

        If chkLipStrap.CheckState = 1 Then
            FrontStartX = xyLipStrapTop.X
            FrontStartY = xyLipStrapTop.Y
            PR_DrawLine(xyLipStrapTop, xyLipStrapBottom)
            PR_AddEntityID("ChinStrapFront")
            PR_DrawLine(xyMouthBack, xyLipStrapBottom)
            PR_AddEntityID("Mouth")
            XDistance = System.Math.Abs(xyLipStrapTop.X)
            YDistance = xyFacelineTop.Y - xyLipStrapTop.Y

        Else
            XDistance = System.Math.Abs(xyMouthBack.X)
            YDistance = xyFacelineTop.Y - xyMouthFront.Y
            FrontStartX = xyMouthBack.X
            FrontStartY = xyMouthBack.Y
        End If


        XStep = XDistance / 7
        YStep = YDistance / 20


        ChinStrapFrontX(1) = FrontStartX
        ChinStrapFrontY(1) = FrontStartY
        ChinStrapFrontX(2) = FrontStartX + (XStep)
        ChinStrapFrontY(2) = FrontStartY + (YStep * YStep)
        ChinStrapFrontX(3) = FrontStartX + (XStep * 2)
        ChinStrapFrontY(3) = FrontStartY + ((2 * YStep) * (2 * YStep))
        ChinStrapFrontX(4) = FrontStartX + (XStep * 3)
        ChinStrapFrontY(4) = FrontStartY + ((3 * YStep) * (3 * YStep))
        ChinStrapFrontX(5) = FrontStartX + (XStep * 4)
        ChinStrapFrontY(5) = FrontStartY + ((4 * YStep) * (4 * YStep))
        ChinStrapFrontX(6) = FrontStartX + (XStep * 5)
        ChinStrapFrontY(6) = FrontStartY + ((5 * YStep) * (5 * YStep))
        ChinStrapFrontX(7) = FrontStartX + (XStep * 6)
        ChinStrapFrontY(7) = FrontStartY + ((6 * YStep) * (6 * YStep))
        ChinStrapFrontX(8) = xyChinStrapTopFront.X
        ChinStrapFrontY(8) = xyChinStrapTopFront.Y

        PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "fitted" & HEADNECK1.QQ)
        For ii = 1 To 8
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(ChinStrapFrontX(ii)) & HEADNECK1.CC & "xyStart.y+" & Str(ChinStrapFrontY(ii)))
        Next ii
        PrintLine(CInt(txtfNum.Text), ");")
        PR_AddEntityID("ChinStrapFront")
        ''To Draw Spline
        Dim ptColl As Point3dCollection = New Point3dCollection()
        For ii = 1 To 8
            ptColl.Add(New Point3d(xyInsertion.X + ChinStrapFrontX(ii), xyInsertion.Y + ChinStrapFrontY(ii), 0))
        Next ii
        PR_DrawSpline(ptColl)
    End Sub

    Private Sub PR_DrawCSLipCovering()

        'Uses module level variable xyChinTopCen & xyChinBotCen
        '
        Dim xyLipCoverTopLeft, xyLipCoverTopRight As HEADNECK1.xy
        Dim xyLipCoverBottomRight, xyLipCoverBottomMid As HEADNECK1.xy
        Dim xy2, xy6, xyTmp, xy7, xy3, xy1 As HEADNECK1.xy

        Dim nn As Short
        Dim aAngle, nOffSet, BotRadius, aStart As Double
        Dim XPoint, YPoint As Double
        Dim ChinVal As Short
        Dim j, i, ii, DrawArcPart As Short
        Dim DistStepX, DistStepY As Double
        Dim ChinCheckY, XOffset As Double
        Dim xyLipCoveringBotCen As HEADNECK1.xy
        Dim LipDistX, LipDistY As Double


        'If chkLipCovering.Value = 1 And chkLipStrap.Value = 1 Then
        Dim Point1, Point2 As HEADNECK1.xy
        Dim length As Double
        Dim VelcroText As Object
        Dim TextPoint As HEADNECK1.xy
        If chkLipCovering.CheckState = 1 Then

            XOffset = -System.Math.Abs(CDbl(txtxyChinTopX.Text))
            BotRadius = 0.9 / 2.54

            ' Set Top Left & Right Points
            XPoint = CDbl(txtxyChinTopX.Text)
            YPoint = CDbl(txtxyChinTopY.Text)
            'FUDGE Start - GG 18.Sep.95
            '    PR_MakeXY xyLipCoverTopLeft, XPoint + .125, YPoint + 1.25
            '    PR_MakeXY xyLipCoverTopRight, -.875, YPoint + 1.25
            '    PR_MakeXY xyLipCoverBottomRight, -.875, .25
            PR_MakeXY(xyLipCoverTopLeft, XPoint + 0.25, YPoint + 1.25)
            PR_MakeXY(xyLipCoverTopRight, -0.75, YPoint + 1.25)
            PR_MakeXY(xyLipCoverBottomRight, -0.75, 0.25)
            'FUDGE End - GG 18.Sep.95
            XPoint = CDbl(txtxyNeckTopFrontX.Text)
            YPoint = CDbl(txtxyNeckTopFrontY.Text)

            'PR_DrawMarker xyLipCoverTopLeft
            'PR_DrawText "xyLipCoverTopLeft", xyLipCoverTopLeft, .05, 0
            'PR_DrawMarker xyLipCoverTopRight
            'PR_DrawText "xyLipCoverTopRight", xyLipCoverTopRight, .05, 0
            'PR_DrawMarker xyLipCoverBottomRight
            'PR_DrawText "xyLipCoverBottomRight", xyLipCoverBottomRight, .05, 0

            'Find Bottom Mid Point
            'This is the intesection with the neck line (y=0)
            'at xy6
            PR_MakeXY(xy6, ChinStrapChinX(6), ChinStrapChinY(6))
            PR_MakeXY(xy7, ChinStrapChinX(7), ChinStrapChinY(7))

            'Take intersection point at end of horizontal line
            DrawArcPart = True
            aStart = FN_CalcAngle(xyChinBotCen, xy6) * (180 / HEADNECK1.PI)
            aAngle = aStart - ((0.375 / (2 * HEADNECK1.PI * BotRadius)) * 360)
            PR_CalcPolar(xyChinBotCen, aAngle, BotRadius, xyLipCoverBottomMid)

            xyLipCoverBottomMid.X = xyLipCoverBottomMid.X + 0.125

            ' Get Points For Curve, From Mid to Bottom Right
            ' Adjust to include Lip
            ChinStrapLipX(1) = xyLipCoverTopLeft.X
            ChinStrapLipY(1) = xyLipCoverTopLeft.Y

            ChinStrapLipX(2) = ChinStrapChinX(12) + 0.125
            ChinStrapLipY(2) = ChinStrapChinY(12)
            ChinStrapLipX(3) = ChinStrapChinX(11) + 0.125
            ChinStrapLipY(3) = ChinStrapChinY(11)
            nn = 4
            For ii = 10 To 7 Step -1
                Select Case nn
                    Case 4
                        nOffSet = 0.08
                    Case 5
                        nOffSet = 0.125
                    Case 6
                        nOffSet = 0.125
                    Case 7
                        nOffSet = 0.06
                End Select
                PR_MakeXY(xyTmp, ChinStrapChinX(ii) + 0.125, ChinStrapChinY(ii))
                PR_CalcPolar(xyChinTopCen, (FN_CalcAngle(xyChinTopCen, xyTmp) * (180 / HEADNECK1.PI)), FN_CalcLength(xyChinTopCen, xyTmp) - nOffSet, xyTmp)
                ChinStrapLipX(nn) = xyTmp.X
                ChinStrapLipY(nn) = xyTmp.Y
                nn = nn + 1
            Next ii


            If DrawArcPart Then
                ChinStrapLipX(8) = ChinStrapChinX(6) + 0.125
                ChinStrapLipY(8) = ChinStrapChinY(6)

                PR_CalcPolar(xyChinBotCen, aAngle + ((aStart - aAngle) / 3), BotRadius, xyTmp)
                ChinStrapLipX(10) = xyTmp.X + 0.125
                ChinStrapLipY(10) = xyTmp.Y

                PR_CalcPolar(xyChinBotCen, aAngle + (2 * ((aStart - aAngle) / 3)), BotRadius, xyTmp)
                ChinStrapLipX(9) = xyTmp.X + 0.125
                ChinStrapLipY(9) = xyTmp.Y

                ChinStrapLipX(11) = xyLipCoverBottomMid.X
                ChinStrapLipY(11) = xyLipCoverBottomMid.Y
            Else
                ChinStrapLipX(8) = xyLipCoverBottomMid.X
                ChinStrapLipY(8) = xyLipCoverBottomMid.Y
            End If

            'Add Offset to curve
            For i = 1 To 11
                ChinStrapLipX(i) = ChinStrapLipX(i) + XOffset
            Next i

            'Indicate position on Pattern
            'Somewhere above we have an extra .125" add in the X axis
            'I don't know why etc. etc. etc.
            PR_SetLayer("Notes")

            PR_MakeXY(Point1, xyLipCoverTopRight.X - 0.125, xyLipCoverTopRight.Y)
            PR_MakeXY(Point2, Point1.X - 0.125, Point1.Y)
            PR_DrawLine(Point1, Point2)
            PR_MakeXY(Point2, Point1.X, Point1.Y - 0.125)
            PR_DrawLine(Point1, Point2)

            PR_MakeXY(Point1, xyLipCoverBottomRight.X - 0.125, xyLipCoverBottomRight.Y)
            PR_MakeXY(Point2, Point1.X - 0.125, Point1.Y)
            PR_DrawLine(Point1, Point2)
            PR_MakeXY(Point2, Point1.X, Point1.Y + 0.125)
            PR_DrawLine(Point1, Point2)

            PR_MakeXY(Point1, xyLipCoverBottomMid.X - 0.125, xyLipCoverBottomMid.Y)
            PR_DrawLine(Point1, xyLipCoverBottomMid)

            PR_SetLayer("TemplateLeft")

            PR_MakeXY(xyLipCoverTopLeft, xyLipCoverTopLeft.X + XOffset, xyLipCoverTopLeft.Y)
            PR_MakeXY(xyLipCoverTopRight, xyLipCoverTopRight.X + XOffset, xyLipCoverTopRight.Y)
            PR_MakeXY(xyLipCoverBottomRight, xyLipCoverBottomRight.X + XOffset, xyLipCoverBottomRight.Y)
            PR_MakeXY(xyLipCoverBottomMid, xyLipCoverBottomMid.X + XOffset, xyLipCoverBottomMid.Y)

            nOffSet = FN_CalcLength(xyLipCoverBottomMid, xyLipCoverBottomRight) / 2
            aAngle = FN_CalcAngle(xyLipCoverBottomMid, xyLipCoverBottomRight) * (180 / HEADNECK1.PI)
            PR_CalcPolar(xyLipCoverBottomMid, aAngle, nOffSet, xyTmp)
            PR_CalcPolar(xyTmp, aAngle + 270, 9, xyLipCoveringBotCen)
            PR_DrawArc(xyLipCoveringBotCen, xyLipCoverBottomRight, xyLipCoverBottomMid)
            PR_AddEntityID("LipCoveringBot")

            'PR_DrawMarker xyLipCoverBottomMid
            'PR_DrawText "xyLipCoverBottomMid", xyLipCoverBottomMid, .05, 0


            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "polyline" & HEADNECK1.QQ)
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyLipCoverTopLeft.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyLipCoverTopLeft.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyLipCoverTopRight.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyLipCoverTopRight.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyLipCoverBottomRight.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyLipCoverBottomRight.Y))
            PrintLine(CInt(txtfNum.Text), ");")
            PR_AddEntityID("LipCovering")
            ''To Draw Polyline
            Dim StrapX(3), StrapY(3) As Double
            StrapX(1) = xyLipCoverTopLeft.X
            StrapY(1) = xyLipCoverTopLeft.Y
            StrapX(2) = xyLipCoverTopRight.X
            StrapY(2) = xyLipCoverTopRight.Y
            StrapX(3) = xyLipCoverBottomRight.X
            StrapY(3) = xyLipCoverBottomRight.Y
            PR_DrawPoly(StrapX, StrapY, 3)

            'Draw chin area
            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "fitted" & HEADNECK1.QQ)
            If DrawArcPart Then
                For i = 1 To 11
                    PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(ChinStrapLipX(i)) & HEADNECK1.CC & "xyStart.y+" & Str(ChinStrapLipY(i)))
                Next i
            Else
                For i = 1 To 8
                    PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(ChinStrapLipX(i)) & HEADNECK1.CC & "xyStart.y+" & Str(ChinStrapLipY(i)))
                Next i
            End If
            PrintLine(CInt(txtfNum.Text), ");")
            PR_AddEntityID("ChinStrapLipCoveringCurve")
            ''To Draw Spline
            Dim ptCollection As Point3dCollection = New Point3dCollection()
            Dim nCount As Integer = 8
            If DrawArcPart Then
                nCount = 11
            End If
            For i = 1 To nCount
                ptCollection.Add(New Point3d(xyInsertion.X + ChinStrapLipX(i), xyInsertion.Y + ChinStrapLipY(i), 0))
            Next i
            PR_DrawSpline(ptCollection)


            PR_MakeXY(TextPoint, (xyLipCoverTopLeft.X + 0.25), (xyLipCoverTopLeft.Y - 0.1))
            PR_TemplateDetails(TextPoint, "Chin Strap Lip Covering")

            ' Add Velcro
            PR_SetLayer("Notes")
            PR_MakeXY(xy1, xyLipCoverTopRight.X - 0.75, xyLipCoverTopRight.Y)

            If xy1.X > xyLipCoverBottomMid.X Then
                PR_MakeXY(xy2, xy1.X, xyLipCoveringBotCen.Y)
                ii = FN_CirLinInt(xy1, xy2, xyLipCoveringBotCen, FN_CalcLength(xyLipCoverBottomRight, xyLipCoveringBotCen), xy2)
            Else
                PR_MakeXY(xy2, xy1.X, xyChinBotCen.Y)
                PR_MakeXY(xyTmp, (xyChinBotCen.X + XOffset) + 0.125, xyChinBotCen.Y)
                ii = FN_CirLinInt(xy1, xy2, xyTmp, BotRadius, xy2)
                'PR_DrawMarker xyTmp
                'PR_DrawText "xyTmp", xyTmp, .05, 0
            End If

            PR_DrawLine(xy1, xy2)
            length = FN_CalcLength(xy1, xy2) + 0.25
            'UPGRADE_WARNING: Couldn't resolve default property of object VelcroText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VelcroText = FN_InchesToText(ARMEDDIA1.fnRoundInches(length))
            'UPGRADE_WARNING: Couldn't resolve default property of object VelcroText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VelcroText = VelcroText & Chr(34) & "  Velcro"
            PR_MakeXY(TextPoint, xy1.X - 0.25, xy1.Y - 0.25)
            PR_DrawText(VelcroText, TextPoint, 0.1, (-90 * HEADNECK1.PI / 180))
            length = FN_CalcLength(xyLipCoverTopRight, xyLipCoverBottomRight) + 0.25
            'UPGRADE_WARNING: Couldn't resolve default property of object VelcroText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VelcroText = FN_InchesToText(ARMEDDIA1.fnRoundInches(length))
            'UPGRADE_WARNING: Couldn't resolve default property of object VelcroText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VelcroText = VelcroText & Chr(34) & "  Velcro"
            PR_MakeXY(TextPoint, xyLipCoverTopRight.X - 0.25, xy1.Y - 0.25)
            PR_DrawText(VelcroText, TextPoint, 0.1, (-90 * HEADNECK1.PI / 180))


            'Draw an offset chin curve on
            PR_SetLayer("Construct")
            XOffset = XOffset + 0.125
            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "fitted" & HEADNECK1.QQ)
            For i = 3 To 12
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(ChinStrapChinX(i) + XOffset) & HEADNECK1.CC & "xyStart.y+" & Str(ChinStrapChinY(i)))
            Next i
            PrintLine(CInt(txtfNum.Text), ");")
            PR_AddEntityID("LipStrapChinConstructLine")
            ''To Draw Spline
            Dim ptColl As Point3dCollection = New Point3dCollection()
            For i = 3 To 12
                ptColl.Add(New Point3d(xyInsertion.X + ChinStrapChinX(i) + XOffset, xyInsertion.Y + ChinStrapChinY(i), 0))
            Next i
            PR_DrawSpline(ptColl)
            PR_SetLayer("TemplateLeft")

        End If

    End Sub

    Private Sub PR_DrawCSModBackCurve(ByRef xyNeckTopVelcro As HEADNECK1.xy, ByRef xyChinStrapTopFront As HEADNECK1.xy)

        Dim BackStartX, BackStartY As Double
        Dim XDistance, YDistance As Double
        Dim XStep, YStep As Double
        Dim ii As Short

        ' Set Front Curve
        BackStartX = xyNeckTopVelcro.X - 1
        BackStartY = xyNeckTopVelcro.Y
        XDistance = xyNeckTopVelcro.X - 2.5
        YDistance = xyChinStrapTopFront.Y

        XStep = XDistance / 15
        YStep = YDistance / 20

        ChinStrapBackX(1) = BackStartX
        ChinStrapBackY(1) = BackStartY
        ChinStrapBackX(2) = BackStartX - (XStep)
        ChinStrapBackY(2) = BackStartY + (YStep * YStep)
        ChinStrapBackX(3) = BackStartX - (XStep * 2)
        ChinStrapBackY(3) = BackStartY + ((2 * YStep) * (2 * YStep))
        ChinStrapBackX(4) = BackStartX - (XStep * 3)
        ChinStrapBackY(4) = BackStartY + ((3 * YStep) * (3 * YStep))
        ChinStrapBackX(5) = BackStartX - (XStep * 4)
        ChinStrapBackY(5) = BackStartY + ((4 * YStep) * (4 * YStep))
        ChinStrapBackX(6) = BackStartX - (XStep * 5)
        ChinStrapBackY(6) = BackStartY + ((5 * YStep) * (5 * YStep))
        ChinStrapBackX(7) = BackStartX - (XStep * 6)
        ChinStrapBackY(7) = BackStartY + ((6 * YStep) * (6 * YStep))
        ChinStrapBackX(8) = xyChinStrapTopFront.X + 2.5
        ChinStrapBackY(8) = xyChinStrapTopFront.Y

        PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "fitted" & HEADNECK1.QQ)
        For ii = 1 To 8
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(ChinStrapBackX(ii)) & HEADNECK1.CC & "xyStart.y+" & Str(ChinStrapBackY(ii)))
        Next ii
        PrintLine(CInt(txtfNum.Text), ");")
        PR_AddEntityID("ChinStrapBack")
        ''To Draw Spline
        Dim ptColl As Point3dCollection = New Point3dCollection()
        For ii = 1 To 8
            ptColl.Add(New Point3d(xyInsertion.X + ChinStrapBackX(ii), xyInsertion.Y + ChinStrapBackY(ii), 0))
        Next ii
        PR_DrawSpline(ptColl)
    End Sub

    Private Sub PR_DrawEarFlaps()

        ' This procedure Draws the ear Flap

        Dim xyFlapTopLeft, xyFlapTopRight As HEADNECK1.xy
        Dim xyFlapBottomLeft, xyFlapBottomRight As HEADNECK1.xy
        Dim FlapX, DartAng, FlapY As Double
        Dim Point2, Point1, Notes As HEADNECK1.xy
        Dim PI As Double
        Dim Y1, X1, X2, y2 As Double
        Dim RightXOffSet, LeftXOffSet, YOffSet As Double
        Dim Left3, Left1, Left2, Left4 As HEADNECK1.xy
        Dim Right3, Right1, Right2, Right4 As HEADNECK1.xy

        PI = 3.141592654
        If optModifiedChinStrap.Checked = True Or optFaceMask.Checked = True Or optOpenFaceMask.Checked = True Then

            ' Calculate Bottom points
            If optModifiedChinStrap.Checked = False Then
                X1 = Val(txtDartStartX.Text)
                Y1 = Val(txtDartStartY.Text)
                X2 = Val(txtDartEndX.Text)
                y2 = Val(txtDartEndY.Text)
                PR_MakeXY(Point1, X1, Y1)
                PR_MakeXY(Point2, X2, y2)
                DartAng = FN_CalcAngle(Point1, Point2)
                'DartAng = DartAng * (PI / 180)
                FlapX = System.Math.Abs(CDbl(txtDartStartX.Text)) + 2
                FlapY = FlapX * System.Math.Tan(DartAng)
                PR_MakeXY(xyFlapBottomRight, Point1.X + FlapX, Point1.Y + FlapY + 0.25)
                PR_MakeXY(xyFlapBottomLeft, -0.625, xyFlapBottomRight.Y)
            Else
                PR_MakeXY(xyFlapBottomRight, 2.1875, 0.25)
                PR_MakeXY(xyFlapBottomLeft, -0.375, 0.25)
            End If
            ' Set OffSets
            LeftXOffSet = -3
            RightXOffSet = 1.625
            If chkOpenHeadMask.CheckState = 0 Then
                YOffSet = TopArcOpp(CInt(txtRadiusNo.Text)) + System.Math.Abs(xyFlapBottomLeft.Y) + 1.25
            Else
                YOffSet = System.Math.Abs(xyFlapBottomLeft.Y) + 1.75
            End If
            If optModifiedChinStrap.Checked = True Then
                YOffSet = Val(txtCSChinAngle.Text) + System.Math.Abs(xyFlapBottomLeft.Y) + 0.25
            End If

            ' Draw Left Ear
            If chkLeftEarFlap.CheckState = 1 Then
                ' Set Top Points
                If optModifiedChinStrap.Checked = False Then
                    PR_MakeXY(xyFlapTopLeft, -0.625, CDbl(txtTopLeftEar.Text) + 0.125)
                    PR_MakeXY(xyFlapTopRight, 2, CDbl(txtTopLeftEar.Text) + 0.125)
                Else
                    PR_MakeXY(xyFlapTopLeft, -0.375, CDbl(txtTopLeftEar.Text) + 0.125)
                    PR_MakeXY(xyFlapTopRight, 2.1875, CDbl(txtTopLeftEar.Text) + 0.125)
                End If

                'Indicate position on Pattern
                PR_SetLayer("Notes")
                PR_MakeXY(Point1, xyFlapTopLeft.X + 0.125, xyFlapTopLeft.Y)
                PR_MakeXY(Point2, xyFlapTopLeft.X, xyFlapTopLeft.Y - 0.125)
                PR_DrawLine(Point1, xyFlapTopLeft)
                PR_DrawLine(Point2, xyFlapTopLeft)

                PR_MakeXY(Point1, xyFlapTopRight.X - 0.125, xyFlapTopRight.Y)
                PR_MakeXY(Point2, xyFlapTopRight.X, xyFlapTopRight.Y - 0.125)
                PR_DrawLine(Point1, xyFlapTopRight)
                PR_DrawLine(Point2, xyFlapTopRight)

                PR_MakeXY(Point1, xyFlapBottomRight.X - 0.125, xyFlapBottomRight.Y)
                PR_MakeXY(Point2, xyFlapTopRight.X, xyFlapBottomRight.Y + 0.125)
                PR_DrawLine(Point1, xyFlapBottomRight)
                PR_DrawLine(Point2, xyFlapBottomRight)

                PR_MakeXY(Point1, xyFlapBottomLeft.X + 0.125, xyFlapBottomLeft.Y)
                PR_MakeXY(Point2, xyFlapBottomLeft.X, xyFlapBottomLeft.Y + 0.125)
                PR_DrawLine(Point1, xyFlapBottomLeft)
                PR_DrawLine(Point2, xyFlapBottomLeft)

                PR_MakeXY(Right1, xyFlapTopLeft.X + RightXOffSet, xyFlapTopLeft.Y + YOffSet)
                PR_MakeXY(Right2, xyFlapTopRight.X + RightXOffSet, xyFlapTopRight.Y + YOffSet)
                PR_MakeXY(Right3, xyFlapBottomRight.X + RightXOffSet, xyFlapBottomRight.Y + YOffSet)
                PR_MakeXY(Right4, xyFlapBottomLeft.X + RightXOffSet, xyFlapBottomLeft.Y + YOffSet)

                PR_SetLayer("TemplateLeft")

                PR_MakeXY(Left1, xyFlapTopLeft.X + LeftXOffSet, xyFlapTopLeft.Y + YOffSet)
                PR_MakeXY(Left2, xyFlapTopRight.X + LeftXOffSet, xyFlapTopRight.Y + YOffSet)
                PR_MakeXY(Left3, xyFlapBottomRight.X + LeftXOffSet, xyFlapBottomRight.Y + YOffSet)
                PR_MakeXY(Left4, xyFlapBottomLeft.X + LeftXOffSet, xyFlapBottomLeft.Y + YOffSet)

                PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "polyline" & HEADNECK1.QQ)
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Left1.X) & HEADNECK1.CC & "xyStart.y+" & Str(Left1.Y))
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Left2.X) & HEADNECK1.CC & "xyStart.y+" & Str(Left2.Y))
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Left3.X) & HEADNECK1.CC & "xyStart.y+" & Str(Left3.Y))
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Left4.X) & HEADNECK1.CC & "xyStart.y+" & Str(Left4.Y))
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Left1.X) & HEADNECK1.CC & "xyStart.y+" & Str(Left1.Y))
                PrintLine(CInt(txtfNum.Text), ");")
                ''To Draw Polyline
                Dim PolyStart(5), PolyEnd(5) As Double
                PolyStart(1) = Left1.X
                PolyEnd(1) = Left1.Y
                PolyStart(2) = Left2.X
                PolyEnd(2) = Left2.Y
                PolyStart(3) = Left3.X
                PolyEnd(3) = Left3.Y
                PolyStart(4) = Left4.X
                PolyEnd(4) = Left4.Y
                PolyStart(5) = Left1.X
                PolyEnd(5) = Left1.Y
                PR_DrawPoly(PolyStart, PolyEnd, 5)
                PR_AddEntityID("LeftEarFlap")

                ' Add notes
                Left1.X = Left1.X + 0.1
                Left1.Y = Left1.Y - 0.1
                PR_TemplateDetails(Left1, "Left Ear Flap")
            End If

            ' Draw Right Ear
            If chkRightEarFlap.CheckState = 1 Then
                ' Set Top Points
                If optModifiedChinStrap.Checked = False Then
                    PR_MakeXY(xyFlapTopLeft, -0.625, CDbl(txtTopRightEar.Text) + 0.125)
                    PR_MakeXY(xyFlapTopRight, 2, CDbl(txtTopRightEar.Text) + 0.125)
                Else
                    PR_MakeXY(xyFlapTopLeft, -0.375, CDbl(txtTopRightEar.Text) + 0.125)
                    PR_MakeXY(xyFlapTopRight, 2.1875, CDbl(txtTopRightEar.Text) + 0.125)
                End If
                'Indicate position on Pattern
                PR_SetLayer("Notes")
                PR_MakeXY(Point1, xyFlapTopLeft.X + 0.125, xyFlapTopLeft.Y)
                PR_MakeXY(Point2, xyFlapTopLeft.X, xyFlapTopLeft.Y - 0.125)
                PR_DrawLine(Point1, xyFlapTopLeft)
                PR_DrawLine(Point2, xyFlapTopLeft)

                PR_MakeXY(Point1, xyFlapTopRight.X - 0.125, xyFlapTopRight.Y)
                PR_MakeXY(Point2, xyFlapTopRight.X, xyFlapTopRight.Y - 0.125)
                PR_DrawLine(Point1, xyFlapTopRight)
                PR_DrawLine(Point2, xyFlapTopRight)

                PR_MakeXY(Point1, xyFlapBottomRight.X - 0.125, xyFlapBottomRight.Y)
                PR_MakeXY(Point2, xyFlapTopRight.X, xyFlapBottomRight.Y + 0.125)
                PR_DrawLine(Point1, xyFlapBottomRight)
                PR_DrawLine(Point2, xyFlapBottomRight)

                PR_MakeXY(Point1, xyFlapBottomLeft.X + 0.125, xyFlapBottomLeft.Y)
                PR_MakeXY(Point2, xyFlapBottomLeft.X, xyFlapBottomLeft.Y + 0.125)
                PR_DrawLine(Point1, xyFlapBottomLeft)
                PR_DrawLine(Point2, xyFlapBottomLeft)

                PR_MakeXY(Right1, xyFlapTopLeft.X + RightXOffSet, xyFlapTopLeft.Y + YOffSet)
                PR_MakeXY(Right2, xyFlapTopRight.X + RightXOffSet, xyFlapTopRight.Y + YOffSet)
                PR_MakeXY(Right3, xyFlapBottomRight.X + RightXOffSet, xyFlapBottomRight.Y + YOffSet)
                PR_MakeXY(Right4, xyFlapBottomLeft.X + RightXOffSet, xyFlapBottomLeft.Y + YOffSet)

                PR_SetLayer("TemplateRight")

                PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "polyline" & HEADNECK1.QQ)
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Right1.X) & HEADNECK1.CC & "xyStart.y+" & Str(Right1.Y))
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Right2.X) & HEADNECK1.CC & "xyStart.y+" & Str(Right2.Y))
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Right3.X) & HEADNECK1.CC & "xyStart.y+" & Str(Right3.Y))
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Right4.X) & HEADNECK1.CC & "xyStart.y+" & Str(Right4.Y))
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Right1.X) & HEADNECK1.CC & "xyStart.y+" & Str(Right1.Y))
                PrintLine(CInt(txtfNum.Text), ");")
                ''To Draw Polyline
                Dim PolyStart(5), PolyEnd(5) As Double
                PolyStart(1) = Right1.X
                PolyEnd(1) = Right1.Y
                PolyStart(2) = Right2.X
                PolyEnd(2) = Right2.Y
                PolyStart(3) = Right3.X
                PolyEnd(3) = Right3.Y
                PolyStart(4) = Right4.X
                PolyEnd(4) = Right4.Y
                PolyStart(5) = Right1.X
                PolyEnd(5) = Right1.Y
                PR_DrawPoly(PolyStart, PolyEnd, 5)
                PR_AddEntityID("RightEarFlap")

                ' Add notes
                Right1.X = Right1.X + 0.1
                Right1.Y = Right1.Y - 0.1
                PR_TemplateDetails(Right1, "Right Ear Flap")
            End If
        End If
    End Sub

    Private Sub PR_DrawEarHole(ByRef EarX As Double, ByRef EarTopY As Double, ByRef EarBotY As Double, ByRef Template As String)

        Dim xyTopOfEarRight, xyBotOfEarRight As HEADNECK1.xy
        Dim xyTopOfEarLeft, xyBotOfEarLeft As HEADNECK1.xy
        Dim xyCen As HEADNECK1.xy
        Dim HalfEarWidth As Double
        Dim nStartAng, nDeltaAng As Double
        Dim nRad, nEndAng As Double

        nStartAng = 0
        nDeltaAng = -180

        If optModifiedChinStrap.Checked = False Then
            HalfEarWidth = 0.125
            nRad = 0.125
        Else
            HalfEarWidth = 0.09375
            nRad = 0.09375
        End If

        PR_SetLayer(Template)

        ' Set points and Draw Right of Ear
        PR_MakeXY(xyTopOfEarRight, EarX + HalfEarWidth, EarTopY)
        PR_MakeXY(xyBotOfEarRight, EarX + HalfEarWidth, EarBotY)
        PR_DrawLine(xyTopOfEarRight, xyBotOfEarRight)
        PR_AddEntityID("EarRight")

        ' Set points and Draw Left of Ear
        PR_MakeXY(xyTopOfEarLeft, EarX - HalfEarWidth, EarTopY)
        PR_MakeXY(xyBotOfEarLeft, EarX - HalfEarWidth, EarBotY)
        PR_DrawLine(xyTopOfEarLeft, xyBotOfEarLeft)
        PR_AddEntityID("EarLeft")

        ' Draw Top Curve of Ear
        PR_MakeXY(xyCen, EarX, EarTopY)
        PR_DrawArc(xyCen, xyTopOfEarRight, xyTopOfEarLeft)
        PR_AddEntityID("EarTop")

        'Draw Bottom Curve of Ear
        PR_MakeXY(xyCen, EarX, EarBotY)
        PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "arc" & HEADNECK1.QC & "xyStart.x +" & Str(xyCen.X) & HEADNECK1.CC & "xyStart.y +" & Str(xyCen.Y) & HEADNECK1.CC & Str(nRad) & HEADNECK1.CC & Str(nEndAng) & HEADNECK1.CC & Str(nDeltaAng) & ");")
        PR_AddEntityID("EarBottom")
        PR_DrawArc(xyCen, xyBotOfEarLeft, xyBotOfEarRight)
        PR_SetLayer("TemplateLeft")

    End Sub

    Private Sub PR_DrawEars()

        Dim xyTextPoint As HEADNECK1.xy
        Dim YEarStep As Double
        Dim EarBotY, EarTopY, EarX As Double

        If optModifiedChinStrap.Checked = True Or optFaceMask.Checked = True Or optOpenFaceMask.Checked = True Then
            If optModifiedChinStrap.Checked = False Then
                YEarStep = 0.125
            ElseIf optModifiedChinStrap.Checked = True Then
                YEarStep = 0.1875
            End If

            If chkEarSize.CheckState = 0 Then
                If optModifiedChinStrap.Checked = False Then
                    EarTopY = Val(txtLowerEarHeight.Text) - YEarStep
                    EarBotY = Val(txtMouthHeight.Text) + YEarStep
                    EarX = -0.125
                    txtTopLeftEar.Text = CStr(EarTopY + YEarStep)
                    txtTopRightEar.Text = CStr(EarTopY + YEarStep)
                ElseIf optModifiedChinStrap.Checked = True Then
                    EarTopY = Val(txtCSEarTopHeight.Text) - YEarStep
                    EarBotY = Val(txtCSEarBotHeight.Text) + YEarStep
                    EarX = 0.09375
                    txtTopLeftEar.Text = CStr(EarTopY + YEarStep)
                    txtTopRightEar.Text = CStr(EarTopY + YEarStep)
                End If

                If chkLeftEarClosed.CheckState = 0 Then
                    PR_DrawEarHole(EarX, EarTopY, EarBotY, "TemplateLeft")
                ElseIf chkLeftEarClosed.CheckState = 1 Then
                    PR_SetLayer("Notes")
                    PR_MakeXY(xyTextPoint, 0.25, EarBotY)
                    PR_DrawText("Left Ear Closed", xyTextPoint, 0.1, 0)
                    PR_SetLayer("TemplateLeft")
                End If

                If chkRightEarClosed.CheckState = 0 Then
                    PR_DrawEarHole(EarX, EarTopY, EarBotY, "TemplateRight")
                    PR_SetLayer("TemplateLeft")
                ElseIf chkRightEarClosed.CheckState = 1 Then
                    PR_SetLayer("Notes")
                    PR_MakeXY(xyTextPoint, 0.25, EarBotY + 0.25)
                    PR_DrawText("Right Ear Closed", xyTextPoint, 0.1, 0)
                    PR_SetLayer("TemplateLeft")
                End If

            ElseIf chkEarSize.CheckState = 1 Then
                ' Left Ear
                If Val(txtLeftEarLength.Text) = 0 Then
                    If optModifiedChinStrap.Checked = False Then
                        EarTopY = Val(txtLowerEarHeight.Text) - YEarStep
                        EarBotY = Val(txtMouthHeight.Text) + YEarStep
                        EarX = -0.125
                        txtTopLeftEar.Text = CStr(EarTopY + YEarStep)
                    ElseIf optModifiedChinStrap.Checked = True Then
                        EarTopY = Val(txtCSEarTopHeight.Text) - YEarStep
                        EarBotY = Val(txtCSEarBotHeight.Text) + YEarStep
                        EarX = 0.09375
                        txtTopLeftEar.Text = CStr(EarTopY + YEarStep)
                    End If

                ElseIf Val(txtLeftEarLength.Text) > 0 Then
                    If optModifiedChinStrap.Checked = False Then
                        EarBotY = Val(txtMouthHeight.Text) + YEarStep
                        EarX = -0.125
                        'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        EarTopY = EarBotY + FN_CmToInches(txtUnits, txtLeftEarLength) - 0.25
                        txtTopLeftEar.Text = CStr(EarTopY + YEarStep)
                    ElseIf optModifiedChinStrap.Checked = True Then
                        EarBotY = Val(txtCSEarBotHeight.Text) + YEarStep
                        EarX = 0.09375
                        'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        EarTopY = EarBotY + FN_CmToInches(txtUnits, txtLeftEarLength) - 0.25
                        txtTopLeftEar.Text = CStr(EarTopY + YEarStep)
                    End If
                End If

                If chkLeftEarClosed.CheckState = 0 Then
                    PR_DrawEarHole(EarX, EarTopY, EarBotY, "TemplateLeft")
                ElseIf chkLeftEarClosed.CheckState = 1 Then
                    PR_SetLayer("Notes")
                    PR_MakeXY(xyTextPoint, 0.25, EarBotY)
                    PR_DrawText("Left Ear Closed", xyTextPoint, 0.1, 0)
                    PR_SetLayer("TemplateLeft")
                End If

                ' Right Ear
                If Val(txtRightEarLength.Text) = 0 Then
                    If optModifiedChinStrap.Checked = False Then
                        EarTopY = Val(txtLowerEarHeight.Text) - YEarStep
                        EarBotY = Val(txtMouthHeight.Text) + YEarStep
                        EarX = -0.125
                        txtTopRightEar.Text = CStr(EarTopY + 0.125)
                    ElseIf optModifiedChinStrap.Checked = True Then
                        EarTopY = Val(txtCSEarTopHeight.Text) - YEarStep
                        EarBotY = Val(txtCSEarBotHeight.Text) + YEarStep
                        EarX = 0.09375
                        txtTopRightEar.Text = CStr(EarTopY + YEarStep)
                    End If
                ElseIf Val(txtRightEarLength.Text) > 0 Then
                    If optModifiedChinStrap.Checked = False Then
                        EarBotY = Val(txtMouthHeight.Text) + YEarStep
                        EarX = -0.125
                        'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        EarTopY = EarBotY + FN_CmToInches(txtUnits, txtRightEarLength) - 0.25
                        txtTopRightEar.Text = CStr(EarTopY + YEarStep)
                    ElseIf optModifiedChinStrap.Checked = True Then
                        EarBotY = Val(txtCSEarBotHeight.Text) + YEarStep
                        EarX = 0.09375
                        'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        EarTopY = EarBotY + FN_CmToInches(txtUnits, txtRightEarLength) - 0.25
                        txtTopRightEar.Text = CStr(EarTopY + YEarStep)
                    End If
                End If
                If chkRightEarClosed.CheckState = 0 Then
                    PR_DrawEarHole(EarX, EarTopY, EarBotY, "TemplateRight")
                    PR_SetLayer("TemplateLeft")
                ElseIf chkRightEarClosed.CheckState = 1 Then
                    PR_SetLayer("Notes")
                    PR_MakeXY(xyTextPoint, 0.25, EarBotY)
                    PR_DrawText("Right Ear Closed", xyTextPoint, 0.1, 0)
                    PR_SetLayer("TemplateLeft")
                End If
            End If
        End If
    End Sub

    Private Sub PR_DrawEyeFlaps()
        ' This procedure Draws the eye Flap
        ' Globals Variables
        '
        '       xyNoseDiagTop As xy, xyNoseDiagBot As xy
        '
        Dim xyFlapTopLeft, xyFlapTopRight As HEADNECK1.xy
        Dim xySideLineInt, xyFlapBottomLeft, xyFlapBottomRight, xyBotLineInt As HEADNECK1.xy
        Dim FlapX, DartAng, FlapY As Double
        Dim PI As Double
        Dim Notes As HEADNECK1.xy
        Dim y2, X2, X1, Y1, Dist As Double
        Dim Point1, Point2 As HEADNECK1.xy
        Dim RightXOffSet, LeftXOffSet, YOffSet As Double
        Dim Left5, Left3, Left1, Left2, Left4, Left6 As HEADNECK1.xy
        Dim Right5, Right3, Right1, Right2, Right4, Right6 As HEADNECK1.xy
        Dim RightYOffSet, LeftYOffSet, XOffset As Double

        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.CQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        HEADNECK1.CQ = HEADNECK1.CC & HEADNECK1.QQ
        PI = 3.141592654
        X1 = Val(txtBotOfEyeX.Text)
        Y1 = Val(txtBotOfEyeY.Text)
        Dist = Val(txtEyeWidth.Text)

        ' Set Top Points
        PR_MakeXY(xyFlapTopLeft, X1 + 0.25, Y1 + 1.125)
        PR_MakeXY(xyFlapTopRight, X1 + Dist + 0.25, Y1 + 1.125)

        ' Calculate Bottom points
        PR_MakeXY(xyFlapBottomRight, xyFlapTopRight.X, Y1 - 0.875)
        PR_MakeXY(xyFlapBottomLeft, xyFlapTopLeft.X, Y1 - 0.875)

        'Check if Flap intersect nose line
        'Step 1 is to extend nose line to ensure intersection

        Dim iBot, iSide, Intersection As Short
        Dim aAngle As Double
        Dim xyPt1, xyPt2 As HEADNECK1.xy

        aAngle = FN_CalcAngle(xyNoseDiagBot, xyNoseDiagTop)
        PR_CalcPolar(xyNoseDiagTop, aAngle, 20, xyPt1)
        aAngle = FN_CalcAngle(xyNoseDiagTop, xyNoseDiagBot)
        PR_CalcPolar(xyNoseDiagBot, aAngle, 20, xyPt2)

        iSide = FN_LinLinInt(xyPt1, xyPt2, xyFlapBottomLeft, xyFlapTopLeft, xySideLineInt)
        iBot = FN_LinLinInt(xyPt1, xyPt2, xyFlapBottomLeft, xyFlapBottomRight, xyBotLineInt)

        'Cut only if no nose covering
        If iBot = True And iSide = True And chkNoseCovering.CheckState = 0 Then
            Intersection = True
        Else
            Intersection = False
        End If


        ' Set OffSets
        LeftYOffSet = -0.25
        RightYOffSet = 2.5
        XOffset = TopArcOpp(CInt(txtRadiusNo.Text))
        If chkLeftEyeFlap.CheckState = 1 Or chkRightEyeFlap.CheckState = 1 Then
            'Indicate position on Pattern
            PR_SetLayer("Notes")
            PR_MakeXY(Point1, xyFlapTopLeft.X + 0.125, xyFlapTopLeft.Y)
            PR_MakeXY(Point2, xyFlapTopLeft.X, xyFlapTopLeft.Y - 0.125)
            PR_DrawLine(Point1, xyFlapTopLeft)
            PR_DrawLine(Point2, xyFlapTopLeft)

            PR_MakeXY(Point1, xyFlapTopRight.X - 0.125, xyFlapTopRight.Y)
            PR_MakeXY(Point2, xyFlapTopRight.X, xyFlapTopRight.Y - 0.125)
            PR_DrawLine(Point1, xyFlapTopRight)
            PR_DrawLine(Point2, xyFlapTopRight)

            PR_MakeXY(Point1, xyFlapBottomRight.X - 0.125, xyFlapBottomRight.Y)
            PR_MakeXY(Point2, xyFlapTopRight.X, xyFlapBottomRight.Y + 0.125)
            PR_DrawLine(Point1, xyFlapBottomRight)
            PR_DrawLine(Point2, xyFlapBottomRight)

            If Intersection Then
                PR_MakeXY(Point1, xyBotLineInt.X + 0.125, xyBotLineInt.Y)
                PR_DrawLine(Point1, xyBotLineInt)
            Else
                PR_MakeXY(Point1, xyFlapBottomLeft.X + 0.125, xyFlapBottomLeft.Y)
                PR_MakeXY(Point2, xyFlapBottomLeft.X, xyFlapBottomLeft.Y + 0.125)
                PR_DrawLine(Point1, xyFlapBottomLeft)
                PR_DrawLine(Point2, xyFlapBottomLeft)
            End If
            PR_MakeXY(Right1, xyFlapTopLeft.X + RightXOffSet, xyFlapTopLeft.Y + YOffSet)
            PR_MakeXY(Right2, xyFlapTopRight.X + RightXOffSet, xyFlapTopRight.Y + YOffSet)
            PR_MakeXY(Right3, xyFlapBottomRight.X + RightXOffSet, xyFlapBottomRight.Y + YOffSet)
            PR_MakeXY(Right4, xyFlapBottomLeft.X + RightXOffSet, xyFlapBottomLeft.Y + YOffSet)

        End If

        ' Draw Left Eye
        If chkLeftEyeFlap.CheckState = 1 Then
            PR_SetLayer("TemplateLeft")
            PR_MakeXY(Left1, xyFlapTopLeft.X - XOffset, xyFlapTopLeft.Y + LeftYOffSet)
            PR_MakeXY(Left2, xyFlapTopRight.X - XOffset, xyFlapTopRight.Y + LeftYOffSet)
            PR_MakeXY(Left3, xyFlapBottomRight.X - XOffset, xyFlapBottomRight.Y + LeftYOffSet)
            PR_MakeXY(Left4, xyFlapBottomLeft.X - XOffset, xyFlapBottomLeft.Y + LeftYOffSet)
            If Intersection Then
                PR_MakeXY(Left5, xySideLineInt.X - XOffset, xySideLineInt.Y + LeftYOffSet)
                PR_MakeXY(Left6, xyBotLineInt.X - XOffset, xyBotLineInt.Y + LeftYOffSet)
            End If

            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "polyline" & HEADNECK1.QQ)
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Left1.X) & HEADNECK1.CC & "xyStart.y+" & Str(Left1.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Left2.X) & HEADNECK1.CC & "xyStart.y+" & Str(Left2.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Left3.X) & HEADNECK1.CC & "xyStart.y+" & Str(Left3.Y))
            If Intersection Then
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Left6.X) & HEADNECK1.CC & "xyStart.y+" & Str(Left6.Y))
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Left5.X) & HEADNECK1.CC & "xyStart.y+" & Str(Left5.Y))
            Else
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Left4.X) & HEADNECK1.CC & "xyStart.y+" & Str(Left4.Y))
            End If
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Left1.X) & HEADNECK1.CC & "xyStart.y+" & Str(Left1.Y))
            PrintLine(CInt(txtfNum.Text), ");")
            ''To Draw Polyline
            Dim Count As Integer = 5
            If Intersection Then
                Count = 6
            End If
            Dim PolyStart(Count), PolyEnd(Count) As Double
            PolyStart(1) = Left1.X
            PolyEnd(1) = Left1.Y
            PolyStart(2) = Left2.X
            PolyEnd(2) = Left2.Y
            PolyStart(3) = Left3.X
            PolyEnd(3) = Left3.Y
            If Intersection Then
                PolyStart(4) = Left6.X
                PolyEnd(4) = Left6.Y
                PolyStart(5) = Left5.X
                PolyEnd(5) = Left5.Y
            Else
                PolyStart(4) = Left4.X
                PolyEnd(4) = Left4.Y
            End If
            PolyStart(Count) = Left1.X
            PolyEnd(Count) = Left1.Y
            PR_DrawPoly(PolyStart, PolyEnd, Count)
            PR_AddEntityID("LeftEyeFlap")

            ' Add notes
            Left1.X = Left1.X + 0.1
            Left1.Y = Left1.Y - 0.1
            PR_TemplateDetails(Left1, "Left Eye Flap")

        End If

        ' Draw Right Eye
        If chkRightEyeFlap.CheckState = 1 Then
            PR_SetLayer("TemplateRight")
            PR_MakeXY(Right1, xyFlapTopLeft.X - XOffset, xyFlapTopLeft.Y + RightYOffSet)
            PR_MakeXY(Right2, xyFlapTopRight.X - XOffset, xyFlapTopRight.Y + RightYOffSet)
            PR_MakeXY(Right3, xyFlapBottomRight.X - XOffset, xyFlapBottomRight.Y + RightYOffSet)
            PR_MakeXY(Right4, xyFlapBottomLeft.X - XOffset, xyFlapBottomLeft.Y + RightYOffSet)
            If Intersection Then
                PR_MakeXY(Right5, xySideLineInt.X - XOffset, xySideLineInt.Y + RightYOffSet)
                PR_MakeXY(Right6, xyBotLineInt.X - XOffset, xyBotLineInt.Y + RightYOffSet)
            End If

            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "polyline" & HEADNECK1.QQ)
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Right1.X) & HEADNECK1.CC & "xyStart.y+" & Str(Right1.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Right2.X) & HEADNECK1.CC & "xyStart.y+" & Str(Right2.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Right3.X) & HEADNECK1.CC & "xyStart.y+" & Str(Right3.Y))
            If Intersection Then
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Right6.X) & HEADNECK1.CC & "xyStart.y+" & Str(Right6.Y))
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Right5.X) & HEADNECK1.CC & "xyStart.y+" & Str(Right5.Y))
            Else
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Right4.X) & HEADNECK1.CC & "xyStart.y+" & Str(Right4.Y))
            End If
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Right1.X) & HEADNECK1.CC & "xyStart.y+" & Str(Right1.Y))
            PrintLine(CInt(txtfNum.Text), ");")
            ''To Draw Polyline
            Dim Count As Integer = 5
            If Intersection Then
                Count = 6
            End If
            Dim PolyStart(Count), PolyEnd(Count) As Double
            PolyStart(1) = Right1.X
            PolyEnd(1) = Right1.Y
            PolyStart(2) = Right2.X
            PolyEnd(2) = Right2.Y
            PolyStart(3) = Right3.X
            PolyEnd(3) = Right3.Y
            If Intersection Then
                PolyStart(4) = Right6.X
                PolyEnd(4) = Right6.Y
                PolyStart(5) = Right5.X
                PolyEnd(5) = Right5.Y
            Else
                PolyStart(4) = Right4.X
                PolyEnd(4) = Right4.Y
            End If
            PolyStart(Count) = Right1.X
            PolyEnd(Count) = Right1.Y
            PR_DrawPoly(PolyStart, PolyEnd, Count)
            PR_AddEntityID("RightEyeFlap")

            ' Add notes
            Right1.X = Right1.X + 0.1
            Right1.Y = Right1.Y - 0.1

            PR_TemplateDetails(Right1, "Right Eye Flap")
        End If

    End Sub

    Private Sub PR_DrawEyeOpening()

        'Start To form Eye Opening, Jobst Procedure No.5

        Dim xyTopOfEye, xyBotOfEye As HEADNECK1.xy
        Dim xyTopOfEyeStart, xyBotOfEyeStart As HEADNECK1.xy
        Dim xyTopOfEyeEnd, xyBotOfEyeEnd As HEADNECK1.xy
        Dim xyCen, xyCen2 As HEADNECK1.xy
        Dim i As Short
        Dim Tempx, TempY As Double
        Dim MidToEyeBot, EyeWidthFromFace As Double
        Dim nStartAng, nDeltaAng As Object
        Dim nRad, nEndAng As Double

        Tempx = BotFaceAdj(CInt(txtRadiusNo.Text))
        PR_MakeXY(xyTopOfEye, Tempx, -(System.Math.Abs(CDbl(txtMidToEyeTop.Text))))
        PR_MakeXY(xyBotOfEye, Tempx, -(System.Math.Abs(CDbl(txtMidToEyeTop.Text))) - 0.25)

        ' Check Chart 2 for Width of eye from faceline
        For i = 1 To 16
            If FaceMaskChart2(1, i) = Val(txtCircumferenceTotal.Text) Then
                EyeWidthFromFace = FaceMaskChart2(3, i)
                txtEyeWidth.Text = CStr(EyeWidthFromFace)
                Exit For
            End If
        Next i


        ' Set points for top of eye
        PR_MakeXY(xyTopOfEyeStart, xyTopOfEye.X + 0.5, xyTopOfEye.Y)
        PR_MakeXY(xyTopOfEyeEnd, xyTopOfEye.X + EyeWidthFromFace - 0.125, xyTopOfEye.Y)

        ' Set points for bottom of eye
        PR_MakeXY(xyBotOfEyeStart, xyBotOfEye.X + 0.5, xyBotOfEye.Y)
        PR_MakeXY(xyBotOfEyeEnd, xyBotOfEye.X + EyeWidthFromFace - 0.125, xyBotOfEye.Y)

        ' Set points for Left Curve of eye
        PR_MakeXY(xyCen, xyTopOfEyeStart.X, xyTopOfEyeStart.Y - 0.125)

        ' Set points for Right Curve of eye
        PR_MakeXY(xyCen2, xyTopOfEyeEnd.X, xyTopOfEyeEnd.Y - 0.125)
        'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nStartAng = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nDeltaAng = 90
        nRad = 0.125

        ' Set txtLowerEarHeight with the y-axis value
        txtLowerEarHeight.Text = CStr(xyBotOfEyeEnd.Y)

        ' Set Coords of Bottom of Eye
        txtBotOfEyeX.Text = CStr(xyBotOfEye.X)
        txtBotOfEyeY.Text = CStr(xyBotOfEye.Y)


        If optFaceMask.Checked = True Or (optOpenFaceMask.Checked = True And chkEyes.CheckState = 1) Then
            ' Close Face Line in front of eyes
            If chkNoseCovering.CheckState = 0 Then
                PR_DrawLine(xyTopOfEye, xyBotOfEye)
                PR_AddEntityID("NoseBridge")
            End If

            PR_DrawLine(xyTopOfEyeStart, xyTopOfEyeEnd)
            PR_AddEntityID("EyeTop")

            PR_DrawLine(xyBotOfEyeStart, xyBotOfEyeEnd)
            PR_AddEntityID("EyeBottom")

            PR_DrawArc(xyCen, xyTopOfEyeStart, xyBotOfEyeStart)
            PR_AddEntityID("EyeLeft")

            'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "arc" & HEADNECK1.QC & "xyStart.x +" & Str(xyCen2.X) & HEADNECK1.CC & "xyStart.y +" & Str(xyCen.Y) & HEADNECK1.CC & Str(nRad) & HEADNECK1.CC & Str(nEndAng) & HEADNECK1.CC & Str(nDeltaAng) & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nDeltaAng = -90
            PR_AddEntityID("EyeRightTop")
            PR_DrawArc(xyCen2, xyBotOfEyeEnd, xyTopOfEyeEnd)
            'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "arc" & HEADNECK1.QC & "xyStart.x +" & Str(xyCen2.X) & HEADNECK1.CC & "xyStart.y +" & Str(xyCen.Y) & HEADNECK1.CC & Str(nRad) & HEADNECK1.CC & Str(nEndAng) & HEADNECK1.CC & Str(nDeltaAng) & ");")
            PR_AddEntityID("EyeRightBottom")

        End If

        ' End Drawing Eye Opening

    End Sub

    Private Sub PR_DrawForeHead()

        Dim xyForeHeadBot, xyForeHeadTop As HEADNECK1.xy
        Dim Tempx, TempY As Double

        If chkOpenHeadMask.CheckState = 0 Then
            Tempx = -LeftArcAdjOut(CInt(txtRadiusNo.Text))
            TempY = LeftArcOppOut(CInt(txtRadiusNo.Text))
            PR_MakeXY(xyForeHeadBot, Tempx, 0)
            PR_MakeXY(xyForeHeadTop, Tempx, TempY)

            PR_DrawLine(xyForeHeadBot, xyForeHeadTop)
            PR_AddEntityID("ForeHead")
        End If
    End Sub

    Private Sub PR_DrawHeadBand()
        Dim xyStartLeftArc, xyCentre, xyEndLeftArc As HEADNECK1.xy
        Dim xyStartRightArc, xyEndRightArc As HEADNECK1.xy
        Dim HeadBand As Double
        Dim RightArcEnd, RightArcStart, RightArcStep As Double
        Dim Across, Upwards As Double
        Dim xyTmp, xyTextPoint, xyNotes, xyTmp1 As HEADNECK1.xy

        Dim nRadius As Double
        Dim iError As Short

        Dim AcrossOut, UpwardsOut As Double


        If optHeadBand.Checked = True Then

            AcrossOut = LeftArcAdjOut(CInt(txtRadiusNo.Text))
            UpwardsOut = LeftArcOppOut(CInt(txtRadiusNo.Text))
            PR_MakeXY(xyStartLeftArc, -AcrossOut, UpwardsOut)

            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Upwards = FN_CmToInches(txtUnits, txtHeadBandDepth) + 0.5
            RightArcStart = 9 / 2.54
            RightArcEnd = 17.05 / 2.54
            RightArcStep = (RightArcEnd - RightArcStart) / 16
            RightArcStep = RightArcStep * (Val(txtRadiusNo.Text) - 1)

            PR_MakeXY(xyCentre, 0, 0)

            nRadius = FN_CalcLength(xyStartLeftArc, xyCentre)
            PR_MakeXY(xyTmp, 0, Upwards)
            PR_MakeXY(xyTmp1, xyStartLeftArc.X - 10, Upwards)
            iError = FN_CirLinInt(xyTmp, xyTmp1, xyCentre, nRadius, xyEndLeftArc)

            PR_MakeXY(xyStartRightArc, RightArcStart + RightArcStep, 0)
            Across = System.Math.Sqrt((System.Math.Abs(xyStartRightArc.X) * System.Math.Abs(xyStartRightArc.X)) - (Upwards * Upwards))
            PR_MakeXY(xyEndRightArc, Across, Upwards)

            PR_DrawArc(xyCentre, xyEndLeftArc, xyStartLeftArc)
            PR_AddEntityID("HeadBandLeftArc")
            PR_DrawArc(xyCentre, xyStartRightArc, xyEndRightArc)
            PR_AddEntityID("HeadBandRightArc")
            PR_DrawLine(xyEndLeftArc, xyEndRightArc)
            PR_AddEntityID("HeadBandTop")

            PR_MakeXY(xyTmp, xyStartLeftArc.X, xyStartRightArc.Y)
            PR_DrawLine(xyTmp, xyStartRightArc)
            PR_AddEntityID("HeadBandBottom")
            PR_DrawLine(xyTmp, xyStartLeftArc)
            PR_AddEntityID("HeadBandLeftLine")

            'add closure
            PR_SetLayer("Notes")
            PR_MakeXY(xyTextPoint, xyEndRightArc.X - 0.1, Upwards / 2)
            PR_DrawText("VELCRO", xyTextPoint, 0.1, FN_CalcAngle(xyEndRightArc, xyStartRightArc))

            ' Add notes
            PR_MakeXY(xyTextPoint, xyEndLeftArc.X + 0.125, xyEndLeftArc.Y - 0.13)
            PR_TemplateDetails(xyTextPoint, "Head Band")


        End If
    End Sub

    Private Sub PR_DrawLeftArc()
        ' This Procedure Draws the Left Arc of the Face Mask
        ' It uses the value held in txtRadiusNo to determine which
        ' Arc to Draw.  The Value Held in txtRadius must lie between
        ' 1 and 19 inclusive.

        Dim xyEndArc, xyStartArc, xyCentreArc As HEADNECK1.xy
        Dim AcrossIn, UpwardsIn As Double
        Dim AcrossOut, UpwardsOut As Double

        If chkOpenHeadMask.CheckState = 0 Then
            PR_MakeXY(xyCentreArc, 0, 0)

            AcrossIn = LeftArcAdjIn(CInt(txtRadiusNo.Text))
            UpwardsIn = LeftArcOppIn(CInt(txtRadiusNo.Text))

            AcrossOut = LeftArcAdjOut(CInt(txtRadiusNo.Text))
            UpwardsOut = LeftArcOppOut(CInt(txtRadiusNo.Text))

            PR_MakeXY(xyStartArc, -AcrossIn, UpwardsIn)
            PR_MakeXY(xyEndArc, -AcrossOut, UpwardsOut)
            PR_DrawArc(xyCentreArc, xyStartArc, xyEndArc)
            PR_AddEntityID("LeftArc")
        End If
    End Sub

    Private Sub PR_DrawLeftCutOut()
        Dim AcrossLeft, UpwardsLeft As Double
        Dim AcrossRight, UpwardsRight As Double
        Dim CentreX, CentreY As Double
        Dim LeftTip, Centre, RightTip As HEADNECK1.xy

        If chkOpenHeadMask.CheckState = 0 Then
            CentreX = -3.8 / 2.54
            CentreY = 3.8 / 2.54

            PR_MakeXY(Centre, CentreX, CentreY)

            AcrossLeft = TopArcAdj(CInt(txtRadiusNo.Text))
            UpwardsLeft = TopArcOpp(CInt(txtRadiusNo.Text))
            AcrossRight = TopRightArcAdjIn(CInt(txtRadiusNo.Text))
            UpwardsRight = TopRightArcOppIn(CInt(txtRadiusNo.Text))

            PR_MakeXY(LeftTip, -AcrossLeft, UpwardsLeft)
            PR_MakeXY(RightTip, -AcrossRight, UpwardsRight)

            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "polyline" & HEADNECK1.QQ)
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(LeftTip.X) & HEADNECK1.CC & "xyStart.y+" & Str(LeftTip.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Centre.X) & HEADNECK1.CC & "xyStart.y+" & Str(Centre.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(RightTip.X) & HEADNECK1.CC & "xyStart.y+" & Str(RightTip.Y))
            PrintLine(CInt(txtfNum.Text), ");")
            ''To Draw Polyline
            Dim StartPt(3), EndPt(3) As Double
            StartPt(1) = LeftTip.X
            EndPt(1) = LeftTip.Y
            StartPt(2) = Centre.X
            EndPt(2) = Centre.Y
            StartPt(3) = RightTip.X
            EndPt(3) = RightTip.Y
            PR_DrawPoly(StartPt, EndPt, 3)
            PR_AddEntityID("LeftCutOut")

        End If
    End Sub

    Private Sub PR_DrawLine(ByRef xyStart As HEADNECK1.xy, ByRef xyFinish As HEADNECK1.xy)
        'To the DRAFIX macro file (given by the global txtfNum).
        'Write the syntax to draw a LINE between two points.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    HEADNECK1.XY      xyStart
        '    HANDLE  hEnt
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
            ''Create a Line object
            'Dim acLine As Line = New Line(New Point3d(xyStart.X, xyStart.Y, 0), New Point3d(xyFinish.X, xyFinish.Y, 0))
            Dim acLine As Line = New Line(New Point3d(xyStart.X + xyInsertion.X, xyStart.Y + xyInsertion.Y, 0), New Point3d(xyFinish.X + xyInsertion.X, xyFinish.Y + xyInsertion.Y, 0))
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                acLine.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
            End If

            '' Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)
            '' Save the new object to the database
            acTrans.Commit()
        End Using

        PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(")
        PrintLine(CInt(txtfNum.Text), HEADNECK1.QQ & "line" & HEADNECK1.QC)
        PrintLine(CInt(txtfNum.Text), "xyStart.x+" & Str(xyStart.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyStart.Y) & HEADNECK1.CC)
        PrintLine(CInt(txtfNum.Text), "xyStart.x+" & Str(xyFinish.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyFinish.Y))
        PrintLine(CInt(txtfNum.Text), ");")

    End Sub
    Sub PR_DrawPoly(ByRef ChinStrapX As Double(), ByRef ChinStrapY As Double(), ByRef nCount As Short)
        'To the DRAFIX macro file (given by the global fNum)
        'write the syntax to draw a POLYLINE through the points
        'given in Profile.
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    XY      xyStart
        '    HANDLE  hEnt
        '
        'Note:-
        '    fNum, CC, QQ, NL are globals initialised by FN_Open
        '
        '
        Dim ii As Short

        'Exit if nothing to draw
        'If Profile.n <= 1 Then Exit Sub

        ''UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'PrintLine(fNum, "hEnt = AddEntity(" & QQ & "poly" & QCQ & "polyline" & QQ)
        'For ii = 1 To Profile.n
        '    'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '    PrintLine(fNum, CC & "xyStart.x+" & Str(Profile.X(ii)) & CC & "xyStart.y+" & Str(Profile.y(ii)))
        'Next ii
        ''UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'PrintLine(fNum, ");")
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                            OpenMode.ForWrite)

            '' Create a polyline with two segments (3 points)
            Using acPoly As Polyline = New Polyline()
                For ii = 1 To nCount
                    acPoly.AddVertexAt(ii - 1, New Point2d(xyInsertion.X + ChinStrapX(ii), xyInsertion.Y + ChinStrapY(ii)), 0, 0, 0)
                Next ii

                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acPoly.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
            End Using

            '' Save the new object to the database
            acTrans.Commit()
        End Using

    End Sub

    Private Sub PR_DrawLipCovering()

        ' This procedure Draws the Lip Covering

        Dim xyFlapTopRight, xyFlapTopLeft, xyFlapBottomMid As HEADNECK1.xy
        Dim xyFlapBottomRight, xyFlapBottomLeft, xyTextPoint As HEADNECK1.xy
        Dim xy2, Notes, xy1, xyTmp As HEADNECK1.xy
        Dim xy6, xyDartStart, xyDartEnd, xy5 As HEADNECK1.xy
        Dim aAngle, aStart, nOffSet As Double
        Dim y2, X2, PI, DistStepX, Upwards, X1, Y1, DistX As Double
        Dim XOffset, DistY, DistStepY, ChinVal, ChinCheckY As Double
        Dim nAlong As Double
        Dim Y3, X3, DartAngle As Double
        Dim i, ii, nn As Short
        Dim xyLipCoveringBotCen As HEADNECK1.xy
        Dim LipDistX, LipDistY As Double
        Dim xyDartInt As HEADNECK1.xy
        Dim length, nArc As Double
        Dim VelcroText As Object
        Dim TextPoint As HEADNECK1.xy

        PI = 3.141592654
        nAlong = 0.375
        '    BotRadius = .9 / 2.54
        '    TopRadius = 1.8 / 2.54


        Dim Point1, Point2 As HEADNECK1.xy
        If chkLipCovering.CheckState = 1 And optChinStrap.Checked = False And optModifiedChinStrap.Checked = False Then

            XOffset = -LeftArcAdjOut(CInt(txtRadiusNo.Text)) - 0.5

            ' Set Top Left Point
            Y1 = CDbl(txtNoseBottomY.Text)
            PR_MakeXY(xyFlapTopLeft, -LeftArcAdjOut(CInt(txtRadiusNo.Text)) + 0.375, Y1 + 0.5)


            ' Set Top Right Point
            PR_MakeXY(xyFlapTopRight, -0.875, xyFlapTopLeft.Y)

            ' Calculate Bottom Right point
            X1 = Val(txtDartStartX.Text)
            Y1 = Val(txtDartStartY.Text)
            X2 = xyFlapTopRight.X
            X3 = Val(txtDartEndX.Text)
            Y3 = Val(txtDartEndY.Text)
            PR_MakeXY(xyDartStart, X1, Y1)
            PR_MakeXY(xyDartEnd, X3, Y3)

            If X1 >= X2 Then
                PR_MakeXY(xyFlapBottomRight, xyFlapTopRight.X, Y1 + 0.25)
            Else
                DistX = System.Math.Abs(System.Math.Abs(X2) - System.Math.Abs(X1))
                DartAngle = FN_CalcAngle(xyDartStart, xyDartEnd) * (180 / PI)
                DartAngle = DartAngle * (PI / 180)
                Upwards = System.Math.Abs(DistX * System.Math.Tan(DartAngle))
                PR_MakeXY(xyFlapBottomRight, xyFlapTopRight.X, Y1 + Upwards + 0.25)
            End If


            ' Find Bottom Mid Point
            PR_MakeXY(xy5, ChinProfileX(5), ChinProfileY(5))
            PR_MakeXY(xy6, ChinProfileX(6), ChinProfileY(6))
            'PR_DrawMarker xy5
            'PR_DrawText "xy5", xy5, .05, 0
            'PR_DrawMarker xy6
            'PR_DrawText "xy6", xy6, .05, 0
            'PR_DrawMarker xyDartStart
            'PR_DrawText "xyDartStart", xyDartStart, .05, 0

            If xyDartStart.Y > xy6.Y Then
                'Calculate intersection point on toparc
                PR_MakeXY(xy1, xyChinTopCen.X, xyDartStart.Y)
                PR_MakeXY(xy2, xyChinTopCen.X - 10, xyDartStart.Y)
                'PR_DrawLine xy1, xy2
                'Print #txtfNum, "hEnt = AddEntity("; HEADNECK1.qq; "circle"; HEADNECK1.qq; ",xyStart.x+"; Str$(xyChinTopCen.x); CC; "xyStart.y+"; Str$(xyChinTopCen.y); ","; g_nChinTopRadius; ");"
                i = FN_CirLinInt(xy1, xy2, xyChinTopCen, g_nChinTopRadius, xyDartInt)
                'PR_DrawMarker xyDartInt
                'PR_DrawText "xyDartInt", xyDartInt, .05, 0
                aAngle = (FN_CalcAngle(xyChinTopCen, xy6) * (180 / PI)) - (FN_CalcAngle(xyChinTopCen, xyDartInt) * (180 / PI))
                nArc = ((PI / 180) * aAngle) * g_nChinTopRadius
                nOffSet = FN_CalcLength(xy6, xy5)
                If nArc > nAlong Then
                    'This is an impossible case so we don't deal with it
                ElseIf (nOffSet + nArc) > nAlong Then
                    PR_CalcPolar(xy6, (FN_CalcAngle(xy6, xy5) * (180 / PI)), nAlong - nArc, xyFlapBottomMid)
                    'PR_DrawMarker xyFlapBottomMid
                    'PR_DrawText "xyFlapBottomMid on line", xyFlapBottomMid, .05, 0
                Else
                    nOffSet = nAlong - (nOffSet + nArc)
                    aStart = FN_CalcAngle(xyChinBotCen, xy5) * (180 / PI)
                    aAngle = aStart - ((nOffSet / (2 * PI * g_nChinBotRadius)) * 360)
                    PR_CalcPolar(xyChinBotCen, aAngle, g_nChinBotRadius, xyFlapBottomMid)
                    'PR_DrawMarker xyFlapBottomMid
                    'PR_DrawText "xyFlapBottomMid on bottom arc", xyFlapBottomMid, .05, 0
                End If

            ElseIf xyDartStart.Y > xy5.Y Then
                'Calculate intersection point on line
                'PR_MakeXY(xyDartInt, xy6.X - (System.Math.Abs(xy6.Y - xyDartStart.Y) / System.Math.Tan(FN_CalcAngle(xy5, xy6) * (PI / 180))), xyDartStart.Y)
                PR_MakeXY(xyDartInt, xy6.X - (System.Math.Abs(xy6.Y - xyDartStart.Y) / System.Math.Tan(FN_CalcAngle(xy5, xy6))), xyDartStart.Y)
                nOffSet = FN_CalcLength(xyDartInt, xy5)
                If nOffSet < nAlong Then
                    aStart = FN_CalcAngle(xyChinBotCen, xy5) * (180 / PI)
                    aAngle = aStart - (((nAlong - nOffSet) / (2 * PI * g_nChinBotRadius)) * 360)
                    PR_CalcPolar(xyChinBotCen, aAngle, g_nChinBotRadius, xyFlapBottomMid)
                Else
                    PR_CalcPolar(xyDartInt, (FN_CalcAngle(xy6, xy5) * (180 / PI)), nAlong, xyFlapBottomMid)
                End If
                'PR_DrawMarker xyDartInt
                'PR_DrawText "xyDartInt", xyDartInt, .05, 0
            Else
                'Calculate intersection point on the bottom arc
                PR_MakeXY(xy1, xyChinBotCen.X, xyDartEnd.Y)
                PR_MakeXY(xy2, xyChinBotCen.X + 10, xyDartEnd.Y)
                'PR_DrawLine xy1, xy2
                'Print #txtfNum, "hEnt = AddEntity("; HEADNECK1.qq; "circle"; HEADNECK1.qq; ",xyStart.x+"; Str$(xyChinBotCen.x); CC; "xyStart.y+"; Str$(xyChinBotCen.y); ","; g_nChinBotRadius; ");"
                i = FN_CirLinInt(xy1, xy2, xyChinBotCen, g_nChinBotRadius, xyDartInt)
                'PR_DrawMarker xyDartInt
                'PR_DrawText "xyDartInt", xyDartInt, .05, 0
                aStart = FN_CalcAngle(xyChinBotCen, xyDartInt) * (180 / PI)
                aAngle = aStart - ((nAlong / (2 * PI * g_nChinBotRadius)) * 360)
                PR_CalcPolar(xyChinBotCen, aAngle, g_nChinBotRadius, xyFlapBottomMid)

            End If

            ' Adjust to include Lip
            LipProfileX(1) = xyFlapTopLeft.X
            LipProfileY(1) = xyFlapTopLeft.Y
            LipProfileX(2) = ChinProfileX(10)
            LipProfileY(2) = ChinProfileY(10)
            LipProfileX(3) = ChinProfileX(9)
            LipProfileY(3) = ChinProfileY(9)

            nn = 4
            For ii = 8 To 6 Step -1
                Select Case nn
                    Case 4
                        nOffSet = 0.08
                    Case 5
                        nOffSet = 0.125
                    Case 6
                        nOffSet = 0.08
                End Select
                PR_MakeXY(xyTmp, ChinProfileX(ii), ChinProfileY(ii))
                PR_CalcPolar(xyChinTopCen, (FN_CalcAngle(xyChinTopCen, xyTmp) * (180 / PI)), FN_CalcLength(xyChinTopCen, xyTmp) - nOffSet, xyTmp)
                LipProfileX(nn) = xyTmp.X
                LipProfileY(nn) = xyTmp.Y
                nn = nn + 1
            Next ii

            LipProfileX(7) = xyFlapBottomMid.X
            LipProfileY(7) = xyFlapBottomMid.Y


            ' Add Offset to curve
            For i = 1 To 7
                LipProfileX(i) = LipProfileX(i) + XOffset
            Next i

            'Indicate position on Pattern
            PR_SetLayer("Notes")

            PR_MakeXY(Point1, xyFlapTopLeft.X + 0.125, xyFlapTopLeft.Y)
            PR_MakeXY(Point2, xyFlapTopLeft.X, xyFlapTopLeft.Y - 0.125)
            PR_DrawLine(Point1, xyFlapTopLeft)
            PR_DrawLine(Point2, xyFlapTopLeft)

            PR_MakeXY(Point1, xyFlapTopRight.X - 0.125, xyFlapTopRight.Y)
            PR_MakeXY(Point2, xyFlapTopRight.X, xyFlapTopRight.Y - 0.125)
            PR_DrawLine(Point1, xyFlapTopRight)
            PR_DrawLine(Point2, xyFlapTopRight)

            PR_MakeXY(Point1, xyFlapBottomRight.X - 0.125, xyFlapBottomRight.Y)
            PR_MakeXY(Point2, xyFlapTopRight.X, xyFlapBottomRight.Y + 0.125)
            PR_DrawLine(Point1, xyFlapBottomRight)
            PR_DrawLine(Point2, xyFlapBottomRight)

            PR_MakeXY(Point1, xyFlapBottomMid.X + 0.125, xyFlapBottomMid.Y)
            PR_MakeXY(Point2, xyFlapBottomMid.X, xyFlapBottomMid.Y + 0.125)
            PR_DrawLine(Point1, xyFlapBottomMid)

            PR_SetLayer("TemplateLeft")

            ' Add OffSet to Points
            PR_MakeXY(xyFlapTopLeft, xyFlapTopLeft.X + XOffset, xyFlapTopLeft.Y)
            PR_MakeXY(xyFlapTopRight, xyFlapTopRight.X + XOffset, xyFlapTopRight.Y)
            PR_MakeXY(xyFlapBottomRight, xyFlapBottomRight.X + XOffset, xyFlapBottomRight.Y)
            PR_MakeXY(xyFlapBottomMid, xyFlapBottomMid.X + XOffset, xyFlapBottomMid.Y)

            nOffSet = FN_CalcLength(xyFlapBottomMid, xyFlapBottomRight) / 2
            aAngle = FN_CalcAngle(xyFlapBottomMid, xyFlapBottomRight) * (180 / PI)
            PR_CalcPolar(xyFlapBottomMid, aAngle, nOffSet, xyTmp)
            PR_CalcPolar(xyTmp, aAngle + 270, 9, xyLipCoveringBotCen)
            PR_DrawArc(xyLipCoveringBotCen, xyFlapBottomRight, xyFlapBottomMid)
            PR_AddEntityID("LipCoveringBot")


            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "polyline" & HEADNECK1.QQ)
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyFlapTopLeft.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyFlapTopLeft.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyFlapTopRight.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyFlapTopRight.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyFlapBottomRight.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyFlapBottomRight.Y))
            PrintLine(CInt(txtfNum.Text), ");")
            ''To Draw Polyline
            Dim PolyStart(3), PolyEnd(3) As Double
            PolyStart(1) = xyFlapTopLeft.X
            PolyEnd(1) = xyFlapTopLeft.Y
            PolyStart(2) = xyFlapTopRight.X
            PolyEnd(2) = xyFlapTopRight.Y
            PolyStart(3) = xyFlapBottomRight.X
            PolyEnd(3) = xyFlapBottomRight.Y
            PR_DrawPoly(PolyStart, PolyEnd, 3)

            PR_AddEntityID("LipCovering")

            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "fitted" & HEADNECK1.QQ)
            For ii = 1 To 7
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(LipProfileX(ii)) & HEADNECK1.CC & "xyStart.y+" & Str(LipProfileY(ii)))
            Next ii
            PrintLine(CInt(txtfNum.Text), ");")
            ''To Draw Polyline
            Dim StartPt(7), EndPt(7) As Double
            For ii = 1 To 7
                StartPt(ii) = LipProfileX(ii)
                EndPt(ii) = LipProfileY(ii)
            Next
            'PR_DrawPoly(StartPt, EndPt, 7)
            ''To Draw Spline
            Dim PointColl As Point3dCollection = New Point3dCollection()
            For ii = 1 To 7
                PointColl.Add(New Point3d(xyInsertion.X + LipProfileX(ii), xyInsertion.Y + LipProfileY(ii), 0))
            Next
            PR_DrawSpline(PointColl)
            PR_AddEntityID("LipCoveringCurve")

            ' Add Velcro
            PR_SetLayer("Notes")
            PR_MakeXY(xy1, xyFlapTopRight.X - 0.75, xyFlapTopRight.Y)
            PR_MakeXY(xy2, xy1.X, xyLipCoveringBotCen.Y)
            ii = FN_CirLinInt(xy1, xy2, xyLipCoveringBotCen, FN_CalcLength(xyLipCoveringBotCen, xyFlapBottomRight), xy2)
            PR_DrawLine(xy1, xy2)
            length = FN_CalcLength(xy1, xy2) + 0.25
            'UPGRADE_WARNING: Couldn't resolve default property of object VelcroText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VelcroText = FN_InchesToText(ARMEDDIA1.fnRoundInches(length))
            'UPGRADE_WARNING: Couldn't resolve default property of object VelcroText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VelcroText = VelcroText & Chr(34) & " Velcro"
            PR_MakeXY(TextPoint, xy1.X - 0.25, xy1.Y - 0.25)
            PR_DrawText(VelcroText, TextPoint, 0.1, -90 * (PI / 180))
            length = FN_CalcLength(xyFlapTopRight, xyFlapBottomRight) + 0.25
            'UPGRADE_WARNING: Couldn't resolve default property of object VelcroText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VelcroText = FN_InchesToText(ARMEDDIA1.fnRoundInches(length))
            'UPGRADE_WARNING: Couldn't resolve default property of object VelcroText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VelcroText = VelcroText & Chr(34) & " Velcro"
            PR_MakeXY(TextPoint, xyFlapTopRight.X - 0.25, xy1.Y - 0.25)
            PR_DrawText(VelcroText, TextPoint, 0.1, -90 * (PI / 180))

            ' Add Notes
            PR_MakeXY(xyTextPoint, xyFlapTopLeft.X + 0.25, xyFlapTopLeft.Y - 0.1)
            PR_TemplateDetails(xyTextPoint, "Lip Covering")

            PR_SetLayer("Construct")
            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "fitted" & HEADNECK1.QQ)
            For ii = 3 To 10
                PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(ChinProfileX(ii) + XOffset) & HEADNECK1.CC & "xyStart.y+" & Str(ChinProfileY(ii)))
            Next ii
            PrintLine(CInt(txtfNum.Text), ");")
            ''To Draw Polyline
            Dim xyStartPt(8), xyEndPt(8) As Double
            For ii = 3 To 10
                xyStartPt(ii - 2) = ChinProfileX(ii) + XOffset
                xyEndPt(ii - 2) = ChinProfileY(ii)
            Next
            '' PR_DrawPoly(xyStartPt, xyEndPt, 8)
            ''To Draw Spline
            Dim ptColl As Point3dCollection = New Point3dCollection()
            For ii = 3 To 10
                ptColl.Add(New Point3d(xyInsertion.X + ChinProfileX(ii) + XOffset, xyInsertion.Y + ChinProfileY(ii), 0))
            Next
            PR_DrawSpline(ptColl)
            PR_SetLayer("TemplateLeft")
        End If

    End Sub

    Private Sub PR_DrawMarker(ByRef xyPoint As HEADNECK1.xy)
        'Draw a Marker at the given point
        PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "marker" & HEADNECK1.QCQ & "xmarker" & HEADNECK1.QC & "xyStart.x+" & xyPoint.X & HEADNECK1.CC & "xyStart.y+" & xyPoint.Y & HEADNECK1.CC & "0.125);")
    End Sub

    Private Sub PR_DrawNeckAndDart()

        ' Procedure to Add Neck Piece, Jobst Procedure No.7

        Dim NeckCircum, FiguredNeck As Double
        Dim NeckRightSideTop, NeckLeftSideTop As HEADNECK1.xy
        Dim NeckRightSideBot, NeckLeftSideBot As HEADNECK1.xy
        Dim DartStart, DartEnd As HEADNECK1.xy
        Dim Tempx, TempY As Double
        Dim TempX2, TempY2 As Double
        Dim TempDepth, NeckDepth As Double
        Dim DartEndX, Dart, DartEndY As Double
        Dim Circum, LineLen As Double
        Dim ZipBit As Double
        Dim i As Short

        Circum = Val(txtRadiusNo.Text)
        ZipBit = 0
        ' Set For Zipper
        If chkZipper.CheckState = 1 And Val(txtRadiusNo.Text) > 3 Then
            Circum = Circum - 3
            ZipBit = 0.625
        ElseIf chkZipper.CheckState = 1 And Circum <= 3 Then
            Circum = Circum + 17
            ZipBit = 0.625
        End If

        ' Set For 2" wide Velcro
        If chkVelcro.CheckState = 1 Then
            Circum = Circum + 2
            ZipBit = -0.375
        End If


        LineLen = 1.35 / 2.54
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        NeckCircum = FN_CmToInches(txtUnits, txtCircOfNeck)

        ' Check Chart 3 for Figured Neck
        For i = 1 To 46
            If (NeckCircum >= FaceMaskChart3(1, i)) And (NeckCircum <= FaceMaskChart3(2, i)) And (NeckCircum < FaceMaskChart3(1, i + 1)) Then
                FiguredNeck = FaceMaskChart3(3, i)
            ElseIf NeckCircum >= 20 Then
                FiguredNeck = FaceMaskChart3(47, 3)
            End If
        Next i

        ' Top Right point
        Tempx = NeckRightAdj(CInt(txtRadiusNo.Text))
        TempY = NeckRightOpp(CInt(txtRadiusNo.Text))
        PR_MakeXY(NeckRightSideTop, Tempx - ZipBit, TempY)

        ' Top Left Point
        TempX2 = Tempx - FiguredNeck
        PR_MakeXY(NeckLeftSideTop, TempX2, TempY)

        ' Sets Min depth of 2" for adults and 1.5" Under ten
        If txtThroatToSternal.Text = "" Then
            If Val(txtAge.Text) >= 10 Then
                NeckDepth = 2
            Else
                NeckDepth = 1.5
            End If
        ElseIf txtThroatToSternal.Text <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            NeckDepth = FN_CmToInches(txtUnits, txtThroatToSternal)
        End If

        TempY2 = -(System.Math.Abs(TempY) + NeckDepth)
        PR_MakeXY(NeckRightSideBot, Tempx - ZipBit, TempY2)

        ' Bottom Left Point
        PR_MakeXY(NeckLeftSideBot, TempX2, TempY2)

        ' Set txtChinLeftBotX and txtChinLeftBotY
        txtChinLeftBotX.Text = CStr(NeckLeftSideBot.X)
        txtChinLeftBotY.Text = CStr(NeckLeftSideBot.Y)

        ' Draw Dart
        Dart = Tempx - FiguredNeck + 0.625
        PR_MakeXY(DartStart, Dart, TempY)
        txtDartStartX.Text = CStr(DartStart.X)
        txtDartStartY.Text = CStr(DartStart.Y)
        DartEndX = (RightArcAdjDown(Circum)) - LineLen
        DartEndY = -RightArcOppDown(Circum)
        PR_MakeXY(DartEnd, DartEndX, DartEndY)
        txtDartEndX.Text = CStr(DartEnd.X)
        txtDartEndY.Text = CStr(DartEnd.Y)

        ' Draw Neck Piece
        PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "polyline" & HEADNECK1.QQ)
        PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(DartEnd.X) & HEADNECK1.CC & "xyStart.y+" & Str(DartEnd.Y))
        PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(DartStart.X) & HEADNECK1.CC & "xyStart.y+" & Str(DartStart.Y))
        PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(NeckRightSideTop.X) & HEADNECK1.CC & "xyStart.y+" & Str(NeckRightSideTop.Y))
        PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(NeckRightSideBot.X) & HEADNECK1.CC & "xyStart.y+" & Str(NeckRightSideBot.Y))
        PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(NeckLeftSideBot.X) & HEADNECK1.CC & "xyStart.y+" & Str(NeckLeftSideBot.Y))
        PrintLine(CInt(txtfNum.Text), ");")
        ''To Draw Polyline
        Dim PolyStart(5), PolyEnd(5) As Double
        PolyStart(1) = DartEnd.X
        PolyEnd(1) = DartEnd.Y
        PolyStart(2) = DartStart.X
        PolyEnd(2) = DartStart.Y
        PolyStart(3) = NeckRightSideTop.X
        PolyEnd(3) = NeckRightSideTop.Y
        PolyStart(4) = NeckRightSideBot.X
        PolyEnd(4) = NeckRightSideBot.Y
        PolyStart(5) = NeckLeftSideBot.X
        PolyEnd(5) = NeckLeftSideBot.Y
        PR_DrawPoly(PolyStart, PolyEnd, 5)
        PR_AddEntityID("DartAndNeck")

        'If 2" velcro then add this
        If chkVelcro.CheckState = 1 Then
            PR_SetLayer("Notes")
            DartEnd.Y = DartStart.Y - 1
            DartEnd.X = DartEnd.X - 1.5
            PR_DrawText("2" & Chr(34) & " VELCRO", DartEnd, 0.1, 0)
            PR_SetLayer("TemplateLeft")
        End If



    End Sub

    Private Sub PR_DrawNose()

        Dim xyNoseTopLeft, xyNoseTopRight As HEADNECK1.xy
        'Dim xyNoseDiagTop As xy, xyNoseDiagBot As xy   'Now Global for use in drawing
        'eye flaps
        Dim xyNoseBotRight, xyNoseBotLeft, xyNoseBotRight2 As HEADNECK1.xy
        Dim xyTopCen, xyBotCen As HEADNECK1.xy
        Dim nRad, nStartAng, nDeltaAng, nEndAng As Double
        Dim Tempx, TempY As Double
        Dim TempX2, TempY2 As Double

        ' Top Part
        Tempx = Val(txtBotOfEyeX.Text)
        TempY = Val(txtBotOfEyeY.Text)
        PR_MakeXY(xyNoseTopLeft, Tempx, TempY)
        PR_MakeXY(xyNoseTopRight, Tempx + 0.125, TempY)

        ' Bottom Part
        TempY2 = Val(txtNoseBottomY.Text)
        PR_MakeXY(xyNoseBotLeft, xyNoseTopLeft.X, TempY2)
        PR_MakeXY(xyNoseBotRight, xyNoseTopLeft.X + 0.25, TempY2)

        ' Diagonal Line
        PR_MakeXY(xyNoseDiagTop, xyNoseTopRight.X + 0.125, xyNoseTopRight.Y - 0.125)
        PR_MakeXY(xyNoseDiagBot, xyNoseBotRight.X + 0.125, TempY2 + 0.125)

        ' Top Curve
        PR_MakeXY(xyTopCen, xyNoseTopRight.X, xyNoseDiagTop.Y)

        'Bottom Curve
        PR_MakeXY(xyBotCen, xyNoseBotRight.X, xyNoseDiagBot.Y)
        nStartAng = 0
        nDeltaAng = -90
        nRad = 0.125

        ' Draw Nose Opening if No nose covering is requested
        If chkNoseCovering.CheckState = 0 Then
            PR_DrawLine(xyNoseTopLeft, xyNoseTopRight)
            PR_AddEntityID("NoseTop")
            PR_DrawLine(xyNoseBotLeft, xyNoseBotRight)
            PR_AddEntityID("NoseBottom")
            PR_DrawLine(xyNoseDiagTop, xyNoseDiagBot)
            PR_AddEntityID("NoseSide")
            PR_DrawArc(xyTopCen, xyNoseDiagTop, xyNoseTopRight)
            PR_AddEntityID("NoseTopCurve")
            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "arc" & HEADNECK1.QC & "xyStart.x +" & Str(xyBotCen.X) & HEADNECK1.CC & "xyStart.y +" & Str(xyBotCen.Y) & HEADNECK1.CC & Str(nRad) & HEADNECK1.CC & Str(nEndAng) & HEADNECK1.CC & Str(nDeltaAng) & ");")
            PR_AddEntityID("NoseBottomCurve")
            PR_DrawArc(xyBotCen, xyNoseBotRight, xyNoseDiagBot)
        Else
            PR_MakeXY(xyNoseBotRight2, xyNoseBotRight.X + 0.125, xyNoseBotRight.Y)
            PR_DrawLine(xyNoseBotLeft, xyNoseBotRight2)
            PR_AddEntityID("NoseCoveringBottom")
        End If

    End Sub

    Private Sub PR_DrawNoseCovering()

        ' This Procedure Calculates and Draws the Nose Covering
        ' A Nose covering can only be drawn with the "Face Mask" and "Open Head Mask"

        Dim xyNoseTop, xyNoseTip, xyTempTip, xyNoseTipOut As HEADNECK1.xy
        Dim xyNoseLenStartArc, xyNoseLenCenArc, xyNoseLenEndArc As HEADNECK1.xy
        Dim xyNoseWidthStartArc, xyNoseWidthCenArc, xyNoseWidthEndArc As HEADNECK1.xy
        Dim Opp2, NoseDiff, Adj1, Opp1, Adj2, NoseStepAcross As Double
        Dim X2, Delta, Theta, NoseLen, NoseStepUp, NoseWidth, y2 As Double
        Dim TipX, TipAng, X1, PI, Opp3, Y1, TipHyp, TipY As Double
        Dim TipOpp, XA1, TopY, TopX, TipAng2, YA1, TipAdj As Double
        Dim i As Short
        Dim xyNoseRightTip, HeadPt1, HeadPt2, xyNoseRightTip2 As HEADNECK1.xy
        Dim TipAng3 As Double
        Dim NostralPt1, NostralPt2 As HEADNECK1.xy

        Dim xyTmp As HEADNECK1.xy
        Dim nLength, aStartAngle, aAngle, aTip, aAngleInc, nRad As Double
        Dim ii As Short


        PI = 3.141592654

        If chkNoseCovering.CheckState = 1 Then

            ' Set Values For Nose Length Arc
            Adj1 = Val(txtBotOfEyeX.Text)
            Opp1 = Val(txtBotOfEyeY.Text)
            PR_MakeXY(xyNoseLenCenArc, Adj1 + 0.1875, Opp1 + 0.25)
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            NoseLen = FN_CmToInches(txtUnits, txtLengthOfNose)

            ' Set Values for Nose Width Arc
            Opp1 = Val(txtNoseBottomY.Text)
            Adj1 = Val(txtBotOfEyeX.Text) + 0.375
            PR_MakeXY(xyNoseWidthCenArc, Adj1, Opp1)
            NoseWidth = (FN_CmToInches(txtUnits, txtTipOfNose)) / 2
            'Get intersection point
            If Not FN_CalcCirCirInt(xyNoseLenCenArc, NoseLen, xyNoseWidthCenArc, NoseWidth, xyTempTip, xyTmp) Then
                MsgBox("Can't form nose with the given dimensions for the Nose length and width. The arcs do not intersect.", 48, "Head & Neck")
            End If

            'PR_DrawMarker xyTempTip
            'PR_DrawText "xyTempTip", xyTempTip, .05, 0
            'PR_DrawMarker xyNoseLenCenArc
            'PR_DrawText "xyNoseLenCenArc", xyNoseLenCenArc, .05, 0
            'PR_DrawMarker xyNoseWidthCenArc
            'PR_DrawText "xyNoseWidthCenArc", xyNoseWidthCenArc, .05, 0
            'Print #txtfNum, "hEnt = AddEntity("; HEADNECK1.qq; "circle"; HEADNECK1.qq; ",xyStart.x+"; Str$(xyNoseLenCenArc.x); CC; "xyStart.y+"; Str$(xyNoseLenCenArc.y); ","; NoseLen; ");"
            'Print #txtfNum, "hEnt = AddEntity("; HEADNECK1.qq; "circle"; HEADNECK1.qq; ",xyStart.x+"; Str$(xyNoseWidthCenArc.x); CC; "xyStart.y+"; Str$(xyNoseWidthCenArc.y); ","; NoseWidth; ");"

            ' Jobst Procedure 11.5.4 Extending Nose Length
            TipAng = 270 - (FN_CalcAngle(xyNoseLenCenArc, xyTempTip) * (180 / PI))
            TipAng = TipAng * (PI / 180)
            TipX = (NoseLen + 0.25) * System.Math.Sin(TipAng)
            TipY = (NoseLen + 0.25) * System.Math.Cos(TipAng)
            PR_MakeXY(xyNoseTip, xyNoseLenCenArc.X - TipX, xyNoseLenCenArc.Y - TipY)


            TipAng2 = FN_CalcAngle(xyNoseTip, xyTempTip) * (180 / PI)
            TipAng2 = 90 - TipAng2
            TipAng2 = TipAng2 * (PI / 180)
            TipOpp = 0.1875 * (System.Math.Sin(TipAng2))
            TipAdj = 0.1875 * (System.Math.Cos(TipAng2))

            PR_MakeXY(xyNoseTipOut, xyNoseTip.X - TipAdj, xyNoseTip.Y + TipOpp)
            TipAng = FN_CalcAngle(xyTempTip, xyNoseLenCenArc) ' * (PI / 180)
            nLength = (xyNoseLenCenArc.Y - xyNoseTipOut.Y) - 0.125
            nLength = nLength / System.Math.Sin(TipAng)
            PR_CalcPolar(xyNoseTipOut, (FN_CalcAngle(xyTempTip, xyNoseLenCenArc) * (180 / PI)), nLength, xyNoseTop)
            'PR_DrawMarker xyNoseTop
            'PR_DrawText "xyNoseTop-New", xyNoseTop, .05, 0
            'PR_DrawMarker xyNoseTipOut
            'PR_DrawText "xyNoseTipOut", xyNoseTipOut, .05, 0
            'PR_DrawMarker xyNoseTip
            'PR_DrawText "xyNoseTip", xyNoseTip, .05, 0

            ' Blend Nose into forehead
            TopX = CDbl(txtNoseCoverX.Text)
            TopY = CDbl(txtNoseCoverY.Text)
            PR_MakeXY(HeadPt1, TopX, TopY)
            PR_MakeXY(HeadPt2, TopX - 0.03125, TopY - 0.25)
            'PR_DrawMarker HeadPt1
            'PR_DrawText "HeadPt1", HeadPt1, .05, 0
            'PR_DrawMarker HeadPt2
            'PR_DrawText "HeadPt2", HeadPt2, .05, 0
            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "open" & HEADNECK1.QQ)
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(HeadPt1.X) & HEADNECK1.CC & "xyStart.y+" & Str(HeadPt1.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(HeadPt2.X) & HEADNECK1.CC & "xyStart.y+" & Str(HeadPt2.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyNoseTop.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyNoseTop.Y))
            PrintLine(CInt(txtfNum.Text), ");")
            ''To Draw Polyline
            Dim PolyStart(3), PolyEnd(3) As Double
            PolyStart(1) = HeadPt1.X
            PolyEnd(1) = HeadPt1.Y
            PolyStart(2) = HeadPt2.X
            PolyEnd(2) = HeadPt2.Y
            PolyStart(3) = xyNoseTop.X
            PolyEnd(3) = xyNoseTop.Y
            PR_DrawPoly(PolyStart, PolyEnd, 3)
            PR_AddEntityID("NoseCoveringBridge")

            TipX = (0.25) * System.Math.Cos(TipAng2)
            TipY = (0.25) * System.Math.Sin(TipAng2)
            'TipX = 0.1875 * System.Math.Cos(TipAng2)
            'TipY = 0.1875 * System.Math.Sin(TipAng2)
            PR_MakeXY(xyNoseRightTip, xyNoseTip.X + TipX, xyNoseTip.Y - TipY)

            TipAng3 = (90 - TipAng2 * (180 / PI)) * (PI / 180)
            TipX = 0.25 * System.Math.Cos(TipAng3)
            TipY = 0.25 * System.Math.Sin(TipAng3)
            PR_MakeXY(xyNoseRightTip2, xyNoseRightTip.X + TipX, xyNoseRightTip.Y + TipY)

            ' Draw nose
            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "polyline" & HEADNECK1.QQ)
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyNoseTop.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyNoseTop.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyNoseTipOut.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyNoseTipOut.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyNoseTip.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyNoseTip.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyNoseRightTip.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyNoseRightTip.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(xyNoseRightTip2.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyNoseRightTip2.Y))
            PrintLine(CInt(txtfNum.Text), ");")
            ''To Draw Polyline
            Dim StartPt(5), EndPt(5) As Double
            StartPt(1) = xyNoseTop.X
            EndPt(1) = xyNoseTop.Y
            StartPt(2) = xyNoseTipOut.X
            EndPt(2) = xyNoseTipOut.Y
            StartPt(3) = xyNoseTip.X
            EndPt(3) = xyNoseTip.Y
            StartPt(4) = xyNoseRightTip.X
            EndPt(4) = xyNoseRightTip.Y
            StartPt(5) = xyNoseRightTip2.X
            EndPt(5) = xyNoseRightTip2.Y
            PR_DrawPoly(StartPt, EndPt, 5)

            PR_AddEntityID("NoseCovering")

            ' Calculate and draw nostril
            ' Mods GG 4.Mar.96
            aTip = (FN_CalcAngle(xyNoseLenCenArc, xyNoseTip) * (180 / PI)) + 90
            nLength = FN_CalcLength(xyNoseRightTip2, xyNoseWidthCenArc)
            aAngle = FN_CalcAngle(xyNoseRightTip2, xyNoseWidthCenArc) * (180 / PI)
            If aAngle > 180 Then
                aAngle = aAngle - aTip
            Else
                aAngle = aAngle + (360 - aTip)
            End If
            nRad = (nLength / 2) / System.Math.Cos(aAngle * (PI / 180))
            PR_CalcPolar(xyNoseRightTip2, aTip, nRad, NostralPt1)
            aStartAngle = FN_CalcAngle(NostralPt1, xyNoseRightTip2) * (180 / PI)
            aAngle = FN_CalcAngle(NostralPt1, xyNoseWidthCenArc) * (180 / PI)
            If aAngle > 270 Then aAngle = aAngle - 360 'NB Sign in this case will add
            aAngleInc = (aStartAngle - aAngle) / 3

            'PR_DrawMarker xyNoseRightTip2
            'PR_DrawText "xyNoseRightTip2", xyNoseRightTip2, .05, 0
            'PR_DrawMarker xyNoseTip
            'PR_DrawText "xyNoseTip", xyNoseTip, .05, 0
            'PR_DrawMarker xyNoseWidthCenArc
            'PR_DrawText "xyNoseWidthCenArc", xyNoseWidthCenArc, .05, 0
            'PR_DrawMarker NostralPt1
            'PR_DrawText "NostralPt1", NostralPt1, .05, 0
            'PR_DrawMarker xyNoseRightTip2
            'PR_DrawText "xyNoseRightTip2", xyNoseRightTip2, .05, 0
            'PR_DrawMarker xyNoseLenCenArc
            'PR_DrawText "xyNoseLenCenArc", xyNoseLenCenArc, .05, 0

            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "fitted" & HEADNECK1.QQ)
            'For ii = 0 To 3
            '    PR_CalcPolar(NostralPt1, aStartAngle - (aAngleInc * ii), nRad, NostralPt2)
            '    PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(NostralPt2.X) & HEADNECK1.CC & "xyStart.y+" & Str(NostralPt2.Y))
            'Next ii
            PrintLine(CInt(txtfNum.Text), ");")
            ''To Draw Spline
            Dim ptColl As Point3dCollection = New Point3dCollection()
            For ii = 0 To 3
                PR_CalcPolar(NostralPt1, aStartAngle - (aAngleInc * ii), nRad, NostralPt2)
                ptColl.Add(New Point3d(xyInsertion.X + NostralPt2.X, xyInsertion.Y + NostralPt2.Y, 0))
            Next
            PR_DrawSpline(ptColl)
            PR_AddEntityID("Nostril")

        End If

    End Sub



    Private Sub PR_DrawOpenFace()

        Dim xyOpeningTop, xyFaceStart, xyOpeningBottom As HEADNECK1.xy
        Dim xyLipBottom, xyLipTop As HEADNECK1.xy
        Dim xyPoint1, xyPoint2 As HEADNECK1.xy
        Dim xDist, YDist As Double
        Dim X1, Y1 As Double
        Dim ii As Short
        Dim XStep, YStep As Double

        PR_MakeXY(xyFaceStart, -LeftArcAdjOut(CInt(txtRadiusNo.Text)), 0)
        PR_MakeXY(xyOpeningTop, -LeftArcAdjOut(CInt(txtRadiusNo.Text)), -0.5)
        PR_DrawLine(xyFaceStart, xyOpeningTop)
        PR_AddEntityID("OpenFaceEyeBrow")

        OpenFaceProfileX(1) = xyOpeningTop.X
        OpenFaceProfileY(1) = xyOpeningTop.Y
        XStep = (0.375 + CDbl(txtEyeWidth.Text)) / 2
        OpenFaceProfileX(2) = xyOpeningTop.X + XStep
        OpenFaceProfileY(2) = xyOpeningTop.Y
        XStep = CDbl((txtEyeWidth).Text)

        OpenFaceProfileX(3) = xyOpeningTop.X + XStep
        'OpenFaceProfileY(3) = xyOpeningTop.Y - 0.0625
        OpenFaceProfileY(3) = xyOpeningTop.Y - 0.125
        XStep = CDbl(txtEyeWidth.Text) + 0.25

        YStep = -0.375
        OpenFaceProfileX(4) = xyOpeningTop.X + XStep
        OpenFaceProfileY(4) = xyOpeningTop.Y + YStep

        If chkLipStrap.CheckState = 1 Then
            X1 = -LeftArcAdjOut(CInt(txtRadiusNo.Text))
            Y1 = Val(txtMouthHeight.Text)
            PR_MakeXY(xyLipBottom, X1, Y1)
            PR_MakeXY(xyLipTop, X1, Y1 + Val(txtLipStrapWidth.Text))
            PR_DrawLine(xyLipBottom, xyLipTop)
            PR_AddEntityID("OpenFaceLipStrap")
            PR_MakeXY(xyOpeningBottom, xyLipTop.X, xyLipTop.Y)
        Else
            X1 = CDbl(txtMouthRightX.Text)
            Y1 = Val(txtMouthHeight.Text)
            PR_MakeXY(xyOpeningBottom, X1, Y1)
        End If

        PR_MakeXY(xyPoint1, OpenFaceProfileX(4), OpenFaceProfileY(4))
        PR_MakeXY(xyPoint2, xyOpeningBottom.X, OpenFaceProfileY(4))
        xDist = FN_CalcLength(xyPoint1, xyPoint2)
        XStep = -xDist / 6

        PR_MakeXY(xyPoint1, OpenFaceProfileX(4), OpenFaceProfileY(4))
        PR_MakeXY(xyPoint2, OpenFaceProfileX(4), xyOpeningBottom.Y)
        YDist = FN_CalcLength(xyPoint1, xyPoint2)
        YStep = -YDist / 5

        OpenFaceProfileX(5) = OpenFaceProfileX(4) + 0.0625
        OpenFaceProfileY(5) = OpenFaceProfileY(4) + YStep

        OpenFaceProfileX(6) = OpenFaceProfileX(5) - 0.0625
        OpenFaceProfileY(6) = OpenFaceProfileY(5) + YStep

        OpenFaceProfileX(7) = OpenFaceProfileX(6) + XStep + 0.0625
        OpenFaceProfileY(7) = OpenFaceProfileY(6) + YStep

        OpenFaceProfileX(8) = OpenFaceProfileX(7) + XStep - 0.0625
        OpenFaceProfileY(8) = OpenFaceProfileY(7) + YStep

        OpenFaceProfileX(9) = OpenFaceProfileX(8) + (2 * XStep)
        OpenFaceProfileY(9) = xyOpeningBottom.Y + 0.09375

        OpenFaceProfileX(10) = xyOpeningBottom.X
        OpenFaceProfileY(10) = xyOpeningBottom.Y

        PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "fitted" & HEADNECK1.QQ)
        For ii = 1 To 10
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(OpenFaceProfileX(ii)) & HEADNECK1.CC & "xyStart.y+" & Str(OpenFaceProfileY(ii)))
        Next ii
        Dim PolyStart(10), PolyEnd(10) As Double
        For ii = 1 To 10
            PolyStart(ii) = OpenFaceProfileX(ii)
            PolyEnd(ii) = OpenFaceProfileY(ii)
        Next
        'PR_DrawPoly(PolyStart, PolyEnd, 10)
        ''To Draw Spline
        Dim ptColl As Point3dCollection = New Point3dCollection()
        For ii = 1 To 10
            ptColl.Add(New Point3d(xyInsertion.X + OpenFaceProfileX(ii), xyInsertion.Y + OpenFaceProfileY(ii), 0))
        Next
        PR_DrawSpline(ptColl)
        PrintLine(CInt(txtfNum.Text), ");")
        PR_AddEntityID("OpenFace")

    End Sub

    Private Sub PR_DrawOpenFaceWithEyes()

        Dim xyOpeningTop, xyFaceStart, xyOpeningBottom As HEADNECK1.xy
        Dim xyCen, xyMouthRight As HEADNECK1.xy
        Dim X1, Y1, y2, nEndAng As Double
        Dim nStartAng, nDeltaAng, nRad As Double

        Y1 = CDbl(txtBotOfEyeY.Text)
        y2 = CDbl(txtMouthHeight.Text)
        X1 = CDbl(txtMouthRightX.Text)
        PR_MakeXY(xyFaceStart, -LeftArcAdjOut(CInt(txtRadiusNo.Text)), 0)
        PR_MakeXY(xyOpeningTop, -LeftArcAdjOut(CInt(txtRadiusNo.Text)), Y1)
        PR_MakeXY(xyOpeningBottom, X1, y2 + 0.125)
        PR_MakeXY(xyMouthRight, X1 - 0.125, y2)
        PR_MakeXY(xyCen, xyOpeningBottom.X - 0.125, xyOpeningBottom.Y)

        'Bottom Curve
        nStartAng = 0
        nDeltaAng = -90
        nRad = 0.125
        PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "arc" & HEADNECK1.QC & "xyStart.x +" & Str(xyCen.X) & HEADNECK1.CC & "xyStart.y +" & Str(xyCen.Y) & HEADNECK1.CC & Str(nRad) & HEADNECK1.CC & Str(nEndAng) & HEADNECK1.CC & Str(nDeltaAng) & ");")
        PR_AddEntityID("OpenFaceBottomCurve")
        PR_DrawArc(xyCen, xyMouthRight, xyOpeningBottom)
        PR_DrawLine(xyFaceStart, xyOpeningTop)
        PR_AddEntityID("OpenFaceEyeBrow")

        PR_DrawLine(xyOpeningBottom, xyOpeningTop)
        PR_AddEntityID("OpenFace")

    End Sub

    Private Sub PR_DrawOpenHead()

        Dim Point2, Point1, Point3 As HEADNECK1.xy
        Dim xyClosureText, Cen As HEADNECK1.xy
        Dim Circum, NormCircum As Double
        Dim Total As Object
        Dim ArcAngle2, Elastic, ArcAngle1, PI As Double
        Dim xyTextPoint As HEADNECK1.xy
        Dim Pt1, Pt2 As HEADNECK1.xy
        Dim TotalArcLen, Radius, ArcLen, ArcAngle As Double

        PI = 3.141592654
        PR_MakeXY(Cen, 0, 0)
        Circum = Val(txtRadiusNo.Text)
        NormCircum = Val(txtRadiusNo.Text)

        ' Set For Zipper
        If chkZipper.CheckState = 1 And Val(txtRadiusNo.Text) > 3 Then
            Circum = Circum - 3
        ElseIf chkZipper.CheckState = 1 And Val(txtRadiusNo.Text) <= 3 Then
            Circum = Circum + 17
        End If

        If chkOpenHeadMask.CheckState = 1 Then
            PR_MakeXY(Point1, -LeftArcAdjOut(NormCircum), 0)
            PR_MakeXY(Point2, -LeftArcAdjOut(NormCircum), 0.5)
            PR_MakeXY(Point3, RightArcOpenAdjUp(Circum), 0.5)
            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "polyline" & HEADNECK1.QQ)
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Point1.X) & HEADNECK1.CC & "xyStart.y+" & Str(Point1.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Point2.X) & HEADNECK1.CC & "xyStart.y+" & Str(Point2.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Point3.X) & HEADNECK1.CC & "xyStart.y+" & Str(Point3.Y))
            PrintLine(CInt(txtfNum.Text), ");")
            ''To Draw Polyline
            Dim StartPt(3), EndPt(3) As Double
            StartPt(1) = Point1.X
            EndPt(1) = Point1.Y
            StartPt(2) = Point2.X
            EndPt(2) = Point2.Y
            StartPt(3) = Point3.X
            EndPt(3) = Point3.Y
            PR_DrawPoly(StartPt, EndPt, 3)
            PR_AddEntityID("OpenHeadTop")

            'top arc
            PR_MakeXY(Pt1, -TopArcAdj(CInt(txtRadiusNo.Text)), TopArcOpp(CInt(txtRadiusNo.Text)))
            PR_MakeXY(Pt2, TopArcAdj(CInt(txtRadiusNo.Text)), TopArcOpp(CInt(txtRadiusNo.Text)))
            ArcAngle1 = FN_CalcAngle(Cen, Pt1) * (180 / PI)
            ArcAngle2 = FN_CalcAngle(Cen, Pt2) * (180 / PI)
            ArcAngle = (ArcAngle1 - ArcAngle2) * (PI / 180)
            Radius = TopArcHyp(CInt(txtRadiusNo.Text))
            ArcLen = Radius * ArcAngle
            TotalArcLen = ArcLen

            ' Left Arc
            PR_MakeXY(Pt1, -LeftArcAdjOut(CInt(txtRadiusNo.Text)), LeftArcOppOut(CInt(txtRadiusNo.Text)))
            PR_MakeXY(Pt2, -LeftArcAdjIn(CInt(txtRadiusNo.Text)), LeftArcOppIn(CInt(txtRadiusNo.Text)))
            ArcAngle1 = FN_CalcAngle(Cen, Pt1) * (180 / PI)
            ArcAngle2 = FN_CalcAngle(Cen, Pt2) * (180 / PI)
            ArcAngle = (ArcAngle1 - ArcAngle2) * (PI / 180)
            Radius = LeftArcHyp(CInt(txtRadiusNo.Text))
            ArcLen = Radius * ArcAngle
            TotalArcLen = TotalArcLen + ArcLen


            ' Left Line
            ArcLen = Pt1.Y - Point2.Y
            TotalArcLen = TotalArcLen + ArcLen

            'Top Right arc
            PR_MakeXY(Pt1, TopRightArcAdjIn(CInt(txtRadiusNo.Text)), TopRightArcOppIn(CInt(txtRadiusNo.Text)))
            PR_MakeXY(Pt2, TopRightArcAdjOut(CInt(txtRadiusNo.Text)), TopRightArcOppOut(CInt(txtRadiusNo.Text)))
            ArcAngle1 = FN_CalcAngle(Cen, Pt1) * (180 / PI)
            ArcAngle2 = FN_CalcAngle(Cen, Pt2) * (180 / PI)
            ArcAngle = (ArcAngle1 - ArcAngle2) * (PI / 180)
            Radius = TopArcHyp(CInt(txtRadiusNo.Text))
            ArcLen = Radius * ArcAngle
            TotalArcLen = TotalArcLen + ArcLen


            'Right arc
            PR_MakeXY(Pt1, RightArcAdjUp(CInt(txtRadiusNo.Text)), RightArcOppUp(CInt(txtRadiusNo.Text)))
            ArcAngle1 = FN_CalcAngle(Cen, Pt1) * (180 / PI)
            ArcAngle2 = FN_CalcAngle(Cen, Point3) * (180 / PI)
            ArcAngle = (ArcAngle1 - ArcAngle2) * (PI / 180)

            Radius = System.Math.Sqrt((RightArcOpenAdjUp(Circum) * RightArcOpenAdjUp(Circum)) + 0.25)
            ArcLen = Radius * ArcAngle
            TotalArcLen = TotalArcLen + ArcLen
            'UPGRADE_WARNING: Couldn't resolve default property of object Total. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Total = FN_InchesToText(ARMEDDIA1.fnRoundInches(TotalArcLen))

            PR_MakeXY(xyTextPoint, Point1.X + 2, Point1.Y)
            PR_SetLayer("Notes")
            'UPGRADE_WARNING: Couldn't resolve default property of object Total. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PR_DrawText(Total & Chr(39) & Chr(39) & " Elastic Strap", xyTextPoint, 0.1, 0)
            PR_SetLayer("TemplateLeft")

        End If

    End Sub

    Private Sub PR_DrawRightArc()

        Dim nStartAng, nRad, nEndAng As Double
        Dim nDeltaAng, Circum As Double
        Dim xyEndArc, xyStartArc, xyCentreArc As HEADNECK1.xy
        Dim xyClosureText As HEADNECK1.xy
        Dim OpenAcross As Double
        Dim AcrossUp, UpwardsUp As Double
        Dim AcrossDown, UpwardsDown As Double

        Circum = Val(txtRadiusNo.Text)
        PR_MakeXY(xyCentreArc, 0, 0)

        ' Set For Zipper
        If chkZipper.CheckState = 1 And Val(txtRadiusNo.Text) > 3 Then
            Circum = Circum - 3
            '    PR_MakeXY xyClosureText, RightArcAdjUp(Circum) - .35, .25
            '    PR_SetLayer "Notes"
            '    PR_DrawText "Zipper", xyClosureText, .25, -90
            '    PR_SetLayer "TemplateLeft"
        ElseIf chkZipper.CheckState = 1 And Val(txtRadiusNo.Text) <= 3 Then
            Circum = Circum + 17
            '    PR_MakeXY xyClosureText, RightArcAdjUp(Circum) - .35, .25
            '    PR_SetLayer "Notes"
            '    PR_DrawText "Zipper", xyClosureText, .25, -90
            '    PR_SetLayer "TemplateLeft"
        End If

        ' Set For 2" wide Velcro
        If chkVelcro.CheckState = 1 Then
            Circum = Circum + 2
        End If

        'Draw text for both ZIPPER and VELCRO"
        PR_MakeXY(xyClosureText, RightArcAdjUp(Circum) - 0.35, 0.25)
        PR_SetLayer("Notes")
        Dim nAngle As Double = 90 * (HEADNECK1.PI / 180)
        If chkZipper.CheckState = 1 Then
            PR_DrawText("ZIPPER", xyClosureText, 0.25, -nAngle)
        Else
            PR_DrawText("VELCRO", xyClosureText, 0.25, -nAngle)
        End If
        PR_SetLayer("TemplateLeft")

        AcrossUp = RightArcAdjUp(Circum)
        UpwardsUp = RightArcOppUp(Circum)
        OpenAcross = RightArcOpenAdjUp(Circum)
        AcrossDown = RightArcAdjDown(Circum)
        UpwardsDown = RightArcOppDown(Circum)

        If chkOpenHeadMask.CheckState = 1 Then
            PR_MakeXY(xyEndArc, OpenAcross, 0.5)
        Else
            PR_MakeXY(xyEndArc, AcrossUp, UpwardsUp)
        End If

        PR_MakeXY(xyStartArc, AcrossDown, -UpwardsDown)
        nRad = FN_CalcLength(xyCentreArc, xyStartArc)
        nStartAng = FN_CalcAngle(xyCentreArc, xyStartArc)
        nStartAng = 360 - nStartAng
        nEndAng = FN_CalcAngle(xyCentreArc, xyEndArc)
        nDeltaAng = -(nEndAng + nStartAng)
        PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "arc" & HEADNECK1.QC & "xyStart.x +" & Str(xyCentreArc.X) & HEADNECK1.CC & "xyStart.y +" & Str(xyCentreArc.Y) & HEADNECK1.CC & Str(nRad) & HEADNECK1.CC & Str(nEndAng) & HEADNECK1.CC & Str(nDeltaAng) & ");")
        PR_DrawArc(xyCentreArc, xyStartArc, xyEndArc)
        PR_AddEntityID("RightArc")

    End Sub

    Private Sub PR_DrawRightCutOut()
        Dim AcrossLeft, UpwardsLeft As Double
        Dim AcrossRight, UpwardsRight As Double
        Dim CentreX, CentreY As Double
        Dim LeftTip, Centre, RightTip As HEADNECK1.xy

        If chkOpenHeadMask.CheckState = 0 Then
            CentreX = 3.8 / 2.54
            CentreY = 3.8 / 2.54
            PR_MakeXY(Centre, CentreX, CentreY)

            AcrossLeft = TopArcAdj(CInt(txtRadiusNo.Text))
            UpwardsLeft = TopArcOpp(CInt(txtRadiusNo.Text))
            AcrossRight = TopRightArcAdjIn(CInt(txtRadiusNo.Text))
            UpwardsRight = TopRightArcOppIn(CInt(txtRadiusNo.Text))

            PR_MakeXY(LeftTip, AcrossLeft, UpwardsLeft)
            PR_MakeXY(RightTip, AcrossRight, UpwardsRight)

            PrintLine(CInt(txtfNum.Text), "hEnt = AddEntity(" & HEADNECK1.QQ & "poly" & HEADNECK1.QCQ & "polyline" & HEADNECK1.QQ)
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(LeftTip.X) & HEADNECK1.CC & "xyStart.y+" & Str(LeftTip.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(Centre.X) & HEADNECK1.CC & "xyStart.y+" & Str(Centre.Y))
            PrintLine(CInt(txtfNum.Text), HEADNECK1.CC & "xyStart.x+" & Str(RightTip.X) & HEADNECK1.CC & "xyStart.y+" & Str(RightTip.Y))
            PrintLine(CInt(txtfNum.Text), ");")
            ''To Draw Polyline
            Dim StartPt(3), EndPt(3) As Double
            StartPt(1) = LeftTip.X
            EndPt(1) = LeftTip.Y
            StartPt(2) = Centre.X
            EndPt(2) = Centre.Y
            StartPt(3) = RightTip.X
            EndPt(3) = RightTip.Y
            PR_DrawPoly(StartPt, EndPt, 3)
            PR_AddEntityID("RightCutOut")
        End If

    End Sub

    Private Sub PR_DrawRightJoin()

        Dim AcrossLeft, UpwardsLeft As Double
        Dim AcrossRight, UpwardsRight As Double
        Dim Start, Finish As HEADNECK1.xy
        Dim Circum As Double

        If chkOpenHeadMask.CheckState = 0 Then
            Circum = Val(txtRadiusNo.Text)
            ' Set For Zipper if Circumference is Less than 36
            If chkZipper.CheckState = 1 And Circum > 3 Then
                Circum = Circum - 3
            ElseIf chkZipper.CheckState = 1 And Circum <= 3 Then
                Circum = Circum + 17
            End If
            ' Set For 2" wide Velcro
            If chkVelcro.CheckState = 1 Then
                Circum = Circum + 2
            End If


            AcrossLeft = TopRightArcAdjOut(CInt(txtRadiusNo.Text))
            UpwardsLeft = TopRightArcOppOut(CInt(txtRadiusNo.Text))
            AcrossRight = RightArcAdjUp(Circum)
            UpwardsRight = RightArcOppUp(Circum)

            PR_MakeXY(Start, AcrossLeft, UpwardsLeft)
            PR_MakeXY(Finish, AcrossRight, UpwardsRight)
            PR_DrawLine(Start, Finish)
            PR_AddEntityID("RightArcJoin")
        End If

    End Sub

    Private Sub PR_DrawRightLine()

        Dim xyStart, xyFinish As HEADNECK1.xy
        Dim Across, Down As Double
        Dim Circum, LineLen As Double

        Circum = Val(txtRadiusNo.Text)

        ' Check For Zipper if Circumference is Less than 36
        If chkZipper.CheckState = 1 And Val(txtRadiusNo.Text) > 3 Then
            Circum = Circum - 3
        ElseIf chkZipper.CheckState = 1 And Circum <= 3 Then
            Circum = Circum + 17
        End If
        ' Set For 2" wide Velcro
        If chkVelcro.CheckState = 1 Then
            Circum = Circum + 2
        End If


        LineLen = 1.35 / 2.54
        Across = RightArcAdjDown(Circum)
        Down = -RightArcOppDown(Circum)
        PR_MakeXY(xyStart, Across, Down)
        PR_MakeXY(xyFinish, Across - LineLen, Down)
        PR_DrawLine(xyStart, xyFinish)
        PR_AddEntityID("RightArcNeckJoin")

    End Sub

    Private Sub PR_DrawText(ByRef sText As Object, ByRef xyInsert As HEADNECK1.xy, ByRef nHeight As Object, ByRef nAngle As Object)
        'To the DRAFIX macro file (given by the global txtfNum).
        'Write the syntax to draw TEXT at the given height.
        '
        'For this to work it assumes that the following DRAFIX variables
        'are defined
        '    HEADNECK1.XY      xyStart
        Dim nWidth As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAspect. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nWidth = nHeight * CDbl(txtCurrTextAspect.Text)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                     OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)

            '' Create a single-line text object
            Using acText As DBText = New DBText()
                acText.Position = New Point3d(xyInsert.X + xyInsertion.X, xyInsert.Y + xyInsertion.Y, 0)
                acText.Height = nHeight
                acText.TextString = sText
                acText.Rotation = nAngle
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acText.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If
                acBlkTblRec.AppendEntity(acText)
                acTrans.AddNewlyCreatedDBObject(acText, True)
            End Using

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using




        'UPGRADE_WARNING: Couldn't resolve default property of object nHeight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nWidth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

        PrintLine(CInt(txtfNum.Text), "AddEntity(" & HEADNECK1.QQ & "text" & HEADNECK1.QCQ & sText & HEADNECK1.QC & "xyStart.x+" & Str(xyInsert.X) & HEADNECK1.CC & "xyStart.y+" & Str(xyInsert.Y) & HEADNECK1.CC & nWidth & HEADNECK1.CC & nHeight & HEADNECK1.CC & nAngle & ");")

    End Sub

    Private Sub PR_DrawTopArc()
        ' This Procedure Draws the top Arc of the Face Mask
        ' It uses the value held in txtRadiusNo to determine which
        ' Arc to Draw.  The Value Held in txtRadius must lie between
        ' 1 and 19 inclusive.

        Dim xyEndArc, xyStartArc, xyCentreArc As HEADNECK1.xy
        Dim Across, Upwards As Double

        If chkOpenHeadMask.CheckState = 0 Then
            PR_MakeXY(xyCentreArc, 0, 0)

            Across = TopArcAdj(CInt(txtRadiusNo.Text))
            Upwards = TopArcOpp(CInt(txtRadiusNo.Text))

            PR_MakeXY(xyStartArc, Across, Upwards)
            PR_MakeXY(xyEndArc, -Across, Upwards)
            PR_DrawArc(xyCentreArc, xyStartArc, xyEndArc)
            PR_AddEntityID("TopArc")
        End If
    End Sub

    Private Sub PR_DrawTopRightArc()

        ' This Procedure Draws the top Arc of the Face Mask
        ' It uses the value held in txtRadiusNo to determine which
        ' Arc to Draw.  The Value Held in txtRadius must lie between
        ' 1 and 19 inclusive.

        Dim xyEndArc, xyStartArc, xyCentreArc As HEADNECK1.xy
        Dim AcrossIn, UpwardsIn As Double
        Dim AcrossOut, UpwardsOut As Double
        If chkOpenHeadMask.CheckState = 0 Then
            PR_MakeXY(xyCentreArc, 0, 0)
            AcrossIn = TopRightArcAdjIn(CInt(txtRadiusNo.Text))
            UpwardsIn = TopRightArcOppIn(CInt(txtRadiusNo.Text))

            AcrossOut = TopRightArcAdjOut(CInt(txtRadiusNo.Text))
            UpwardsOut = TopRightArcOppOut(CInt(txtRadiusNo.Text))

            PR_MakeXY(xyStartArc, AcrossIn, UpwardsIn)
            PR_MakeXY(xyEndArc, AcrossOut, UpwardsOut)
            PR_DrawArc(xyCentreArc, xyEndArc, xyStartArc)
            PR_AddEntityID("TopRightArc")
        End If

    End Sub

    Private Sub PR_FaceMask()
        Dim xyTextPoint As HEADNECK1.xy

        If optFaceMask.Checked = True Then
            PR_DrawTopArc()
            PR_DrawRightCutOut()
            PR_DrawTopRightArc()
            PR_DrawRightJoin()
            PR_DrawRightArc()
            PR_DrawRightLine()
            PR_DrawLeftArc()
            PR_DrawLeftCutOut()
            PR_DrawOpenHead()
            PR_SetChart1Details()
            PR_DrawEyeOpening()
            PR_DrawEars()
            PR_DrawNeckAndDart()
            PR_DrawNose()
            PR_DrawForeHead()
            PR_DrawChin()
            PR_DrawNoseCovering()
            PR_DrawEarFlaps()
            PR_DrawEyeFlaps()
            PR_DrawLipCovering()
            PR_StampLining()
            PR_NeckElastic()
            PR_MakeXY(xyTextPoint, 0.3, 0.5)
            PR_TemplateDetails(xyTextPoint, "Face Mask")
            ''To Draw Cross Block
            PR_DrawCrossBlock()

        End If
    End Sub

    Private Sub PR_LoadBotFaceArrays()
        ' This Procedure Loads the Arrays that hold the positions
        ' of the end points for the radials of the Face Line

        Dim BotStart, BotStep As Double
        Dim i As Short

        BotStart = 10.15 / 2.54
        BotStep = (10.1 / 2.54) / 16

        For i = 1 To 19
            BotFaceAdj(i) = -(LeftArcAdjOut(i))
            BotFaceOpp(i) = LeftArcOppOut(i) - BotStart
            BotStart = BotStart + BotStep
        Next i

    End Sub

    Private Sub PR_LoadFaceMaskChart1()

        ' This Procedure Loads Face Mask Chart 1 into an
        ' array Named: FaceMaskChartA

        Dim nMidMouth As Double
        Dim i As Short

        nMidMouth = 2
        For i = 1 To 25
            ' Column 1
            FaceMaskChartA(1, i) = nMidMouth
            nMidMouth = nMidMouth + 0.125
            ' Column 2
            If (i = 1) Or (i = 2) Then
                FaceMaskChartA(2, i) = 0.375
            ElseIf (i >= 3) And (i <= 7) Then
                FaceMaskChartA(2, i) = 0.5
            ElseIf (i >= 8) And (i <= 12) Then
                FaceMaskChartA(2, i) = 0.625
            ElseIf (i >= 13) And (i <= 17) Then
                FaceMaskChartA(2, i) = 0.75
            ElseIf (i >= 18) And (i <= 22) Then
                FaceMaskChartA(2, i) = 0.875
            ElseIf (i >= 23) And (i <= 25) Then
                FaceMaskChartA(2, i) = 1
            End If
            ' Column 3
            If (i = 1) Or (i = 2) Then
                FaceMaskChartA(3, i) = 0.5
            ElseIf (i >= 3) And (i <= 6) Then
                FaceMaskChartA(3, i) = 0.625
            ElseIf (i >= 7) And (i <= 10) Then
                FaceMaskChartA(3, i) = 0.75
            ElseIf (i >= 11) And (i <= 14) Then
                FaceMaskChartA(3, i) = 0.825
            ElseIf (i >= 15) And (i <= 18) Then
                FaceMaskChartA(3, i) = 1
            ElseIf (i >= 19) And (i <= 22) Then
                FaceMaskChartA(3, i) = 1.125
            ElseIf (i >= 23) And (i <= 25) Then
                FaceMaskChartA(3, i) = 1.25
            End If
        Next i

    End Sub

    Private Sub PR_LoadFaceMaskChart2()

        ' This Procedure Loads Face Mask Chart 2 into an
        ' array Named: FaceMaskChart2

        Dim nCircum, i As Short

        nCircum = 30
        For i = 1 To 16
            ' Column 1
            FaceMaskChart2(1, i) = nCircum
            nCircum = nCircum + 2
            ' Column 2
            If (i = 1) Then
                FaceMaskChart2(2, i) = 0.75
            ElseIf (i >= 2) And (i <= 4) Then
                FaceMaskChart2(2, i) = 0.875
            ElseIf (i >= 5) And (i <= 7) Then
                FaceMaskChart2(2, i) = 1
            ElseIf (i >= 8) And (i <= 9) Then
                FaceMaskChart2(2, i) = 1.125
            ElseIf (i >= 10) And (i <= 12) Then
                FaceMaskChart2(2, i) = 1.25
            ElseIf (i >= 13) And (i <= 14) Then
                FaceMaskChart2(2, i) = 1.375
            ElseIf (i >= 15) And (i <= 16) Then
                FaceMaskChart2(2, i) = 1.5
            End If
            ' Column 3
            If (i = 1) Then
                FaceMaskChart2(3, i) = 1.625
            ElseIf (i = 2) Or (i = 3) Then
                FaceMaskChart2(3, i) = 1.75
            ElseIf (i = 4) Or (i = 5) Then
                FaceMaskChart2(3, i) = 1.875
            ElseIf (i = 6) Then
                FaceMaskChart2(3, i) = 2
            ElseIf (i = 7) Or (i = 8) Then
                FaceMaskChart2(3, i) = 2.125
            ElseIf (i = 9) Or (i = 10) Then
                FaceMaskChart2(3, i) = 2.25
            ElseIf (i = 11) Then
                FaceMaskChart2(3, i) = 2.375
            ElseIf (i = 12) Or (i = 13) Then
                FaceMaskChart2(3, i) = 2.5
            ElseIf (i = 14) Then
                FaceMaskChart2(3, i) = 2.625
            ElseIf (i = 15) Or (i = 16) Then
                FaceMaskChart2(3, i) = 2.75
            End If
        Next i

    End Sub

    Private Sub PR_LoadFaceMaskChart3()
        ' This Procedure Loads Face Mask Chart 3 into an
        ' array Named: FaceMaskChart3
        Dim nNeck As Single
        Dim i As Short

        nNeck = 4

        ' Column 3
        For i = 1 To 47
            FaceMaskChart3(3, i) = nNeck
            nNeck = nNeck + 0.125
        Next i

        ' Column 1
        FaceMaskChart3(1, 1) = 8
        FaceMaskChart3(1, 2) = 8.25
        FaceMaskChart3(1, 3) = 8.5
        FaceMaskChart3(1, 4) = 8.75
        FaceMaskChart3(1, 5) = 9
        FaceMaskChart3(1, 6) = 9.25
        FaceMaskChart3(1, 7) = 9.625
        FaceMaskChart3(1, 8) = 9.875
        FaceMaskChart3(1, 9) = 10.125
        FaceMaskChart3(1, 10) = 10.375
        FaceMaskChart3(1, 11) = 10.625
        FaceMaskChart3(1, 12) = 10.875
        FaceMaskChart3(1, 13) = 11.125
        FaceMaskChart3(1, 14) = 11.375
        FaceMaskChart3(1, 15) = 11.625
        FaceMaskChart3(1, 16) = 11.875
        FaceMaskChart3(1, 17) = 12.125
        FaceMaskChart3(1, 18) = 12.375
        FaceMaskChart3(1, 19) = 12.75
        FaceMaskChart3(1, 20) = 13
        FaceMaskChart3(1, 21) = 13.25
        FaceMaskChart3(1, 22) = 13.5
        FaceMaskChart3(1, 23) = 13.75
        FaceMaskChart3(1, 24) = 14
        FaceMaskChart3(1, 25) = 14.25
        FaceMaskChart3(1, 26) = 14.5
        FaceMaskChart3(1, 27) = 14.75
        FaceMaskChart3(1, 28) = 15
        FaceMaskChart3(1, 29) = 15.25
        FaceMaskChart3(1, 30) = 15.5
        FaceMaskChart3(1, 31) = 15.875
        FaceMaskChart3(1, 32) = 16.125
        FaceMaskChart3(1, 33) = 16.375
        FaceMaskChart3(1, 34) = 16.625
        FaceMaskChart3(1, 35) = 16.875
        FaceMaskChart3(1, 36) = 17.125
        FaceMaskChart3(1, 37) = 17.375
        FaceMaskChart3(1, 38) = 17.625
        FaceMaskChart3(1, 39) = 17.875
        FaceMaskChart3(1, 40) = 18.125
        FaceMaskChart3(1, 41) = 18.375
        FaceMaskChart3(1, 42) = 18.625
        FaceMaskChart3(1, 43) = 19
        FaceMaskChart3(1, 44) = 19.25
        FaceMaskChart3(1, 45) = 19.5
        FaceMaskChart3(1, 46) = 19.75
        FaceMaskChart3(1, 47) = 20

        ' Column 2
        FaceMaskChart3(2, 1) = 8.125
        FaceMaskChart3(2, 2) = 8.75
        FaceMaskChart3(2, 3) = 8.625
        FaceMaskChart3(2, 4) = 8.875
        FaceMaskChart3(2, 5) = 9.125
        FaceMaskChart3(2, 6) = 9.5
        FaceMaskChart3(2, 7) = 9.75
        FaceMaskChart3(2, 8) = 10
        FaceMaskChart3(2, 9) = 10.25
        FaceMaskChart3(2, 10) = 10.5
        FaceMaskChart3(2, 11) = 10.75
        FaceMaskChart3(2, 12) = 11
        FaceMaskChart3(2, 13) = 11.25
        FaceMaskChart3(2, 14) = 11.5
        FaceMaskChart3(2, 15) = 11.75
        FaceMaskChart3(2, 16) = 12
        'FaceMaskChart3(2, 17) = 12.125   'GG Bugfix
        FaceMaskChart3(2, 17) = 12.25 'GG Bugfix
        FaceMaskChart3(2, 18) = 12.625
        FaceMaskChart3(2, 19) = 12.875
        FaceMaskChart3(2, 20) = 13.125
        FaceMaskChart3(2, 21) = 13.375
        FaceMaskChart3(2, 22) = 13.625
        FaceMaskChart3(2, 23) = 13.875
        FaceMaskChart3(2, 24) = 14.125
        FaceMaskChart3(2, 25) = 14.375
        FaceMaskChart3(2, 26) = 14.625
        FaceMaskChart3(2, 27) = 14.875
        FaceMaskChart3(2, 28) = 15.125
        FaceMaskChart3(2, 29) = 15.375
        FaceMaskChart3(2, 30) = 15.75
        FaceMaskChart3(2, 31) = 16
        FaceMaskChart3(2, 32) = 16.25
        FaceMaskChart3(2, 33) = 16.5
        FaceMaskChart3(2, 34) = 16.75
        FaceMaskChart3(2, 35) = 17
        FaceMaskChart3(2, 36) = 17.25
        FaceMaskChart3(2, 37) = 17.5
        FaceMaskChart3(2, 38) = 17.75
        FaceMaskChart3(2, 39) = 18
        FaceMaskChart3(2, 40) = 18.25
        FaceMaskChart3(2, 41) = 18.5
        FaceMaskChart3(2, 42) = 18.875
        FaceMaskChart3(2, 43) = 19.125
        FaceMaskChart3(2, 44) = 19.375
        FaceMaskChart3(2, 45) = 19.625
        FaceMaskChart3(2, 46) = 19.875
        FaceMaskChart3(2, 47) = 0

    End Sub

    Private Sub PR_LoadLeftArcArrays()
        ' This Procedure Loads the Arrays that hold the positions
        ' of the start and end points for the radials of the Left Arc

        Dim TopArcUpwards, LeftArcUpwardsIn As Double
        Dim LeftArcAcrossIn, LeftUpwardsIn As Double
        Dim i As Short
        Dim LeftArcAcrossOut, LeftArcUpwardsOut As Double
        Dim LeftAdjOut, LeftAdjIn, LeftHyp As Double

        TopArcUpwards = 0.50625 / 2.54
        LeftArcAcrossIn = 0.456275 / 2.54
        LeftArcAcrossOut = 0.4796875 / 2.54
        LeftArcUpwardsIn = 0.2375 / 2.54
        LeftArcUpwardsOut = 0.146875 / 2.54

        ' Left Arc Values
        LeftAdjIn = 6.25 / 2.54
        LeftAdjOut = 7.65 / 2.54
        LeftHyp = 8.03 / 2.54

        ' Load Arrays
        For i = 1 To 19

            LeftArcAdjIn(i) = LeftAdjIn
            LeftArcAdjOut(i) = LeftAdjOut
            LeftArcHyp(i) = LeftHyp
            LeftArcOppIn(i) = System.Math.Sqrt((LeftHyp * LeftHyp) - (LeftAdjIn * LeftAdjIn))
            LeftArcOppOut(i) = System.Math.Sqrt((LeftHyp * LeftHyp) - (LeftAdjOut * LeftAdjOut))
            LeftAdjIn = LeftAdjIn + LeftArcAcrossIn
            LeftAdjOut = LeftAdjOut + LeftArcAcrossOut
            LeftHyp = LeftHyp + TopArcUpwards

        Next i

    End Sub

    Private Sub PR_LoadNeckArrays()

        ' This Procedure Loads the array NeckRight(index)
        ' with the coordinates of the template for the right
        ' side of the neck

        Dim OppStart, AdjStart As Double
        Dim OppEnd, AdjEnd As Double
        Dim i As Short
        Dim OppStep, AdjStep As Double

        OppStart = -7.65 / 2.54
        OppEnd = -15.25 / 2.54
        AdjStart = 6.35 / 2.54
        AdjEnd = 14 / 2.54
        OppStep = (OppEnd - OppStart) / 16
        AdjStep = (AdjEnd - AdjStart) / 16

        For i = 1 To 19
            NeckRightOpp(i) = OppStart
            NeckRightAdj(i) = AdjStart
            OppStart = OppStart + OppStep
            AdjStart = AdjStart + AdjStep
        Next i

    End Sub

    Private Sub PR_LoadRightArcArrays()

        ' This Procedure Loads the Arrays that hold the positions
        ' of the start and end points for the radials of the Right Arc

        Dim TopAdj, TopArcUpwards, TopHyp As Double
        Dim LeftAdjOut, LeftAdjIn, LeftHyp As Double
        Dim RightArcAcrossTop, RightArcUpwards As Double
        Dim RightArcAcrossBot, RightArcDownwards As Double
        Dim RightAdjDown, RightAdjUp, RightHyp As Double
        Dim RightAdjDown2, RightAdjUp2, RightHyp2 As Double
        Dim OpenAcross As Double
        Dim i, j As Short

        RightArcAcrossTop = 0.45 / 2.54
        RightArcAcrossBot = 0.4625 / 2.54
        RightArcUpwards = 0.23125 / 2.54
        RightArcDownwards = 0.23125 / 2.54
        TopArcUpwards = 0.50625 / 2.54

        ' Right Arc Values
        RightAdjUp = 7.85 / 2.54
        RightAdjDown = 6.4 / 2.54
        RightHyp = 9.03 / 2.54
        OpenAcross = 8.85 / 2.54

        RightAdjUp2 = (7.85 / 2.54) - 3 * (RightArcAcrossTop)
        RightAdjDown2 = (6.4 / 2.54) - 3 * (RightArcAcrossBot)
        RightHyp2 = (9.03 / 2.54) - 3 * (TopArcUpwards)

        ' Load Right Arc Arrays
        For i = 1 To 19
            RightArcAdjUp(i) = RightAdjUp
            RightArcAdjDown(i) = RightAdjDown
            RightArcOpenAdjUp(i) = System.Math.Sqrt((RightHyp * RightHyp) - (0.25))
            RightArcOppUp(i) = System.Math.Sqrt((RightHyp * RightHyp) - (RightAdjUp * RightAdjUp))
            RightArcOppDown(i) = System.Math.Sqrt((RightHyp * RightHyp) - (RightAdjDown * RightAdjDown))
            RightAdjUp = RightAdjUp + RightArcAcrossTop
            RightAdjDown = RightAdjDown + RightArcAcrossBot
            RightHyp = RightHyp + TopArcUpwards
        Next i

        For i = 18 To 20
            j = 20
            RightArcAdjUp(i) = RightAdjUp2
            RightArcAdjDown(i) = RightAdjDown2
            RightHyp2 = RightHyp2
            RightArcOpenAdjUp(i) = System.Math.Sqrt((RightHyp2 * RightHyp2) - (0.25))
            RightArcOppUp(i) = System.Math.Sqrt((RightHyp2 * RightHyp2) - (RightAdjUp2 * RightAdjUp2))
            RightArcOppDown(i) = System.Math.Sqrt((RightHyp2 * RightHyp2) - (RightAdjDown2 * RightAdjDown2))
            RightAdjUp2 = RightAdjUp2 + RightArcAcrossTop
            RightAdjDown2 = RightAdjDown2 + RightArcAcrossBot
            RightHyp2 = RightHyp2 + TopArcUpwards
            j = j - 1
        Next i

    End Sub

    Private Sub PR_LoadTopArcArrays()

        ' This Procedure Loads the Arrays that hold the positions
        ' of the start and end points for the radials of the Top Arc

        Dim TopArcAcross, TopArcUpwards As Double
        Dim TopAdj, TopHyp As Double
        Dim i As Short

        TopArcAcross = 0.23125 / 2.54
        TopArcUpwards = 0.50625 / 2.54

        ' Top Arc Values
        TopAdj = 5.05 / 2.54
        TopHyp = 8 / 2.54

        ' Load Top Arc Arrays
        For i = 1 To 19
            TopArcAdj(i) = TopAdj
            TopArcHyp(i) = TopHyp
            TopArcOpp(i) = System.Math.Sqrt((TopHyp * TopHyp) - (TopAdj * TopAdj))
            TopAdj = TopAdj + TopArcAcross
            TopHyp = TopHyp + TopArcUpwards
        Next i

    End Sub

    Private Sub PR_LoadTopRightArrays()

        ' This Procedure Loads the Arrays that hold the positions
        ' of the start and end points for the radials of the Top Right Arc

        Dim TopRightAdjIn, TopRightAdjOut As Double
        Dim TopRightHyp, TopArcUpwards As Double
        Dim TopRightArcAcrossIn, TopRightArcUpwardsIn As Double
        Dim TopRightArcAcrossOut, TopRightArcUpwardsOut As Double
        Dim i As Short

        TopRightArcAcrossIn = 0.453125 / 2.54
        TopRightArcAcrossOut = 0.4515625 / 2.54
        TopRightArcUpwardsIn = 0.23125 / 2.54
        TopRightArcUpwardsOut = 0.23125 / 2.54
        TopArcUpwards = 0.50625 / 2.54

        ' Top Right Arc Values
        TopRightAdjIn = 6.25 / 2.54
        TopRightAdjOut = 6.955 / 2.54
        TopRightHyp = 8.03 / 2.54

        ' Load Arrays
        For i = 1 To 19
            ' Top Right Arc Arrays
            TopRightArcAdjIn(i) = TopRightAdjIn
            TopRightArcAdjOut(i) = TopRightAdjOut
            TopRightArcOppIn(i) = System.Math.Sqrt((TopRightHyp * TopRightHyp) - (TopRightAdjIn * TopRightAdjIn))
            TopRightArcOppOut(i) = System.Math.Sqrt((TopRightHyp * TopRightHyp) - (TopRightAdjOut * TopRightAdjOut))
            TopRightAdjIn = TopRightAdjIn + TopRightArcAcrossIn
            TopRightAdjOut = TopRightAdjOut + TopRightArcAcrossOut
            TopRightHyp = TopRightHyp + TopArcUpwards
        Next i

    End Sub

    Private Sub PR_MakeXY(ByRef xyReturn As HEADNECK1.xy, ByRef X As Double, ByRef Y As Double)

        'Utility to return a point based on the X and Y values given

        xyReturn.X = X
        xyReturn.Y = Y

    End Sub

    Private Sub PR_NeckElastic()

        Dim xyNeckElastic As HEADNECK1.xy

        ' Stamp For Lining
        If chkNeckElastic.CheckState = 1 Then
            PR_MakeXY(xyNeckElastic, CDbl(txtChinLeftBotX.Text) + 0.4, CDbl(txtChinLeftBotY.Text) + 0.5)
            PR_SetLayer("Notes")
            PR_DrawText("1" & Chr(34) & " Elastic Sewn Underneath", xyNeckElastic, 0.1, 0)
            PR_SetLayer("TemplateLeft")
        End If

    End Sub

    Private Sub PR_OpenFaceMask()

        Dim xyTextPoint As HEADNECK1.xy

        If optOpenFaceMask.Checked = True Then
            PR_DrawTopArc()
            PR_DrawRightCutOut()
            PR_DrawTopRightArc()
            PR_DrawRightJoin()
            PR_DrawRightArc()
            PR_DrawRightLine()
            PR_DrawLeftArc()
            PR_DrawLeftCutOut()
            PR_SetChart1Details()
            PR_DrawEyeOpening()
            PR_DrawEars()
            PR_DrawNeckAndDart()
            PR_DrawForeHead()
            PR_DrawChin()
            PR_DrawOpenHead()
            PR_DrawEarFlaps()
            PR_DrawEyeFlaps()
            PR_StampLining()
            PR_NeckElastic()
            If chkEyes.CheckState = 0 Then
                PR_DrawOpenFace()
            ElseIf chkEyes.CheckState = 1 Then
                PR_DrawOpenFaceWithEyes()
            End If
            PR_DrawLipCovering()
            PR_MakeXY(xyTextPoint, 0.3, 0.5)
            PR_TemplateDetails(xyTextPoint, "Open Face Mask")
            ''To Draw Cross Block
            PR_DrawCrossBlock()
        End If

    End Sub

    Private Sub PR_PutLine(ByRef sLine As String)
        'Puts the contents of sLine to the opened "Macro" file
        'Puts the line with no translation or additions
        '    fNum is global variable
        '
        PrintLine(CInt(txtfNum.Text), "//" & sLine)

    End Sub

    Private Sub PR_SaveDetails()

        Dim sSymbol As String

        sSymbol = "HEADNECK"

        If txtUidHN.Text = "" Then
            'Find "mainpatientdetails" and get position
            PrintLine(CInt(txtfNum.Text), "HEADNECK1.XY     xyMPD_Origin, xyMPD_Scale ;")
            PrintLine(CInt(txtfNum.Text), "STRING sMPD_Name;")
            PrintLine(CInt(txtfNum.Text), "ANGLE  aMPD_Angle;")

            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CInt(txtfNum.Text), "hMPD = UID (" & HEADNECK1.QQ & "find" & HEADNECK1.QC & Val(txtUidMPD.Text) & ");")
            PrintLine(CInt(txtfNum.Text), "if (hMPD)")
            PrintLine(CInt(txtfNum.Text), "  GetGeometry(hMPD, &sMPD_Name, &xyMPD_Origin, &xyMPD_Scale, &aMPD_Angle);")
            PrintLine(CInt(txtfNum.Text), "else")
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CInt(txtfNum.Text), "  Exit(%cancel," & HEADNECK1.QQ & "Can't find > mainpatientdetails < symbol, Insert Patient Data" & HEADNECK1.QQ & ");")

            'Insert headneck
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CInt(txtfNum.Text), "if ( Symbol(" & HEADNECK1.QQ & "find" & HEADNECK1.QCQ & sSymbol & HEADNECK1.QQ & ")){")
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CInt(txtfNum.Text), "  Execute (" & HEADNECK1.QQ & "menu" & HEADNECK1.QCQ & "SetLayer" & HEADNECK1.QC & "Table(" & HEADNECK1.QQ & "find" & HEADNECK1.QCQ & "layer" & HEADNECK1.QCQ & "Data" & HEADNECK1.QQ & "));")
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.QCQ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CInt(txtfNum.Text), "  hFacemask = AddEntity(" & HEADNECK1.QQ & "symbol" & HEADNECK1.QCQ & sSymbol & HEADNECK1.QC & "xyMPD_Origin);")
            PrintLine(CInt(txtfNum.Text), "  }")
            PrintLine(CInt(txtfNum.Text), "else")
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CInt(txtfNum.Text), "  Exit(%cancel, " & HEADNECK1.QQ & "Can't find >" & sSymbol & "< symbol to insert\nCheck your installation, that JOBST.SLB exists!" & HEADNECK1.QQ & ");")
        Else
            'Use existing symbol
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.QC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CInt(txtfNum.Text), "hFacemask = UID (" & HEADNECK1.QQ & "find" & HEADNECK1.QC & Val(txtUidHN.Text) & ");")
            'UPGRADE_WARNING: Couldn't resolve default property of object HEADNECK1.qq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(CInt(txtfNum.Text), "if (!hFacemask) Exit(%cancel," & HEADNECK1.QQ & "Can't find >" & sSymbol & "< symbol to update!" & HEADNECK1.QQ & ");")

        End If


        ' Update data base fields
        PrintLine(CInt(txtfNum.Text), "SetDBData(hFacemask," & HEADNECK1.QQ & "HeadNeck" & HEADNECK1.QCQ & txtMeasurements.Text & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "SetDBData(hFacemask," & HEADNECK1.QQ & "Fabric" & HEADNECK1.QCQ & cboFabric.Text & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "SetDBData(hFacemask," & HEADNECK1.QQ & "WorkOrder" & HEADNECK1.QCQ & txtWorkOrder.Text & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "SetDBData(hFacemask," & HEADNECK1.QQ & "Data" & HEADNECK1.QCQ & txtData.Text & HEADNECK1.QQ & ");")

    End Sub

    Private Sub PR_SelectTemplateRadius()

        ' This Procedure determines the radius
        ' to use on the template

        Dim nTotal, nDecVals As Double
        Dim nEyeBrow, nChinAngle As Double
        Dim nChinRem, nEyeRem, nRemTotal As Double
        Dim nChinInt, nEyeInt, nIntTotal As Short

        'Check Values have been entered
        If Val(txtCircEyeBrow.Text) <= 0 Then
            MsgBox("Invalid or No value has been entered for CIRC above Eyebrow.", 48, "Head & Neck")
            txtCircEyeBrow.Focus()
            Exit Sub
        End If
        If Val(txtCircChinAngle.Text) <= 0 And optHeadBand.Checked = False Then
            MsgBox("Invalid or No value has been entered for CIRC around head at Chin Angle.", 48, "Head & Neck")
            txtCircChinAngle.Focus()
            Exit Sub
        End If

        ' Convert to Inches
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nEyeBrow = FN_CmToInches(txtUnits, txtCircEyeBrow)
        If optHeadBand.Checked = True Then
            nChinAngle = nEyeBrow
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nChinAngle = FN_CmToInches(txtUnits, txtCircChinAngle)
        End If

        ' Add Values and Round up
        nEyeInt = Int(nEyeBrow)
        nEyeRem = nEyeBrow - nEyeInt
        nChinInt = Int(nChinInt)
        nChinRem = nChinAngle - nChinInt
        nIntTotal = nEyeInt + nChinInt
        nRemTotal = nEyeRem + nChinRem
        nDecVals = nRemTotal / 0.125
        If nDecVals >= 8 Then
            nIntTotal = nIntTotal + 1
            nDecVals = nDecVals - 8
        End If
        nDecVals = nDecVals * 0.125
        nTotal = nIntTotal + nDecVals

        ' Determine Radius No.
        If (nTotal <= 30) Then
            txtRadiusNo.Text = CStr(1)
            txtCircumferenceTotal.Text = CStr(30)
        ElseIf (nTotal > 30) And (nTotal <= 32) Then
            txtRadiusNo.Text = CStr(2)
            txtCircumferenceTotal.Text = CStr(32)
        ElseIf (nTotal > 32) And (nTotal <= 34) Then
            txtRadiusNo.Text = CStr(3)
            txtCircumferenceTotal.Text = CStr(34)
        ElseIf (nTotal > 34) And (nTotal <= 36) Then
            txtRadiusNo.Text = CStr(4)
            txtCircumferenceTotal.Text = CStr(36)
        ElseIf (nTotal > 36) And (nTotal <= 38) Then
            txtRadiusNo.Text = CStr(5)
            txtCircumferenceTotal.Text = CStr(38)
        ElseIf (nTotal > 38) And (nTotal <= 40) Then
            txtRadiusNo.Text = CStr(6)
            txtCircumferenceTotal.Text = CStr(40)
        ElseIf (nTotal > 40) And (nTotal <= 42) Then
            txtRadiusNo.Text = CStr(7)
            txtCircumferenceTotal.Text = CStr(42)
        ElseIf (nTotal > 42) And (nTotal <= 44) Then
            txtRadiusNo.Text = CStr(8)
            txtCircumferenceTotal.Text = CStr(44)
        ElseIf (nTotal > 44) And (nTotal <= 46) Then
            txtRadiusNo.Text = CStr(9)
            txtCircumferenceTotal.Text = CStr(46)
        ElseIf (nTotal > 46) And (nTotal <= 48) Then
            txtRadiusNo.Text = CStr(10)
            txtCircumferenceTotal.Text = CStr(48)
        ElseIf (nTotal > 48) And (nTotal <= 50) Then
            txtRadiusNo.Text = CStr(11)
            txtCircumferenceTotal.Text = CStr(50)
        ElseIf (nTotal > 50) And (nTotal <= 52) Then
            txtRadiusNo.Text = CStr(12)
            txtCircumferenceTotal.Text = CStr(52)
        ElseIf (nTotal > 52) And (nTotal <= 54) Then
            txtRadiusNo.Text = CStr(13)
            txtCircumferenceTotal.Text = CStr(54)
        ElseIf (nTotal > 54) And (nTotal <= 56) Then
            txtRadiusNo.Text = CStr(14)
            txtCircumferenceTotal.Text = CStr(56)
        ElseIf (nTotal > 56) And (nTotal <= 58) Then
            txtRadiusNo.Text = CStr(15)
            txtCircumferenceTotal.Text = CStr(58)
        ElseIf (nTotal > 58) And (nTotal <= 60) Then
            txtRadiusNo.Text = CStr(16)
            txtCircumferenceTotal.Text = CStr(60)
        ElseIf (nTotal > 60) And (nTotal <= 62) Then
            txtRadiusNo.Text = CStr(17)
            txtCircumferenceTotal.Text = CStr(62)
        End If

    End Sub

    Private Sub PR_SelectTextInBox(ByRef Text_Box_Name As System.Windows.Forms.Control)

        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Text_Box_Name.SelStart = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Text_Box_Name.SelLength = 4

    End Sub

    Private Sub PR_SeperateMeasurements()

        Dim tempval As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = VB.Left(txtMeasurements.Text, 2) & "." & Mid(txtMeasurements.Text, 3, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtChinToMouth.Text = CStr(Val(tempval))
            PR_ShowInches(tempval, lblInch1)
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtMeasurements.Text, 4, 2) & "." & Mid(txtMeasurements.Text, 6, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtCircEyeBrow.Text = CStr(Val(tempval))
            PR_ShowInches(tempval, lblInch2)
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtMeasurements.Text, 7, 2) & "." & Mid(txtMeasurements.Text, 9, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtCircChinAngle.Text = CStr(Val(tempval))
            PR_ShowInches(tempval, lblInch3)
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtMeasurements.Text, 10, 2) & "." & Mid(txtMeasurements.Text, 12, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtCircOfNeck.Text = CStr(Val(tempval))
            PR_ShowInches(tempval, lblInch4)
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtMeasurements.Text, 13, 2) & "." & Mid(txtMeasurements.Text, 15, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtThroatToSternal.Text = CStr(Val(tempval))
            PR_ShowInches(tempval, lblInch5)
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtMeasurements.Text, 16, 2) & "." & Mid(txtMeasurements.Text, 18, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtTipOfNose.Text = CStr(Val(tempval))
            PR_ShowInches(tempval, lblInch6)
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtMeasurements.Text, 19, 2) & "." & Mid(txtMeasurements.Text, 21, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtLengthOfNose.Text = CStr(Val(tempval))
            PR_ShowInches(tempval, lblInch7)
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtMeasurements.Text, 22, 2) & "." & Mid(txtMeasurements.Text, 24, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtHeadBandDepth.Text = CStr(Val(tempval))
            PR_ShowInches(tempval, lblInch8)
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtMeasurements.Text, 25, 2) & "." & Mid(txtMeasurements.Text, 27, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtLeftEarLength.Text = CStr(Val(tempval))
            PR_ShowInches(tempval, lblInch9)
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtMeasurements.Text, 28, 2) & "." & Mid(txtMeasurements.Text, 30, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtRightEarLength.Text = CStr(Val(tempval))
            PR_ShowInches(tempval, lblInch10)
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtMeasurements.Text, 31, 2) & "." & Mid(txtMeasurements.Text, 33, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtChinCollarMin.Text = CStr(Val(tempval))
            PR_ShowInches(tempval, lblInch11)
        End If

    End Sub

    Private Sub PR_SeperateModifications()

        Dim tempval As Object

        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = VB.Left(txtData.Text, 3)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If tempval = "RFM" Then
            optFaceMask.Checked = True
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf tempval = "RHB" Then
            optHeadBand.Checked = True
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf tempval = "OFM" Then
            optOpenFaceMask.Checked = True
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf tempval = "RCS" Then
            optChinStrap.Checked = True
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf tempval = "MCS" Then
            optModifiedChinStrap.Checked = True
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf tempval = "RCC" Then
            optChinCollar.Checked = True
            'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf tempval = "CCC" Then
            optContouredChinCollar.Checked = True
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 4, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkLeftEarFlap.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 5, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkRightEarFlap.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 6, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkNeckElastic.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 7, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkLeftEyeFlap.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 8, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkRightEyeFlap.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 9, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkOpenHeadMask.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 10, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkLipStrap.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 11, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkLipCovering.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 12, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkNoseCovering.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 13, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkZipper.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 14, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkLining.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 15, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkEarSize.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 16, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkRightEarClosed.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 17, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkLeftEarClosed.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 18, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkEyes.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempval = Mid(txtData.Text, 19, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object tempval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(tempval) = 1 Then
            chkVelcro.CheckState = System.Windows.Forms.CheckState.Checked
        End If

    End Sub

    Private Sub PR_SetChart1Details()

        ' This Procedure Follows the Jobst Procedures No. 4
        ' It Gets the Mouth Height from Measurements,
        ' and checks chart 1 to get width of Lip strap and
        ' eye Height.
        ' These are then marked on the drawing

        Dim xyMouthStart, xyMouthEnd As HEADNECK1.xy
        Dim xyEyeStart, xyEyeEnd As HEADNECK1.xy
        Dim PointOne, PointTwo As Double
        Dim Chart1Ref, MouthHt As Double
        Dim LipWidth, EyeHt As Double
        Dim nArrayVal, nChart1Ref As Double
        Dim nArrayValB As Double
        Dim i As Short

        ' Get start point for Mouth Height
        PointOne = BotFaceAdj(CInt(txtRadiusNo.Text))
        PointTwo = BotFaceOpp(CInt(txtRadiusNo.Text))
        'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        MouthHt = FN_CmToInches(txtUnits, txtChinToMouth)
        PR_MakeXY(xyMouthStart, PointOne, PointTwo + MouthHt)

        ' set txtMouthHeight
        txtMouthHeight.Text = CStr(xyMouthStart.Y)

        ' Check Chart1
        nChart1Ref = System.Math.Abs(PointTwo) - System.Math.Abs(MouthHt)
        If nChart1Ref = (FaceMaskChartA(1, 25)) Then
            LipWidth = (FaceMaskChartA(2, 25))
            EyeHt = (FaceMaskChartA(3, 25))
        Else
            For i = 1 To 24
                nArrayVal = (FaceMaskChartA(1, i))
                nArrayValB = (FaceMaskChartA(1, i + 1))
                If (nChart1Ref >= nArrayVal) And (nChart1Ref < nArrayValB) Then
                    LipWidth = (FaceMaskChartA(2, i))
                    EyeHt = (FaceMaskChartA(3, i))
                End If
            Next i
        End If

        ' Set Details
        txtLipStrapWidth.Text = CStr(LipWidth)
        txtMidToEyeTop.Text = CStr(EyeHt)

        ' Mark Mouth Height
        PR_MakeXY(xyMouthEnd, PointOne, PointTwo + MouthHt + LipWidth)

        If optFaceMask.Checked = True Then
            PR_DrawLine(xyMouthStart, xyMouthEnd)
            PR_AddEntityID("Lip")
        End If
        ' set txtNoseBottomY
        txtNoseBottomY.Text = CStr(PointTwo + MouthHt + LipWidth)

        ' Mark Eye Height
        PR_MakeXY(xyEyeStart, PointOne, 0)
        If chkNoseCovering.CheckState = 1 Then
            PR_MakeXY(xyEyeEnd, PointOne, -EyeHt + 0.25)
            txtNoseCoverX.Text = CStr(xyEyeEnd.X)
            txtNoseCoverY.Text = CStr(xyEyeEnd.Y)
        Else
            PR_MakeXY(xyEyeEnd, PointOne, -EyeHt)
        End If
        If optFaceMask.Checked = True Then
            PR_DrawLine(xyEyeStart, xyEyeEnd)
            PR_AddEntityID("EyeBrow")
        End If

    End Sub

    Private Sub PR_SetLayer(ByRef sNewLayer As String)

        'To the DRAFIX macro file (given by the global txtfNum).
        'Write the syntax to set the current LAYER.
        'For this to work it assumes that hLayer is defined in DRAFIX as
        'a HANDLE.
        '
        'To reduce unessesary writing of DRAFIX code check that the new layer
        'is different from the Current layer, change only if it is different.
        '
        PrintLine(CInt(txtfNum.Text), "hLayer = Table(" & HEADNECK1.QQ & "find" & HEADNECK1.QCQ & "layer" & HEADNECK1.QCQ & sNewLayer & HEADNECK1.QQ & ");")
        PrintLine(CInt(txtfNum.Text), "if ( hLayer > %zero && hLayer != 32768)" & "Execute (" & HEADNECK1.QQ & "menu" & HEADNECK1.QCQ & "SetLayer" & HEADNECK1.QC & "hLayer);")

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Layer table for read
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId,
                                     OpenMode.ForRead)

            ' Dim sLayerName As String = "Center"

            If acLyrTbl.Has(sNewLayer) = True Then
                '' Set the layer Center current
                acCurDb.Clayer = acLyrTbl(sNewLayer)

                '' Save the changes
                acTrans.Commit()
            End If

            '' Dispose of the transaction
        End Using

    End Sub

    Private Sub PR_ShowInches(ByRef TextBox As Object, ByRef InchLabel As Object)

        ' Show Values as Inches
        'UPGRADE_WARNING: Couldn't resolve default property of object TextBox. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Dim inchbit As Double
        If Val(TextBox) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inchbit = FN_CmToInches(txtUnits, TextBox)
            'UPGRADE_WARNING: Couldn't resolve default property of object InchLabel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            InchLabel = FN_InchesToText(inchbit)
        End If

    End Sub

    Private Sub PR_StampLining()

        Dim xyStampLining As HEADNECK1.xy

        ' Stamp For Lining
        If chkLining.CheckState = 1 And chkOpenHeadMask.CheckState <> 1 Then
            PR_MakeXY(xyStampLining, -0.5, TopArcOpp(CInt(txtRadiusNo.Text)) - 0.5)
            PR_SetLayer("Notes")
            PR_DrawText("Lining", xyStampLining, 0.25, 0)
            PR_SetLayer("TemplateLeft")
        End If

    End Sub

    Private Sub PR_TemplateDetails(ByRef xyTextPoint As HEADNECK1.xy, ByRef TitleText As String)
        Dim xyNotes As HEADNECK1.xy
        Dim sText As String

        ' Add notes
        PR_SetLayer("Notes")
        ''sText = TitleText & "\n" & txtPatientName.Text & "\n" & txtWorkOrder.Text & "\n" & cboFabric.Text
        sText = TitleText & Chr(10) & txtPatientName.Text & Chr(10) & txtWorkOrder.Text & Chr(10) & cboFabric.Text
        PR_MakeXY(xyNotes, xyTextPoint.X, xyTextPoint.Y - 0.125)
        ''PR_DrawText(sText, xyNotes, 0.1, 0)
        PR_DrawMText(sText, xyNotes)

        PR_SetLayer("Construct")
        ''sText = txtFileNo.Text & "\n" & txtDiagnosis.Text & "\n" & txtAge.Text & "\n" & txtSex.Text
        sText = txtFileNo.Text & Chr(10) & txtDiagnosis.Text & Chr(10) & txtAge.Text & Chr(10) & txtSex.Text
        PR_MakeXY(xyNotes, xyTextPoint.X, xyTextPoint.Y - 0.9)
        ''PR_DrawText(sText, xyNotes, 0.1, 0)
        PR_DrawMText(sText, xyNotes)

        PR_SetLayer("TemplateLeft")

    End Sub

    Private Function round(ByVal nNumber As Double) As Short
        'Fuction to return the rounded value of a decimal number
        'E.G.
        '    round(1.35)  = 1
        '    round(1.55)  = 2
        '    round(2.50)  = 3
        '    round(-2.50) = -3
        '    round(0)     = 0
        '

        Dim iInt, iSign As Short

        'Avoid extra work. Return 0 if input is 0
        If nNumber = 0 Then
            round = 0
            Exit Function
        End If

        'Split input
        iSign = System.Math.Sign(nNumber)
        nNumber = System.Math.Abs(nNumber)
        iInt = Int(nNumber)

        'Effect rounding
        If (nNumber - iInt) >= 0.5 Then
            round = (iInt + 1) * iSign
        Else
            round = iInt * iSign
        End If

    End Function

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        'This is enabled on Load
        'It is disabled in Link close
        'If the event happens then it is assumed that
        'for some reason the link between DRAFIX and This dialogue
        'has failed'
        'Therefor we "End" here
        'End
        Return
    End Sub

    Private Sub txtChinCollarMin_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChinCollarMin.Enter

        ' Highlight text in box
        PR_SelectTextInBox(txtChinCollarMin)

    End Sub

    Private Sub txtChinCollarMin_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChinCollarMin.Leave

        'Check for Numeric Value
        If Not IsNumeric(txtChinCollarMin.Text) And txtChinCollarMin.Text <> "" Then
            MsgBox("Non-Numeric value  has been entered", 48, "Head & Neck")
            txtChinCollarMin.Focus()
        End If

        ' Show Values as Inches
        Dim inchbit As Double
        If Val(txtChinCollarMin.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inchbit = FN_CmToInches(txtUnits, txtChinCollarMin)
            lblInch11.Text = FN_InchesToText(inchbit)
        Else
            lblInch11.Text = ""
        End If

        ' Ensure Value is not Greater then The Throat to Sternal Notch value
        If txtThroatToSternal.Text <> "" Then
            If Val(txtChinCollarMin.Text) >= Val(txtThroatToSternal.Text) Then
                MsgBox("Chin Collar Min Cannot be Greater than the Throat to Sternal Notch Distance", 48, "Head & Neck")
                txtChinCollarMin.Focus()
            End If
        End If

    End Sub

    Private Sub txtChinToMouth_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChinToMouth.Enter

        ' Highlight text in box
        PR_SelectTextInBox(txtChinToMouth)

    End Sub

    Private Sub txtChinToMouth_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChinToMouth.Leave

        'Check for Numeric Value
        If Not IsNumeric(txtChinToMouth.Text) And txtChinToMouth.Text <> "" Then
            MsgBox("Non-Numeric value  has been entered", 48, "Head & Neck")
            txtChinToMouth.Focus()
        End If

        ' Show Values as Inches
        Dim inchbit As Double
        If Val(txtChinToMouth.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inchbit = FN_CmToInches(txtUnits, txtChinToMouth)
            lblInch1.Text = FN_InchesToText(inchbit)
        Else
            lblInch1.Text = ""
        End If

    End Sub

    Private Sub txtCircChinAngle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCircChinAngle.Enter

        ' Highlight text in box
        PR_SelectTextInBox(txtCircChinAngle)

    End Sub

    Private Sub txtCircChinAngle_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCircChinAngle.Leave

        'Check for Numeric Value
        If Not IsNumeric(txtCircChinAngle.Text) And txtCircChinAngle.Text <> "" Then
            MsgBox("Non-Numeric value  has been entered", 48, "Head & Neck")
            txtCircChinAngle.Focus()
        End If

        ' Show Values as Inches
        Dim inchbit As Double
        If Val(txtCircChinAngle.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inchbit = FN_CmToInches(txtUnits, txtCircChinAngle)
            lblInch3.Text = FN_InchesToText(inchbit)
        Else
            lblInch3.Text = ""
        End If

    End Sub

    Private Sub txtCircEyeBrow_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCircEyeBrow.Enter

        ' Highlight text in box
        PR_SelectTextInBox(txtCircEyeBrow)

    End Sub

    Private Sub txtCircEyeBrow_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCircEyeBrow.Leave

        'Check for Numeric Value
        If Not IsNumeric(txtCircEyeBrow.Text) And txtCircEyeBrow.Text <> "" Then
            MsgBox("Non-Numeric value  has been entered", 48, "Head & Neck")
            txtCircEyeBrow.Focus()
        End If

        ' Show Values as Inches
        Dim inchbit As Double
        If Val(txtCircEyeBrow.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inchbit = FN_CmToInches(txtUnits, txtCircEyeBrow)
            lblInch2.Text = FN_InchesToText(inchbit)
        Else
            lblInch2.Text = ""
        End If

    End Sub

    Private Sub txtCircOfNeck_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCircOfNeck.Enter

        ' Highlight text in box
        PR_SelectTextInBox(txtCircOfNeck)

    End Sub

    Private Sub txtCircOfNeck_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCircOfNeck.Leave

        'Check for Numeric Value
        If Not IsNumeric(txtCircOfNeck.Text) And txtCircOfNeck.Text <> "" Then
            MsgBox("Non-Numeric value  has been entered", 48, "Head & Neck")
            txtCircOfNeck.Focus()
        End If

        ' Show Values as Inches
        Dim inchbit As Double
        If Val(txtCircOfNeck.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inchbit = FN_CmToInches(txtUnits, txtCircOfNeck)
            lblInch4.Text = FN_InchesToText(inchbit)
        Else
            lblInch4.Text = ""
        End If

    End Sub

    Private Sub txtHeadBandDepth_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHeadBandDepth.Enter

        ' Highlight text in box
        PR_SelectTextInBox(txtHeadBandDepth)

    End Sub

    Private Sub txtHeadBandDepth_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHeadBandDepth.Leave

        'Check for Numeric Value
        If Not IsNumeric(txtHeadBandDepth.Text) And txtHeadBandDepth.Text <> "" Then
            MsgBox("Non-Numeric value  has been entered", 48, "Head & Neck")
            txtHeadBandDepth.Focus()
        End If

        ' Show Values as Inches
        Dim inchbit As Double
        If Val(txtHeadBandDepth.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inchbit = FN_CmToInches(txtUnits, txtHeadBandDepth)
            lblInch8.Text = FN_InchesToText(inchbit)
        Else
            lblInch8.Text = ""
        End If


    End Sub

    Private Sub txtLeftEarLength_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeftEarLength.Enter

        ' Highlight text in box
        PR_SelectTextInBox(txtLeftEarLength)

    End Sub

    Private Sub txtLeftEarLength_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLeftEarLength.Leave

        'Check for Numeric Value
        If Not IsNumeric(txtLeftEarLength.Text) And txtLeftEarLength.Text <> "" Then
            MsgBox("Non-Numeric value  has been entered", 48, "Head & Neck")
            txtLeftEarLength.Focus()
        End If

        ' Show Values as Inches
        Dim inchbit As Double
        If Val(txtLeftEarLength.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inchbit = FN_CmToInches(txtUnits, txtLeftEarLength)
            lblInch9.Text = FN_InchesToText(inchbit)
        Else
            lblInch9.Text = ""
        End If


    End Sub

    Private Sub txtLengthOfNose_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLengthOfNose.Enter

        ' Highlight text in box
        PR_SelectTextInBox(txtLengthOfNose)

    End Sub

    Private Sub txtLengthOfNose_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLengthOfNose.Leave

        'Check for Numeric Value
        If Not IsNumeric(txtLengthOfNose.Text) And txtLengthOfNose.Text <> "" Then
            MsgBox("Non-Numeric value  has been entered", 48, "Head & Neck")
            txtLengthOfNose.Focus()
        End If

        ' Show Values as Inches
        Dim inchbit As Double
        If Val(txtLengthOfNose.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inchbit = FN_CmToInches(txtUnits, txtLengthOfNose)
            lblInch7.Text = FN_InchesToText(inchbit)
        Else
            lblInch7.Text = ""
        End If

    End Sub

    Private Sub txtRightEarLength_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRightEarLength.Enter

        ' Highlight text in box
        PR_SelectTextInBox(txtRightEarLength)

    End Sub

    Private Sub txtRightEarLength_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRightEarLength.Leave

        'Check for Numeric Value
        If Not IsNumeric(txtRightEarLength.Text) And txtRightEarLength.Text <> "" Then
            MsgBox("Non-Numeric value  has been entered", 48, "Head & Neck")
            txtRightEarLength.Focus()
        End If

        ' Show Values as Inches
        Dim inchbit As Double
        If Val(txtRightEarLength.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inchbit = FN_CmToInches(txtUnits, txtRightEarLength)
            lblInch10.Text = FN_InchesToText(inchbit)
        Else
            lblInch10.Text = ""
        End If


    End Sub

    Private Sub txtThroatToSternal_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtThroatToSternal.Enter

        ' Highlight text in box
        PR_SelectTextInBox(txtThroatToSternal)

    End Sub

    Private Sub txtThroatToSternal_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtThroatToSternal.Leave

        'Check for Numeric Value
        If Not IsNumeric(txtThroatToSternal.Text) And txtThroatToSternal.Text <> "" Then
            MsgBox("Non-Numeric value  has been entered", 48, "Head & Neck")
            txtThroatToSternal.Focus()
        End If
        Dim totval As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object totval. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        totval = Val(txtThroatToSternal.Text)

        ' Show Values as Inches
        Dim inchbit As Double
        If Val(txtThroatToSternal.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inchbit = FN_CmToInches(txtUnits, txtThroatToSternal)
            lblInch5.Text = FN_InchesToText(inchbit)
        Else
            lblInch5.Text = ""
        End If

        ' Ensure Chin Collar Min is not Greater then The Throat to Sternal Notch value
        If txtThroatToSternal.Text <> "" Then
            If Val(txtChinCollarMin.Text) >= Val(txtThroatToSternal.Text) Then
                MsgBox("Chin Collar Min Cannot be Greater than the Throat to Sternal Notch Distance", 48, "Head & Neck")
                txtChinCollarMin.Focus()
            End If
        End If

    End Sub

    Private Sub txtTipOfNose_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipOfNose.Enter

        ' Highlight text in box
        PR_SelectTextInBox(txtTipOfNose)

    End Sub

    Private Sub txtTipOfNose_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTipOfNose.Leave

        'Check for Numeric Value
        If Not IsNumeric(txtTipOfNose.Text) And txtTipOfNose.Text <> "" Then
            MsgBox("Non-Numeric value  has been entered", 48, "Head & Neck")
            txtTipOfNose.Focus()
        End If

        ' Show Values as Inches
        Dim inchbit As Double
        If Val(txtTipOfNose.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_CmToInches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inchbit = FN_CmToInches(txtUnits, txtTipOfNose)
            lblInch6.Text = FN_InchesToText(inchbit)
        Else
            lblInch6.Text = ""
        End If

    End Sub
    'To Get Insertion point from user
    Private Sub PR_GetInsertionPoint()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        pPtOpts.Message = vbLf & "Indicate Start Point "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        If pPtRes.Status = PromptStatus.Cancel Then
            Exit Sub
        End If
        Dim ptStart As Point3d = pPtRes.Value
        PR_MakeXY(xyInsertion, ptStart.X, ptStart.Y)
    End Sub
    'To Draw Cross Block
    Private Sub PR_DrawCrossBlock()
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has("CROSS") Then
                Dim blkTblRecCross As BlockTableRecord = New BlockTableRecord()
                blkTblRecCross.Name = "CROSS"
                Dim acLine As Line = New Line(New Point3d(0 - 0.0625, 0, 0), New Point3d(0 + 0.0625, 0, 0))
                blkTblRecCross.AppendEntity(acLine)
                acLine = New Line(New Point3d(0, 0 + 0.0625, 0), New Point3d(0, 0 - 0.0625, 0))
                blkTblRecCross.AppendEntity(acLine)
                acBlkTbl.UpgradeOpen()
                acBlkTbl.Add(blkTblRecCross)
                acTrans.AddNewlyCreatedDBObject(blkTblRecCross, True)
                blkRecId = blkTblRecCross.Id
            Else
                blkRecId = acBlkTbl("CROSS")
            End If
            ' Insert the block into the current space
            If blkRecId <> ObjectId.Null Then
                'Create new block reference 
                'Dim ptPosition As Point3d = New Point3d(xyInsertion.X, xyInsertion.Y, 0)
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyInsertion.X, xyInsertion.Y, 0), blkRecId)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout), OpenMode.ForWrite)
                acBlkTblRec.AppendEntity(blkRef)
                acTrans.AddNewlyCreatedDBObject(blkRef, True)
            End If
            '' Save the new object to the database
            acTrans.Commit()
        End Using

    End Sub
    'To Draw Spline
    Private Sub PR_DrawSpline(ByRef PointCollection As Point3dCollection)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)
            '' Get a 3D vector from the point (0.5,0.5,0)
            Dim vecTan As Vector3d = New Point3d(0, 0, 0).GetAsVector
            '' Create a spline through (0, 0, 0), (5, 5, 0), and (10, 0, 0) with a
            '' start and end tangency of (0.5, 0.5, 0.0)
            Using acSpline As Spline = New Spline(PointCollection, vecTan, vecTan, 4, 0.0)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    acSpline.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
                End If
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acSpline)
                acTrans.AddNewlyCreatedDBObject(acSpline, True)
            End Using
            '' Save the new object to the database
            acTrans.Commit()
        End Using
    End Sub
    Private Sub saveInfoToDWG()
        Try
            Dim _sClass As New SurroundingClass()
            If (_sClass.GetXrecord("HeadInfo", "HEADDIC") IsNot Nothing) Then
                _sClass.RemoveXrecord("HeadInfo", "HEADDIC")
            End If

            Dim resbuf As New ResultBuffer

            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtChinToMouth.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtCircEyeBrow.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtCircChinAngle.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtCircOfNeck.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtThroatToSternal.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtTipOfNose.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtLengthOfNose.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtHeadBandDepth.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtLeftEarLength.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtRightEarLength.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtChinCollarMin.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboFabric.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkLeftEyeFlap.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkRightEyeFlap.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkLeftEarFlap.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkRightEarFlap.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkLeftEarClosed.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkRightEarClosed.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkEarSize.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkNeckElastic.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkLipStrap.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkEyes.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkLipCovering.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkLining.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkNoseCovering.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkZipper.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkOpenHeadMask.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), chkVelcro.CheckState))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optFaceMask.Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optOpenFaceMask.Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optChinStrap.Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optModifiedChinStrap.Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optChinCollar.Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optContouredChinCollar.Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optHeadBand.Checked))
            _sClass.SetXrecord(resbuf, "HeadInfo", "HEADDIC")

        Catch ex As Exception

        End Try
    End Sub
    Private Sub readDWGInfo()
        Try

            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
            resbuf = _sClass.GetXrecord("HeadInfo", "HEADDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                txtChinToMouth.Text = arr(0).Value
                txtCircEyeBrow.Text = arr(1).Value
                txtCircChinAngle.Text = arr(2).Value
                txtCircOfNeck.Text = arr(3).Value
                txtThroatToSternal.Text = arr(4).Value
                txtTipOfNose.Text = arr(5).Value
                txtLengthOfNose.Text = arr(6).Value
                txtHeadBandDepth.Text = arr(7).Value
                txtLeftEarLength.Text = arr(8).Value
                txtRightEarLength.Text = arr(9).Value
                txtChinCollarMin.Text = arr(10).Value
                cboFabric.Text = arr(11).Value
                chkLeftEyeFlap.CheckState = arr(12).Value
                chkRightEyeFlap.CheckState = arr(13).Value
                chkLeftEarFlap.CheckState = arr(14).Value
                chkRightEarFlap.CheckState = arr(15).Value
                chkLeftEarClosed.CheckState = arr(16).Value
                chkRightEarClosed.CheckState = arr(17).Value
                chkEarSize.CheckState = arr(18).Value
                chkNeckElastic.CheckState = arr(19).Value
                chkLipStrap.CheckState = arr(20).Value
                chkEyes.CheckState = arr(21).Value
                chkLipCovering.CheckState = arr(22).Value
                chkLining.CheckState = arr(23).Value
                chkNoseCovering.CheckState = arr(24).Value
                chkZipper.CheckState = arr(25).Value
                chkOpenHeadMask.CheckState = arr(26).Value
                chkVelcro.CheckState = arr(27).Value
                optFaceMask.Checked = arr(28).Value
                optOpenFaceMask.Checked = arr(29).Value
                optChinStrap.Checked = arr(30).Value
                optModifiedChinStrap.Checked = arr(31).Value
                optChinCollar.Checked = arr(32).Value
                optContouredChinCollar.Checked = arr(33).Value
                optHeadBand.Checked = arr(34).Value

            End If
            resbuf = _sClass.GetXrecord("NewPatient", "NEWPATIENTDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                Dim pattern As String = "dd-MM-yyyy"
                Dim parsedDate As DateTime
                DateTime.TryParseExact(arr(8).Value, pattern, System.Globalization.CultureInfo.CurrentCulture,
                                  System.Globalization.DateTimeStyles.None, parsedDate)
                DateTimePicker1.Value = parsedDate
                ''CurrentYear - BirthDate  
                Dim startTime As DateTime = Convert.ToDateTime(DateTimePicker1.Value)
                Dim endTime As DateTime = DateTime.Today
                Dim span As TimeSpan = endTime.Subtract(startTime)
                Dim totalDays As Double = span.TotalDays
                Dim totalYears As Double = Math.Truncate(totalDays / 365)
                txtAge.Text = totalYears.ToString()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub HeadNeck_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If g_bIsClose = False Then
            Dim Response As Short
            Dim DataString, MeasurementString As String
            Dim Options As String
            ' Check For any Changes
            MeasurementString = FN_SetMeasurementString()
            DataString = FN_ModificationString()
            Options = FN_ChooseDesign()
            DataString = Options & DataString

            If (DataString = txtData.Text And MeasurementString = txtMeasurements.Text) Or (txtData.Text = "txtData" And txtMeasurements.Text = "txtMeasurements") Then
                Exit Sub
            Else
                ' changes have been made, find out what to do?
                Response = MsgBox("Changes have been made, Save changes before closing", 35, "Head & Neck Details")
                ' Determine response of message box
                If Response = 6 Then
                    ' Yes button selected
                    txtData.Text = DataString
                    txtMeasurements.Text = MeasurementString
                    Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
                    txtfNum.Text = FN_OpenSave(sDrawFile, txtPatientName.Text, txtFileNo.Text)
                    PR_SaveDetails()
                    FileClose(CInt(txtfNum.Text))
                    saveInfoToDWG()
                ElseIf Response = 2 Then
                    e.Cancel = True
                End If
            End If
        End If
        g_bIsClose = False
    End Sub

    Private Sub PR_DrawMText(ByRef sText As Object, ByRef xyInsert As HEADNECK1.xy)
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                     OpenMode.ForRead)
            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(strLayout),
                                        OpenMode.ForWrite)

            Dim mtx As New MText()
            mtx.Location = New Point3d(xyInsert.X + xyInsertion.X, xyInsert.Y + xyInsertion.Y, 0)
            mtx.SetDatabaseDefaults()
            mtx.TextStyleId = acCurDb.Textstyle
            ' current text size
            mtx.TextHeight = 0.1
            ' current textstyle
            mtx.Width = 0.0
            mtx.Rotation = 0
            mtx.Contents = sText
            mtx.Attachment = AttachmentPoint.TopLeft
            mtx.SetAttachmentMovingLocation(mtx.Attachment)
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                mtx.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.Y, 0)))
            End If
            acBlkTblRec.AppendEntity(mtx)
            acTrans.AddNewlyCreatedDBObject(mtx, True)

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
End Class