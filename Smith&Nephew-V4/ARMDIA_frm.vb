Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports VB = Microsoft.VisualBasic

Public Class armdia
    Inherits System.Windows.Forms.Form
    'Project:   ARMDIA.MAK
    'Purpose:   Drawing Arms and Vest/Bodysuit sleeves.
    '
    'Version:   3.07
    'Date:      94/95
    'Author:    Ciaran McKavangh    (Armdia.FRM, mostly)
    '           Gary George         (Utility.BAS)
    '-------------------------------------------------------
    'REVISIONS:
    'Date       By      Action
    '-------------------------------------------------------
    'Feb.95     CMcK    Error and bug fixes
    '                   G.Dunne Fax, 20.Feb.95
    '
    'Mar.95     GG      Error and bug fixes
    '                   G.Dunne Fax, 20.Mar.95 and 03.Apr.95    z
    '
    'Apr.95     GG      Allow multiple ARM styles in a single
    '                   drawing
    '
    'Jul.95     GG      Triton to CAD link
    '
    'Oct.95     GG      Triton to CAD link
    '                   and Sleeveles Vests
    '
    'Dec.95     GG      Mods to decrease code
    '
    'Sep.97     GG      Added support for bodysuit and bodybrief
    '
    'Jan.98     GG      Added support for new mesh method.
    '                   Added support for side scooped necks
    '
    'Dec 98     GG      Ported to VB5
    'Note:-
    '
    '
    'UPGRADE_WARNING: Lower bound of array Measurements was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim Measurements(18) As Object
    'UPGRADE_WARNING: Lower bound of array ChangeChecker was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim ChangeChecker(18) As Object
    'UPGRADE_WARNING: Lower bound of array MMsChecker was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim MMsChecker(18) As Object
    'UPGRADE_WARNING: Lower bound of array tapewidths was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim tapewidths(18) As Object
    'UPGRADE_WARNING: Lower bound of array Pow was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim Pow(19, 23) As Object
    'UPGRADE_WARNING: Lower bound of array Bob was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim Bob(19, 23) As Object
    Dim nine, seven, five, three, one, Zero, two, four, six, eight, ten As Object

    '* Windows API Functions Declarations
    ' Private Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
    ' Private Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
    ' Private Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer
    ' Private Declare Function GetWindowsDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
    ' Private Declare Function GetPrivateProfileString% Lib "Kernel" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)

    'Constanst used by GetWindow
    ' Const GW_CHILD = 5
    ' Const GW_HWNDFIRST = 0
    ' Const GW_HWNDLAST = 1
    ' Const GW_HWNDNEXT = 2
    ' Const GW_HWNDPREV = 3
    ' Const GW_OWNER = 4

    'MsgBox constant
    '    Const IDCANCEL = 2
    '    Const IDYES = 6
    '    Const IDNO = 7

    Dim EmptySleeve As Object 'Flag to suppress calculation
    Dim g_sAlreadyLoad As Boolean
    Dim strFirstTape, strSecondTape, strLastTape, strSecondLastTape As String
    Dim RightPow(19, 23) As Object
    'UPGRADE_WARNING: Lower bound of array Bob was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Dim RightBob(19, 23) As Object
    Dim RightMeasurements(18) As Object
    Dim RightChangeChecker(18) As Object
    Dim Righttapewidths(18) As Object
    Dim RightEmptySleeve As Object
    Dim strWristNo, strPalmNo, strThumbDist, strThumbLen, strTCircum As String
    Dim g_bIsClose As Boolean
    Dim g_sLeftArmValues, g_sRightArmValues As String



    Private Sub BackCalcMMs()
        Dim X As Object
        Dim Modulus As Object
        Dim w As Object
        Dim Backgram As Object
        If Length(w).Text <> "" Then
            ' check MMs
            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (w < 9) And (w = (CDbl(txtFirstTape.Text) - 1)) And txtStump.Text = "False" And (Val(CStr(CDbl(reds(w).Text) > 14))) Then
                reds(w).Text = CStr(14)
                If VB.Left(g_sFabnam, 3) = "Pow" Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modulus, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Backgram = Pow(Modulus, 5)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modulus, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Backgram = Bob(Modulus, 5)
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                gms(w).Text = CStr(ARMDIA1.round(Backgram))
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (w < 9) And (w = (CDbl(txtFirstTape.Text) - 1)) And (Val(CStr(CDbl(reds(w).Text) < 10))) Then
                If VB.Left(g_sFabnam, 3) = "Pow" Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modulus, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Backgram = Pow(Modulus, 1)
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If Val(gms(w).Text) < Backgram Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                        reds(w).Text = CStr(10)
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        gms(w).Text = CStr(ARMDIA1.round(Backgram))
                    End If
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modulus, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Backgram = Bob(Modulus, 1)
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If Val(gms(w).Text) < Backgram Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                        reds(w).Text = CStr(10)
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        gms(w).Text = CStr(ARMDIA1.round(Backgram))
                    End If
                End If
            End If
        End If
        ' end check wrist reduction


    End Sub

    Private Function BobReduction(ByRef Modul As Object, ByRef lblGram As Object) As Object
        Dim lblred As Object
        Dim j As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        one = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object two. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        two = 2
        'UPGRADE_WARNING: Couldn't resolve default property of object nine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nine = 9
        'UPGRADE_WARNING: Couldn't resolve default property of object eight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        eight = 8
        'UPGRADE_WARNING: Couldn't resolve default property of object ten. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ten = 10
        'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        j = one
        'UPGRADE_WARNING: Couldn't resolve default property of object lblGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Bob(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (Val(Bob(Modul, 1)) = Val(lblGram)) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object ten. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            lblred = ten
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object two. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        j = two
        'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While j <= 22
            'UPGRADE_WARNING: Couldn't resolve default property of object lblGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Bob(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (Val(Bob(Modul, j)) = Val(lblGram)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object nine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                lblred = Val(j + nine)
                'UPGRADE_WARNING: Couldn't resolve default property of object lblGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Bob(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (Val(lblGram) - Val(Bob(Modul, j - one))) <= (Val(Bob(Modul, j)) - Val(lblGram)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object eight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                lblred = Val(j + eight)
                Exit Do
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            j = j + one
        Loop
        'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object BobReduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        BobReduction = lblred
    End Function

    Private Sub Calculate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Calculate.Click
        'Force a recalculation
        '
        If txtSleeve.Text = "Right" Then
            RightCalculate_Click()
            Exit Sub
        End If
        Dim ii As Short
        '
        If cboPressure.Text = "" Then
            If Me.Visible = True Then MsgBox("Pressure has not been selected", 48, "Arm Details")
            Exit Sub
        End If
        If cboFabric.Text = "" Then
            If Me.Visible = True Then MsgBox("Fabric has not been selected", 48, "Arm Details")
            Exit Sub
        End If

        PR_CheckValues()

        If cboProximalTape.SelectedIndex = 0 Then cboProximalTape_SelectedIndexChanged(cboProximalTape, New System.EventArgs())
        If cboDistalTape.SelectedIndex = 0 Then cboDistalTape_SelectedIndexChanged(cboDistalTape, New System.EventArgs())
        If ARMDIA1.g_iStyleLastTape < ARMDIA1.g_iStyleFirstTape Then
            If Me.Visible = True Then MsgBox("Chosen distal tape is above Proximal tape!", 48, "Arm Details")
            Exit Sub
        End If

        If chkGauntlets.Checked = True Then
            If cboPalm.SelectedIndex < 0 Then
                If Me.Visible = True Then MsgBox("Palm tape position not given", 48, "Gauntlet Details")
                Exit Sub
            End If
            If cboWrist.SelectedIndex < 0 Then
                If Me.Visible = True Then MsgBox("Wrist tape position not given", 48, "Gauntlet Details")
                Exit Sub
            End If
            PR_GetGauntletDetails()
            PR_GauntletHoleCheck()
        End If

        g_sPressureChange = "True" 'Use this to force a recalculation
        PR_GetStartEndTapes()
        PR_GetPressure()
        PR_GetFabric()
        PR_GetGauntletDetails()
        For ii = 0 To 17
            If Val(Length(ii).Text) > 0 Then
                PR_Pressurecalc(ii)
            End If
        Next ii
        PR_SetRresults()
        PR_SetGlobals()
        g_sPressureChange = "False"

        'Clean up displayed values
        For ii = 0 To CDbl(txtFirstTape.Text) - 2
            If mms(ii).Text = "" Or mms(ii).Text = "0" Then
                mms(ii).Text = ""
                gms(ii).Text = ""
                reds(ii).Text = ""
            End If
        Next ii

        For ii = 17 To CInt(txtLastTape.Text) Step -1
            If mms(ii).Text = "" Or mms(ii).Text = "0" Then
                mms(ii).Text = ""
                gms(ii).Text = ""
                reds(ii).Text = ""
            End If
        Next ii

        If Me.Visible = True Then ARMDIA1.g_Modified = True
        'UPGRADE_WARNING: Couldn't resolve default property of object EmptySleeve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        EmptySleeve = False

    End Sub

    'UPGRADE_WARNING: Event cboDistalTape.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboDistalTape_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboDistalTape.SelectedIndexChanged
        'Set style first tape
        'Modify tape labels to show this (lblTape index starts at 0)
        'NOTE:-
        '    cboDistalTape.ListIndex = 0 = 1st
        '    cboDistalTape.ListIndex = 1 = -6
        '    cboDistalTape.ListIndex = 2 = -4½
        '    cboDistalTape.ListIndex = 3 = -3
        '                .
        '                .
        '                .
        '    cboDistalTape.ListIndex = 18 = 19½
        '
        Dim iTape As Short

        'Reset display of previous to Black
        If ARMDIA1.g_iStyleFirstTape > 0 Then
            lblTape(ARMDIA1.g_iStyleFirstTape - 1).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 0, 0))
        End If

        'set style first tape value
        iTape = cboDistalTape.SelectedIndex - 1
        If iTape < 0 Then
            'Use last tape
            If ARMDIA1.g_iFirstTape > 0 Then
                ARMDIA1.g_iStyleFirstTape = ARMDIA1.g_iFirstTape
            Else
                ARMDIA1.g_iStyleFirstTape = -1
            End If
        Else
            'use tape pointed at by index
            ARMDIA1.g_iStyleFirstTape = cboDistalTape.SelectedIndex
        End If

        'Set colour of tape label to green
        iTape = ARMDIA1.g_iStyleFirstTape - 1 'lblTape Index starts at 0
        If iTape >= 0 Then
            lblTape(iTape).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 255, 0))
        End If

        If Me.Visible = True Then ARMDIA1.g_Modified = True

    End Sub

    'UPGRADE_WARNING: Event cboFabric.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboFabric_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFabric.SelectedIndexChanged
        Dim fabnam, Fabtype As Object
        ' check for fabric modulus
        If cboFabric.Text <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            fabnam = VB.Left(cboFabric.Text, 3)
            'UPGRADE_WARNING: Couldn't resolve default property of object Fabtype. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Fabtype = Mid(cboFabric.Text, 5, 3)
            'UPGRADE_WARNING: Couldn't resolve default property of object Fabtype. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Val(Fabtype) > 340 Or Val(Fabtype) < 160 Then
                MsgBox("Fabric Modulus Not Recognised", 48, "Arm Details")
                cboFabric.Text = ""
                Exit Sub
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (fabnam <> "Pow") And (fabnam <> "Bob") Then
                MsgBox("Fabric Type Not Recognised, Should be either Bobbinette or Powernet", 48, "Arm Details")
                cboFabric.Text = ""
                Exit Sub
            End If
            If Me.Visible = True Then ARMDIA1.g_Modified = True

        End If

    End Sub

    'UPGRADE_WARNING: Event cboFlaps.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboFlaps_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFlaps.SelectedIndexChanged
        Dim w, X As Object
        If cboFlaps.Text = "none" Then
            cboFlaps.Text = ""

            labStrap.Text = ""
            txtCustFlapLength.Text = ""
            labCustFlapLength.Text = ""
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X = Val(txtLastTape.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        w = X - 1
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If X > 1 Then
            If Length(w).Text <> "" And cboFabric.Text <> "" And cboPressure.Text <> "" And gms(w).Text <> "" And reds(w).Text <> "" Then
                txtFlap.Text = cboFlaps.Text
                PR_LastTapeRed(w, X, g_sModulus)
                PR_CheckFlaps(w)
                PR_SetRresults()
                PR_SetGlobals()
            End If
        End If

        'Enable waist circumference for style "D"
        If InStr(1, cboFlaps.Text, "D", 0) > 0 Then
            labWaistCir.Enabled = True 'Inches convesion
            txtWaistCir.Enabled = True 'Edit box
            lblWaistCir.Enabled = True
        ElseIf txtWaistCir.Enabled = True Then
            'Disable only if previously enabled
            txtWaistCir.Text = "" 'Delete display of value
            labWaistCir.Text = "" 'Delete inches display of value
            labWaistCir.Enabled = False
            txtWaistCir.Enabled = False
            lblWaistCir.Enabled = False
        End If

        'Disable all boxes if Vest Raglan chosen
        If VB.Left(cboFlaps.Text, 4) = "Vest" Or VB.Left(cboFlaps.Text, 4) = "Body" Then
            txtStrapLength.Text = ""
            txtStrapLength.Enabled = False
            txtFrontStrapLength.Text = ""
            txtFrontStrapLength.Enabled = False
            lblFrontStrapLength.Enabled = False
            labFrontStrapLength.Text = ""
            labStrap.Text = ""
            lblStrapLength.Enabled = False
            txtCustFlapLength.Text = ""
            txtCustFlapLength.Enabled = False
            labCustFlapLength.Text = ""
            lblCustFlapLength.Enabled = False
        Else
            'Enable if not enbaled only
            If txtStrapLength.Enabled = False Then
                txtStrapLength.Enabled = True
                lblStrapLength.Enabled = True
            End If
            If txtCustFlapLength.Enabled = False Then
                txtCustFlapLength.Enabled = True
                lblCustFlapLength.Enabled = True
            End If
            If txtFrontStrapLength.Enabled = False Then
                txtFrontStrapLength.Enabled = True
                lblFrontStrapLength.Enabled = True
            End If
        End If
        If Me.Visible = True Then ARMDIA1.g_Modified = True


    End Sub

    Private Sub cboPalm_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboPalm.Leave
        If cboWrist.SelectedIndex > -1 Then
            If cboWrist.SelectedIndex <= cboPalm.SelectedIndex Then
                MsgBox("Palm No. must be less than Wrist No.", 48, "Gauntlet Details")
                cboWrist.Focus()
                Exit Sub
            End If
        End If

        If Me.Visible = True Then ARMDIA1.g_Modified = True

    End Sub

    'UPGRADE_WARNING: Event cboPressure.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboPressure_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboPressure.SelectedIndexChanged
        Dim Index As Short

        If Me.Visible = True Then ARMDIA1.g_Modified = True

    End Sub

    'UPGRADE_WARNING: Event cboProximalTape.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboProximalTape_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboProximalTape.SelectedIndexChanged
        'Set style last tape
        'Modify tape labels to show this (lblTape index starts at 0)
        'NOTE:-
        '    cboProximal.ListIndex = 0 = Last
        '    cboProximal.ListIndex = 1 = 19½
        '    cboProximal.ListIndex = 2 = 18
        '    cboProximal.ListIndex = 3 = 16½
        '                .
        '                .
        '                .
        '    cboProximal.ListIndex = 18 = -6
        '
        Dim iTape As Short

        'Reset display of previous to Black
        If ARMDIA1.g_iStyleLastTape > 0 Then
            lblTape(ARMDIA1.g_iStyleLastTape - 1).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 0, 0))
        End If

        'set style last tape value
        iTape = cboProximalTape.SelectedIndex - 1
        If iTape < 0 Then
            'Use last tape
            If ARMDIA1.g_iLastTape > 0 Then
                ARMDIA1.g_iStyleLastTape = ARMDIA1.g_iLastTape
            Else
                ARMDIA1.g_iStyleLastTape = -1
            End If
        Else
            'use tape pointed at by index
            ARMDIA1.g_iStyleLastTape = (18 - cboProximalTape.SelectedIndex) + 1
        End If

        'Set colour of tape label to red
        iTape = ARMDIA1.g_iStyleLastTape - 1 'lblTape Index starts at 0
        If iTape >= 0 Then
            lblTape(iTape).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(255, 0, 0))
        End If


        If Me.Visible = True Then ARMDIA1.g_Modified = True

    End Sub

    Private Sub cboWrist_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboWrist.Leave
        If cboPalm.SelectedIndex > -1 Then
            If cboPalm.SelectedIndex >= cboWrist.SelectedIndex Then
                MsgBox("Wrist No. must be greater than Palm No.", 48, "Gauntlet Details")
                cboPalm.Focus()
                Exit Sub
            End If
        End If

        If Me.Visible = True Then ARMDIA1.g_Modified = True

    End Sub

    'UPGRADE_WARNING: Event chkDetachable.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkDetachable_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDetachable.CheckStateChanged
        'Setup for Detachable gauntlet

        If chkDetachable.CheckState = 1 Then
            'For detachable gauntlets
            'Note:- cboProximalTape gives the tapes in reverse order
            cboProximalTape.SelectedIndex = 13 '+1-1/2 tape

            'If Val(txtAge) > 14 Then
            '    cboProximalTape.ListIndex = 12 '+3 tape for adults
            'Else
            '    cboProximalTape.ListIndex = 13 '+1-1/2 tape for children
            'End If
            cboProximalTape_SelectedIndexChanged(cboProximalTape, New System.EventArgs())
            optProximalTape(0).Checked = 1
            optProximalTape(1).Checked = 0
        ElseIf chkDetachable.CheckState = 0 Then
            'For normal gauntlets us last tape
            cboProximalTape.SelectedIndex = 0
            cboProximalTape_SelectedIndexChanged(cboProximalTape, New System.EventArgs())
        End If
    End Sub

    'UPGRADE_WARNING: Event chkEnclosed.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkEnclosed_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkEnclosed.CheckStateChanged
        If chkEnclosed.CheckState = 1 Then
            chkNoThumb.Enabled = False
            chkNoThumb.CheckState = False
            txtCircum.Enabled = True
            txtThumb.Enabled = True
            lblCircum.Enabled = True
            lblThumb.Enabled = True
        ElseIf chkEnclosed.CheckState = 0 Then
            chkNoThumb.Enabled = True

        End If

        If Me.Visible = True Then ARMDIA1.g_Modified = True

    End Sub

    'UPGRADE_WARNING: Event chkGauntlets.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkGauntlets_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkGauntlets.CheckedChanged
        If eventSender.Checked Then
            If chkGauntlets.Checked = True Then
                frmGauntlet.Enabled = True
                chkStump.Checked = False
                lblPalm.Enabled = True
                lblWrist.Enabled = True
                lblThumb.Enabled = True
                cboPalm.Enabled = True
                cboWrist.Enabled = True
                txtThumb.Enabled = True
                chkEnclosed.Enabled = True
                chkDetachable.Enabled = True
                chkNoThumb.Enabled = True
                lblPalmWrist.Enabled = True
                txtPalmWrist.Enabled = True
                txtCircum.Enabled = True
                lblCircum.Enabled = True
                lblGauntletExtension.Enabled = True
                txtGauntletExtension.Enabled = True

                'Set distal tape position to first
                cboDistalTape.Enabled = True
                cboDistalTape.SelectedIndex = 0
                '        cboDistalTape_Click
                cboDistalTape.Enabled = False

            ElseIf chkGauntlets.Checked = False Then
                frmGauntlet.Enabled = False
                chkNoThumb.Enabled = False
                lblPalmWrist.Enabled = False
                txtPalmWrist.Enabled = False
                lblPalm.Enabled = False
                lblWrist.Enabled = False
                lblThumb.Enabled = False
                cboPalm.Enabled = False
                cboWrist.Enabled = False
                txtThumb.Enabled = False
                chkEnclosed.Enabled = False
                chkDetachable.Enabled = False
                txtCircum.Enabled = False
                lblCircum.Enabled = False
                lblGauntletExtension.Enabled = True
                txtGauntletExtension.Enabled = True

            End If

            If Me.Visible = True Then ARMDIA1.g_Modified = True

        End If
    End Sub

    'UPGRADE_WARNING: Event chkNone.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkNone_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkNone.CheckedChanged
        If eventSender.Checked Then
            'Disable gauntlet
            g_sHoleCheck = "False"
            frmGauntlet.Enabled = False
            lblPalm.Enabled = False
            lblWrist.Enabled = False
            lblThumb.Enabled = False
            cboPalm.SelectedIndex = -1
            cboWrist.SelectedIndex = -1
            cboPalm.Enabled = False
            cboWrist.Enabled = False
            lblGauntletExtension.Enabled = False
            txtGauntletExtension.Enabled = False


            If chkDetachable.CheckState = 1 Then
                'Reset the proximal tape to last
                chkDetachable.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkDetachable_CheckStateChanged(chkDetachable, New System.EventArgs())
            End If

            chkNoThumb.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEnclosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            PWDist.Text = ""
            TCirc.Text = ""
            TLen.Text = ""
            txtThumb.Text = ""
            labGauntletExtension.Text = ""
            txtGauntletExtension.Text = ""
            txtThumb.Enabled = False
            chkEnclosed.Enabled = False
            chkDetachable.Enabled = False
            chkNoThumb.Enabled = False
            lblPalmWrist.Enabled = False
            txtPalmWrist.Text = ""
            txtPalmWrist.Enabled = False
            txtCircum.Text = ""
            txtCircum.Enabled = False
            lblCircum.Enabled = False

            txtGauntletFlag.Text = "False"

            'Enable distal tape position
            cboDistalTape.Enabled = True

            'Disable stump
            chkStump.Checked = False
            txtStump.Text = "False"

            Dim w, X As Object
            If Val(txtFirstTape.Text) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                X = Val(txtFirstTape.Text)
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                w = X - 1
                If Length(w).Text <> "" And cboFabric.Text <> "" And cboPressure.Text <> "" And gms(w).Text <> "" And reds(w).Text <> "" Then
                    PR_WristRed(w, X, g_sModulus)
                    PR_Reductions(w, X)
                    PR_SetRresults()
                    PR_SetGlobals()
                End If
            End If

            If Me.Visible = True Then ARMDIA1.g_Modified = True

        End If
    End Sub

    'UPGRADE_WARNING: Event chkNoThumb.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkNoThumb_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkNoThumb.CheckStateChanged
        If chkNoThumb.CheckState = 1 Then
            chkEnclosed.Enabled = False
            chkEnclosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            txtCircum.Enabled = False
            txtThumb.Enabled = False
            lblCircum.Enabled = False
            lblThumb.Enabled = False
            txtCircum.Text = ""
            TCirc.Text = ""
            txtThumb.Text = ""
            TLen.Text = ""

        ElseIf chkNoThumb.CheckState = 0 And chkGauntlets.Checked = True Then
            txtCircum.Enabled = True
            txtThumb.Enabled = True
            lblCircum.Enabled = True
            lblThumb.Enabled = True
            chkEnclosed.Enabled = True
        End If

        If Me.Visible = True Then ARMDIA1.g_Modified = True

    End Sub

    'UPGRADE_WARNING: Event chkStump.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub chkStump_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkStump.CheckedChanged
        If eventSender.Checked Then
            Dim entries, i As Object
            If chkStump.Checked = True Then
                'Disable gauntlet
                g_sHoleCheck = "False"
                frmGauntlet.Enabled = False
                lblPalm.Enabled = False
                lblWrist.Enabled = False
                lblThumb.Enabled = False
                cboPalm.SelectedIndex = -1
                cboPalm.Enabled = False
                cboWrist.SelectedIndex = -1
                cboWrist.Enabled = False
                If chkDetachable.CheckState = 1 Then
                    'Reset the proximal tape to last
                    chkDetachable.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkDetachable_CheckStateChanged(chkDetachable, New System.EventArgs())
                End If
                chkNoThumb.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkEnclosed.CheckState = System.Windows.Forms.CheckState.Unchecked
                txtGauntletExtension.Enabled = False
                lblGauntletExtension.Enabled = False
                labGauntletExtension.Text = ""
                txtGauntletExtension.Text = ""
                PWDist.Text = ""
                TCirc.Text = ""
                TLen.Text = ""
                txtThumb.Text = ""
                txtThumb.Enabled = False
                chkEnclosed.Enabled = False
                chkDetachable.Enabled = False
                chkNoThumb.Enabled = False
                lblPalmWrist.Enabled = False
                txtPalmWrist.Text = ""
                txtPalmWrist.Enabled = False
                txtCircum.Text = ""
                txtCircum.Enabled = False
                lblCircum.Enabled = False
                txtGauntletFlag.Text = "False"


                'Set distal tape position to first
                cboDistalTape.Enabled = True
                cboDistalTape.SelectedIndex = 0
                cboDistalTape_SelectedIndexChanged(cboDistalTape, New System.EventArgs())
                'cboDistalTape.Enabled = False

            ElseIf chkStump.Checked = 0 Then
                chkGauntlets.Enabled = True
            End If

            Dim w, X As Object
            If Val(txtFirstTape.Text) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                X = Val(txtFirstTape.Text)
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                w = X - 1
                If Length(w).Text <> "" And cboFabric.Text <> "" And cboPressure.Text <> "" And gms(w).Text <> "" And reds(w).Text <> "" Then
                    If chkStump.Checked = True Then
                        PR_Reductions(w, X)
                        PR_Stump(w, X)
                        PR_SetRresults()
                        PR_SetGlobals()
                    End If
                End If
            End If

            If Me.Visible = True Then ARMDIA1.g_Modified = True

        End If
    End Sub
    Private Function FN_LeftArmValuesString() As String
        Dim i As Short
        Dim sString As String = ""
        FN_LeftArmValuesString = ""
        For i = 0 To 17
            sString = sString & Length(i).Text
        Next i
        For i = 0 To 17
            sString = sString & mms(i).Text
        Next i
        For i = 0 To 17
            sString = sString & gms(i).Text
        Next i
        For i = 0 To 17
            sString = sString & reds(i).Text
        Next i
        sString = sString & txtWristPleat1.Text & txtWristPleat2.Text & txtShoulderPleat1.Text & txtShoulderPleat2.Text
        sString = sString & cboPressure.Text & cboFabric.Text & Str(chkNone.Checked) & Str(chkStump.Checked) & Str(chkGauntlets.Checked)
        sString = sString & cboDistalTape.Text & cboPalm.Text & cboWrist.Text
        sString = sString & txtPalmWrist.Text & txtCircum.Text & txtThumb.Text & txtGauntletExtension.Text
        sString = sString & Str(chkDetachable.Checked) & Str(chkNoThumb.Checked) & Str(chkEnclosed.Checked)
        sString = sString & Str(optProximalTape(0).Checked) & Str(optProximalTape(1).Checked) & cboProximalTape.Text & cboFlaps.Text
        sString = sString & txtStrapLength.Text & txtFrontStrapLength.Text & txtCustFlapLength.Text & txtWaistCir.Text
        FN_LeftArmValuesString = sString
    End Function
    Private Function FN_RightArmValuesString() As String
        Dim i As Short
        Dim sString As String = ""
        FN_RightArmValuesString = ""
        For i = 0 To 17
            sString = sString & RightLength(i).Text
        Next i
        For i = 0 To 17
            sString = sString & Rightmms(i).Text
        Next i
        For i = 0 To 17
            sString = sString & Rightgms(i).Text
        Next i
        For i = 0 To 17
            sString = sString & Rightreds(i).Text
        Next i
        sString = sString & RighttxtWristPleat1.Text & RighttxtWristPleat2.Text & RighttxtShoulderPleat1.Text & RighttxtShoulderPleat2.Text
        sString = sString & RightcboPressure.Text & RightcboFabric.Text & Str(RightchkNone.Checked) & Str(RightchkStump.Checked) & Str(RightchkGauntlets.Checked)
        sString = sString & RightcboDistalTape.Text & RightcboPalm.Text & RightcboWrist.Text
        sString = sString & RighttxtPalmWrist.Text & RighttxtCircum.Text & RighttxtThumb.Text & RighttxtGauntletExtension.Text
        sString = sString & Str(RightchkDetachable.Checked) & Str(RightchkNoThumb.Checked) & Str(RightchkEnclosed.Checked)
        sString = sString & Str(RightoptProximalTape(0).Checked) & Str(RightoptProximalTape(1).Checked) & RightcboProximalTape.Text & RightcboFlaps.Text
        sString = sString & RighttxtStrapLength.Text & RighttxtFrontStrapLength.Text & RighttxtCustFlapLength.Text & RighttxtWaistCir.Text
        FN_RightArmValuesString = sString
    End Function

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        Dim Response As Short
        Dim sLeftCurrentValue As String = FN_LeftArmValuesString()
        Dim sRightCurrentValue As String = FN_RightArmValuesString()
        Dim bIsChanged As Boolean = False
        If sLeftCurrentValue <> g_sLeftArmValues Then
            bIsChanged = True
        ElseIf sRightCurrentValue <> g_sRightArmValues Then
            bIsChanged = True
        End If
        ''If ARMDIA1.g_Modified Then
        If bIsChanged = True Then
            Response = MsgBox("Changes have been made, Save changes before closing", 35, "Arm Details")
            Select Case Response
                Case IDYES
                    ''ARMDIA1.fNum = ARMDIA1.FN_Open("c:\jobst\draw.d", txtPatientName.Text, txtFileNo.Text, txtSleeve.Text, "Save")
                    Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
                    ARMDIA1.fNum = ARMDIA1.FN_Open(sDrawFile, txtPatientName.Text, txtFileNo.Text, txtSleeve.Text, "Save")
                    'Check Flaps
                    txtFlap.Text = cboFlaps.Text
                    If txtFlap.Text = "" Then g_sFlap = "None" Else g_sFlap = txtFlap.Text
                    If txtFlap.Text <> "" And txtStrapLength.Text = "" Then
                        If txtinchflag.Text = "cm" Then
                            txtStrap.Text = "61"
                        Else
                            txtStrap.Text = "24"
                        End If
                    End If
                    PR_GetStartEndTapes()
                    PR_GetGauntletDetails()
                    PR_SetRresults()
                    PR_SetGlobals()
                    ARMDIA1.g_sFabric = cboFabric.Text
                    'Update
                    DataBaseDataUpDate("Save")
                    FileClose(ARMDIA1.fNum)
                    'AppActivate(fnGetDrafixWindowTitleText())
                    'System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
                    saveInfoToDWG()
                    saveRightInfoToDWG()
                    g_bIsClose = True
                    Me.Close()
                    'Exit Sub
                Case IDNO
                    g_bIsClose = True
                    Me.Close()
                    'Exit Sub
                Case IDCANCEL
                    'Exit Sub
            End Select
        Else
            g_bIsClose = True
            Me.Close()
        End If
    End Sub

    Private Function cmtoinches(ByRef flag As Object, ByRef cmCtrl As Object) As Object
        Dim realrem As Object
        Dim temprem, inch, rnder As Object
        Dim remainder As Double
        Dim temp As Short
        Dim Value As Double

        Dim cm As Object = DirectCast(cmCtrl, TextBox).Text

        'UPGRADE_WARNING: Couldn't resolve default property of object flag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If flag = "cm" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object inch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inch = Val(cm) / 2.54
            temp = Int(CDbl(inch))
            'UPGRADE_WARNING: Couldn't resolve default property of object inch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            remainder = (inch - temp)
            remainder = remainder / 0.125
            'UPGRADE_WARNING: Couldn't resolve default property of object realrem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            realrem = Int(remainder)
            'UPGRADE_WARNING: Couldn't resolve default property of object temprem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            temprem = remainder - Int(remainder)
            If temprem >= 0.5 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object realrem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                realrem = realrem + 1
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object realrem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            remainder = realrem * 0.125
            'UPGRADE_WARNING: Couldn't resolve default property of object inch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inch = temp + remainder
            'UPGRADE_WARNING: Couldn't resolve default property of object inch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            cmtoinches = inch
            'UPGRADE_WARNING: Couldn't resolve default property of object flag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf flag = "inches" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Value = cm
            ''Value = Val(cm)
            If Value >= 1 Then
                temp = Int(Value)
                remainder = Value - temp
                If remainder > 0.75 Then
                    MsgBox("Invalid Measurements, Inches and Eights only", 48, "Arm Details")
                    'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    cmtoinches = "-1"
                    Exit Function
                Else
                    remainder = remainder * 10
                    remainder = remainder * 0.125
                    'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    cmtoinches = temp + remainder
                End If
            Else
                remainder = Value
                If remainder > 0.75 Then
                    MsgBox("Invalid Measurements, Inches and Eights only", 48, "Arm Details")
                    'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    cmtoinches = "-1"
                    Exit Function
                Else
                    remainder = remainder * 10
                    remainder = remainder * 0.125
                    'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    cmtoinches = remainder
                End If

            End If
        End If

    End Function

    Private Sub Draw_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Draw.Click
        If txtSleeve.Text = "Right" Then
            RightDraw_Click()
            Exit Sub
        End If
        Dim nValue As Object
        Dim n2 As Object
        Dim n1 As Object
        Dim ff As Object
        Dim gg As Object
        ''txtSleeve.Text = "Left"
        'UPGRADE_WARNING: Arrays in structure Gauntend may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        'UPGRADE_WARNING: Arrays in structure Profile may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        'UPGRADE_WARNING: Arrays in structure conures may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Dim conures, Profile, Gauntend As New ARMDIA1.curve
        Dim tps, ii, nn, numtapes As Object
        Dim fNum As Short
        Dim xyPleat, xyThumbhole, xyFin, xyMidPoint, xyTHole1, xyPt, xyPt1, xyCen, xyBegin, xyTHole2, xyPt2, xyElbow1 As ARMDIA1.XY
        Dim nWrist, hole, pl, actfabric, Pleatsgn, GauntSize, nPalm, lenty As Object
        Dim m, nghk, nlen, nAng1, nAdj1, nOpp1, nOpp2, nAdj2, nAng2, midylen, n, valcheck As Object
        Dim xyThumb7, xyThumb5, xyThumb3, xyThumb1, xyThumbStart, xyThumb2, xyThumb4, xyThumb6, xyThumb8 As ARMDIA1.XY
        Dim nTextHt, thmAng1, thm2, thm1, thm3, thmang2, numtap As Object
        Dim xyP As ARMDIA1.XY
        Dim w, SeamDist, X As Object
        Dim v As Short
        Dim xyElbMark1, xyElbMark2 As ARMDIA1.XY
        Dim iPleats As Short
        ''Draw.Enabled = False
        'UPGRADE_WARNING: Couldn't resolve default property of object SeamDist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        SeamDist = 0.1875
        'UPGRADE_WARNING: Couldn't resolve default property of object NoVals(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object valcheck. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        valcheck = NoVals()


        If cboPressure.Text = "" Then
            MsgBox("Pressure has not been selected", 48, "Arm Details")
            Exit Sub
        End If
        If cboFabric.Text = "" Then
            MsgBox("Fabric has not been selected", 48, "Arm Details")
            Exit Sub
        End If

        PR_GetStartEndTapes()
        'Set style first and last tape
        If ARMDIA1.g_iStyleFirstTape = -1 Then ARMDIA1.g_iStyleFirstTape = Val(txtFirstTape.Text)
        If ARMDIA1.g_iStyleLastTape = -1 Then ARMDIA1.g_iStyleLastTape = Val(txtLastTape.Text)


        'UPGRADE_WARNING: Couldn't resolve default property of object valcheck. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If valcheck = 0 Then
            MsgBox("No tape measurements have been entered", 48, "Arm Details")
            Exit Sub
        End If

        If ARMDIA1.g_iStyleLastTape < ARMDIA1.g_iStyleFirstTape Then
            MsgBox("Chosen distal tape is above Proximal tape!", 48, "Arm Details")
            Exit Sub
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object numtapes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        numtapes = (Val(txtLastTape.Text) - Val(txtFirstTape.Text)) + 1

        'UPGRADE_WARNING: Couldn't resolve default property of object numtapes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If numtapes = 1 Then
            MsgBox("At least two tapes have to be entered before drawing", 48, "Arm Details")
            Exit Sub
        End If


        'Check if flap is style D and that a waist circumference is given
        If InStr(1, cboFlaps.Text, "D", 0) > 0 And Val(txtWaistCir.Text) = 0 Then
            MsgBox("A Waist circumference must be given for a D-Style flap", 48, "Arm Details")
            Exit Sub
        End If

        ' Check if tape has been deleted
        Dim tmpFirst, DeleteChecker, offnumtapes, actnumtapes, tmpLast As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object offnumtapes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        offnumtapes = (CDbl(txtLastTape.Text) - CDbl(txtFirstTape.Text)) + 1
        'UPGRADE_WARNING: Couldn't resolve default property of object gg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gg = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object tmpFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tmpFirst = ""
        'UPGRADE_WARNING: Couldn't resolve default property of object gg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While gg < 18
            If Length(gg).Text <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object gg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tmpFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                tmpFirst = gg + 1
                Exit Do
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object gg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            gg = gg + 1
        Loop

        ' secondlast & last tapes
        'UPGRADE_WARNING: Couldn't resolve default property of object ff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ff = 17
        'UPGRADE_WARNING: Couldn't resolve default property of object tmpLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tmpLast = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object ff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While ff > 0
            If Length(ff).Text <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object ff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tmpLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                tmpLast = ff + 1

                Exit Do
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object ff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ff = ff - 1
        Loop
        'UPGRADE_WARNING: Couldn't resolve default property of object tmpFirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object tmpLast. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object actnumtapes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        actnumtapes = (tmpLast - tmpFirst) + 1

        If chkStump.Checked = False Then
            txtStump.Text = CStr(False)
        End If

        PR_GetPressure()
        PR_GetFabric()
        PR_CheckValues()
        PR_GetGauntletDetails()
        PR_GauntletHoleCheck()
        PR_GetStartEndTapes()
        '    PR_CalculateMMSGmsReds 1
        PR_SetRresults()
        PR_SetGlobals()

        'check Gauntlet details
        If chkGauntlets.Checked = True Then
            If cboPalm.SelectedIndex = -1 Then
                MsgBox("No value for palm tape", 48, "Gauntlet Details")
                Exit Sub
            End If
            If cboWrist.SelectedIndex = -1 Then
                MsgBox("No value for wrist tape", 48, "Gauntlet Details")
                Exit Sub
            End If
        End If

        ' set tapewidths
        For w = 0 To 17
            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            X = w + 1
            PR_SetDistances(w, X)
            PR_CheckFlaps(w)
        Next w

        'UPGRADE_WARNING: Couldn't resolve default property of object numtapes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Profile.n = numtapes
        'UPGRADE_WARNING: Couldn't resolve default property of object numtapes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If numtapes > 3 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ii = 1
            ' Profile.X = Nothing
            'Profile.y = Nothing
            Dim X1(Val(txtLastTape.Text) - Val(txtFirstTape.Text) + 1) As Double
            Dim Y1(Val(txtLastTape.Text) - Val(txtFirstTape.Text) + 1) As Double
            For tps = Val(txtFirstTape.Text) To Val(txtLastTape.Text)
                'UPGRADE_WARNING: Couldn't resolve default property of object tps. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(tps). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Dim Measurement As Double = Val(Measurements(tps))
                If Measurement <= 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object tps. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tps = tps + 1
                    Profile.n = Profile.n - 1
                    Measurement = Val(Measurements(tps))
                End If
                actfabric = ((Val(Measurement) * (100 - Val(reds(tps - 1).Text))) / 100) / 2
                ' Profile.X(ii) = Val(tapewidths(tps))
                X1(ii) = Val(tapewidths(tps))
                'Profile.y(ii) = actfabric + SeamDist
                Y1(ii) = actfabric + SeamDist
                ii = ii + 1
            Next tps
            Profile.X = X1
            Profile.y = Y1
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object numtapes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If numtapes = 3 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object tps. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            tps = Val(txtFirstTape.Text)
            Dim X1(3), Y1(3) As Double
            actfabric = ((Val(Measurements(tps)) * (100 - Val(reds(tps - 1).Text))) / 100) / 2
            ''--------Profile.X(1) = Val(tapewidths(tps))
            X1(1) = Val(tapewidths(tps))
            ''------Profile.y(1) = actfabric + SeamDist
            Y1(1) = actfabric + SeamDist
            tps = tps + 2
            actfabric = ((Val(Measurements(tps)) * (100 - Val(reds(tps - 1).Text))) / 100) / 2
            ''-----------Profile.X(3) = Val(tapewidths(tps))
            X1(3) = Val(tapewidths(tps))
            ''----------Profile.y(3) = actfabric + SeamDist
            Y1(3) = actfabric + SeamDist
            tps = tps - 1
            If Measurements(tps) <> 0 Then 'And cboFlaps.Text = ""
                actfabric = ((Val(Measurements(tps)) * (100 - Val(reds(tps - 1).Text))) / 100) / 2
                ''---------Profile.X(2) = Val(tapewidths(tps))
                ''---------Profile.y(2) = actfabric + SeamDist
                X1(2) = Val(tapewidths(tps))
                Y1(2) = actfabric + SeamDist
            End If
            Profile.X = X1
            Profile.y = Y1
        End If
        If numtapes = 3 Then
            tps = Val(txtFirstTape.Text)
            If Str(Measurements(tps + 1)) = "" And chkGauntlets.Checked = True Then
                Profile.n = 2
                actfabric = ((Val(Measurements(tps)) * (100 - Val(reds(tps - 1).Text))) / 100) / 2
                Profile.X(1) = Val(tapewidths(tps))
                Profile.y(1) = actfabric + SeamDist
                tps = tps + 2
                actfabric = ((Val(Measurements(tps)) * (100 - Val(reds(tps - 1).Text))) / 100) / 2
                Profile.X(2) = Profile.X(1) + Val(txtPalmWristDist.Text)
                Profile.y(2) = actfabric + SeamDist
            End If
        End If
        If numtapes = 2 And chkGauntlets.Checked = False Then
            ii = 1
            Dim X1(2), Y1(2) As Double
            Profile.X = X1
            Profile.y = Y1
            For tps = Val(txtFirstTape.Text) To Val(txtLastTape.Text)
                If Measurements(tps) <= 0 Then
                    tps = tps + 1
                    Profile.n = Profile.n - 1
                End If
                actfabric = ((Val(Measurements(tps)) * (100 - Val(reds(tps - 1).Text))) / 100) / 2
                Profile.X(ii) = Val(tapewidths(tps))
                Profile.y(ii) = actfabric + SeamDist
                ii = ii + 1
            Next tps
        End If

        'Check that there are more tape spaces than pleats
        'I.E. So that pleats will fit
        iPleats = 0
        If Val(txtWristPleat1.Text) > 0 And txtWristPleat1.Enabled Then iPleats = iPleats + 1
        If Val(txtWristPleat2.Text) > 0 And txtWristPleat2.Enabled Then iPleats = iPleats + 1
        If Val(txtShoulderPleat1.Text) > 0 And txtShoulderPleat1.Enabled Then iPleats = iPleats + 1
        If Val(txtShoulderPleat2.Text) > 0 And txtShoulderPleat2.Enabled Then iPleats = iPleats + 1
        'As we can't have a pleat between palm and wrist we fudge it here to cut down on code.
        If chkGauntlets.Checked = True Then iPleats = iPleats + 1
        'If Profile.n <= iPleats Then
        '    MsgBox("There are more pleats given, than spaces available between tapes. Remove or temporarily disable pleats, by double clicking on the ""Pleat"" label!", 48, "Arm Details")
        '    Exit Sub
        'End If

        ''ARMDIA1.fNum = ARMDIA1.FN_Open("c:\jobst\draw.d", txtPatientName.Text, txtFileNo.Text, txtSleeve.Text, "Draw")
        Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
        ARMDIA1.fNum = ARMDIA1.FN_Open(sDrawFile, txtPatientName.Text, txtFileNo.Text, txtSleeve.Text, "Draw")
        FileClose(ARMDIA1.fNum)
        Me.Hide()
        PR_GetInsertionPoint()
        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
        ARMDIA1.PR_DrawFitted(Profile, True)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "Profile") 'g_sID from FN_Open
        ' PrintLine(ARMDIA1.fNum, "SetDBData(hEnt," & ARMDIA1.QQ & "curvetype" & ARMDIA1.QCQ & "sleeveprofile" & ARMDIA1.QQ & ");")


        ARMDIA1.PR_NamedHandle("hSleeveProfile")


        ARMDIA1.PR_MakeXY(xyPt1, Profile.X(1), 0)
        ARMDIA1.PR_MakeXY(xyPt2, Profile.X(Profile.n), 0)
        ARMDIA1.PR_DrawLine(xyPt1, xyPt2) 'Fold Line
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "FoldLine") 'sFileNo, sSide from FN_Open

        ARMDIA1.PR_SetLayer("Notes")
        PR_DrawLineOffset(xyPt1, xyPt2, 0.6875) 'Inner tram line
        PR_DrawLineOffset(xyPt1, xyPt2, 0.1875) 'seam line at 3/16ths

        If txtGauntletFlag.Text = "False" Then
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
            ARMDIA1.PR_MakeXY(xyPt1, Profile.X(1), 0)
            If Val(txtFirstTape.Text) = 11 Then
                ARMDIA1.PR_MakeXY(xyPt2, Profile.X(1), Profile.y(1) / 2)
            Else
                ARMDIA1.PR_MakeXY(xyPt2, Profile.X(1), Profile.y(1))
            End If
            ARMDIA1.PR_DrawLine(xyPt1, xyPt2) 'Wrist Line
            ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "WristLine") 'sFileNo, sSide from FN_Open


        End If

        If cboFlaps.Text = "" Or cboFlaps.Text = " " Then
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
            If Val(txtLastTape.Text) = 11 Then
                ARMDIA1.PR_MakeXY(xyPt1, Profile.X(Profile.n), Profile.y(Profile.n) / 2)
            Else
                ARMDIA1.PR_MakeXY(xyPt1, Profile.X(Profile.n), Profile.y(Profile.n))
            End If
            ARMDIA1.PR_MakeXY(xyPt2, Profile.X(Profile.n), 0)
            ARMDIA1.PR_DrawLine(xyPt1, xyPt2) 'Shoulder Line
            ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ShoulderLine") 'sFileNo, sSide from FN_Open

        End If

        ' if Old Drawing add DB Fields
        If g_sOld = "True" Then
            '' PR_OldDrawing()
        End If

        nTextHt = 0.125

        ' Draw Circular Stump
        If txtStump.Text = "True" Then
            ARMDIA1.PR_MakeXY(xyPt1, Profile.X(1) - 5, Profile.y(1) / 2)
            PR_DrawCircularStump(xyPt1, Profile.y(1), txtAge.Text)
            'Add Stump text
            ARMDIA1.PR_SetLayer("Notes")
            ARMDIA1.g_nCurrTextAngle = 90
            ARMDIA1.PR_MakeXY(xyPt1, Profile.X(1) + 0.25, Profile.y(1) / 2)
            ARMDIA1.PR_DrawText("STUMP", xyPt1, nTextHt, 90 * (ARMDIA1.PI / 180))
            ARMDIA1.g_nCurrTextAngle = 0
        End If
        ' Insert Elbow mark
        ii = 1
        For tps = Val(txtFirstTape.Text) To Val(txtLastTape.Text)
            'Elbow Mark
            If tps = 11 And g_sHoleCheck <> "True" Then
                ARMDIA1.PR_SetLayer("Notes")
                ARMDIA1.PR_MakeXY(xyElbMark1, Profile.X(ii), Profile.y(ii) - 0.4)
                ARMDIA1.PR_MakeXY(xyElbMark2, Profile.X(ii), Profile.y(ii) + 0.4)
                ARMDIA1.PR_DrawLine(xyElbMark1, xyElbMark2)
                PR_DrawAssignDrafixVariable("xyElbow.x", Profile.X(ii))
                PR_DrawAssignDrafixVariable("xyElbow.y", Profile.y(ii))
            End If

            If tps = 10 And g_sHoleCheck = "True" Then
                ARMDIA1.PR_SetLayer("Notes")
                ARMDIA1.PR_MakeXY(xyElbMark1, Profile.X(ii), Profile.y(ii) - 0.4)
                ARMDIA1.PR_MakeXY(xyElbMark2, Profile.X(ii), Profile.y(ii) + 0.4)
                ARMDIA1.PR_DrawLine(xyElbMark1, xyElbMark2)
                PR_DrawAssignDrafixVariable("xyElbow.x", Profile.X(ii))
                PR_DrawAssignDrafixVariable("xyElbow.y", Profile.y(ii))
            End If
            ii = ii + 1
        Next tps


        ' insert Pleats
        If Val(txtWristPleat1.Text) > 0 And txtWristPleat1.Enabled Then
            nTextHt = 0.125
            ARMDIA1.PR_SetLayer("Construct")
            If chkGauntlets.Checked = True Then
                n1 = ARMDIA1.max(Val(txtWristNo.Text) - Val(txtFirstTape.Text), 2)
                n2 = n1 + 1
            Else
                n1 = 1
                n2 = 2
            End If
            pl = (Profile.X(n2) - Profile.X(n1)) / 2
            ARMDIA1.PR_MakeXY(xyPleat, Profile.X(n1) + pl, 0.3375)
            ARMDIA1.PR_DrawText("Pleat", xyPleat, nTextHt, 0)
        End If
        If Val(txtWristPleat2.Text) > 0 And txtWristPleat2.Enabled Then
            nTextHt = 0.125
            ARMDIA1.PR_SetLayer("Construct")
            If chkGauntlets.Checked = True Then
                n1 = ARMDIA1.max(Val(txtWristNo.Text) - Val(txtFirstTape.Text), 2) + 1
                n2 = n1 + 1
            Else
                n1 = 2
                n2 = 3
            End If
            pl = (Profile.X(n2) - Profile.X(n1)) / 2
            ARMDIA1.PR_MakeXY(xyPleat, Profile.X(n1) + pl, 0.3375)
            ARMDIA1.PR_DrawText("Pleat", xyPleat, nTextHt, 0)
        End If
        If Val(txtShoulderPleat1.Text) > 0 And txtShoulderPleat1.Enabled Then
            nTextHt = 0.125
            ARMDIA1.PR_SetLayer("Construct")
            n = Profile.n
            m = Profile.n - 1
            pl = (Profile.X(n) - Profile.X(m)) / 2
            ARMDIA1.PR_MakeXY(xyPleat, Profile.X(m) + pl, 0.3375)
            ARMDIA1.PR_DrawText("Pleat", xyPleat, nTextHt, 0)
        End If
        If Val(txtShoulderPleat2.Text) > 0 And txtShoulderPleat2.Enabled Then
            nTextHt = 0.125
            ARMDIA1.PR_SetLayer("Construct")
            n = Profile.n - 1
            m = Profile.n - 2
            pl = (Profile.X(n) - Profile.X(m)) / 2
            ARMDIA1.PR_MakeXY(xyPleat, Profile.X(m) + pl, 0.3375)
            ARMDIA1.PR_DrawText("Pleat", xyPleat, nTextHt, 0)
        End If

        'Draw Gauntlet
        If chkGauntlets.Checked = True Then
            DrawGauntlet(Profile, numtapes)
        End If

        'Patient Details
        TypeDetails(Profile, numtapes)

        ' Elastics
        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)

        Elastic(Profile)

        ' Shoulder Flaps
        If cboFlaps.Text = "" Or VB.Left(cboFlaps.Text, 4) = "None" Then
            g_sFlap = "None"
        ElseIf (VB.Left(cboFlaps.Text, 4) = "Vest") Then
            g_sFlap = ""
            g_sFlapStrap = ""
            g_sFlap = cboFlaps.Text
            PR_DrawRaglan("Vest", txtVestRaglan.Text, txtAge.Text, txtVestID)

        ElseIf (VB.Left(cboFlaps.Text, 4) = "Body") Then
            g_sFlap = ""
            g_sFlapStrap = ""
            g_sFlap = cboFlaps.Text
            PR_DrawRaglan("Body", txtVestRaglan.Text, txtAge.Text, txtVestID)

        Else
            g_sFlap = cboFlaps.Text

            PR_DrawShoulderFlaps(Profile, numtapes)
        End If

        ARMDIA1.g_sFabric = cboFabric.Text

        'Vestarm and Profile Origin Marker data base update
        'For editing purposes the tape lengths are stored with
        'the Profile Origin Marker
        Dim sJustifiedString As String
        For ii = 0 To 17
            nValue = Val(Length(ii).Text)
            If nValue <> 0 And ii >= (Val(txtFirstTape.Text) - 1) And ii <= (Val(txtLastTape.Text) - 1) Then
                nValue = nValue * 10 'Shift decimal place to right
                sJustifiedString = New String(" ", 3)
                sJustifiedString = RSet(Trim(Str(nValue)), Len(sJustifiedString))
            Else
                sJustifiedString = New String(" ", 3)
            End If

            'Tape values
            g_sEditLengths = g_sEditLengths & sJustifiedString
        Next ii

        'Update
        DataBaseDataUpDate("Draw")
        ''Added for #166 in the issue list
        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
        ii = 1
        For tps = Val(txtFirstTape.Text) To Val(txtLastTape.Text)
            Dim Measurement As Double = Val(Measurements(tps))
            If Measurement <= 0 Then
                tps = tps + 1
                Measurement = Val(Measurements(tps))
            End If
            ARMDIA1.PR_MakeXY(xyPt1, Profile.X(ii), 0)
            '---------ARMDIA1.PR_PutTapeLabel(tps, xyPt1, Measurement, mms(tps - 1), gms(tps - 1), reds(tps - 1))
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                ''ARMDIA1.PR_PutTapeLabel(tps, xyPt1, fnDisplayToCM(Measurement), mms(tps - 1), gms(tps - 1), reds(tps - 1))
                ARMDIA1.PR_PutTapeLabel(tps, xyPt1, Val(Length(tps - 1).Text), mms(tps - 1), gms(tps - 1), reds(tps - 1))
            Else
                ARMDIA1.PR_PutTapeLabel(tps, xyPt1, Measurement, mms(tps - 1), gms(tps - 1), reds(tps - 1))
            End If
            PR_DrawRuler(tps, xyPt1, lblTape(tps - 1).Text)
            ii = ii + 1
        Next tps

        'For Body suit et.al. start program to draw the mesh
        PR_DrawMesh()
        saveInfoToDWG()
        g_bIsClose = True
        Me.Close()
        Draw.Enabled = True
        FileClose(ARMDIA1.fNum)
        '' AppActivate(fnGetDrafixWindowTitleText())
        '' System.Windows.Forms.SendKeys.SendWait("@c:\jobst\draw.d{enter}")
        PR_SetLayer("Construct")
        PR_DrawXMarker()
        VestMain.VestMainDlg.Hide()
        VestMain.VestMainDlg.Close()
        Return

    End Sub

    Private Function readDWGInfo()
        Try

            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
            resbuf = _sClass.GetXrecord()
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                _Length_0.Text = arr(0).Value
                Length_Leave(_Length_0, Nothing)
                _Length_1.Text = arr(1).Value
                Length_Leave(_Length_1, Nothing)
                _Length_2.Text = arr(2).Value
                Length_Leave(_Length_2, Nothing)
                _Length_3.Text = arr(3).Value
                _Length_4.Text = arr(4).Value
                _Length_5.Text = arr(5).Value
                _Length_6.Text = arr(6).Value
                _Length_7.Text = arr(7).Value
                _Length_8.Text = arr(8).Value
                _Length_9.Text = arr(9).Value
                _Length_10.Text = arr(10).Value
                _Length_11.Text = arr(11).Value
                _Length_12.Text = arr(12).Value
                _Length_13.Text = arr(13).Value
                _Length_14.Text = arr(14).Value
                _Length_15.Text = arr(15).Value
                _Length_16.Text = arr(16).Value
                _Length_17.Text = arr(17).Value
                cboFabric.Text = arr(18).Value
                txtMM.Text = arr(19).Value

                _mms_0.Text = arr(20).Value
                _mms_1.Text = arr(21).Value
                _mms_2.Text = arr(22).Value
                _mms_3.Text = arr(23).Value
                _mms_4.Text = arr(24).Value
                _mms_5.Text = arr(25).Value
                _mms_6.Text = arr(26).Value
                _mms_7.Text = arr(27).Value
                _mms_8.Text = arr(28).Value
                _mms_9.Text = arr(29).Value
                _mms_10.Text = arr(30).Value
                _mms_11.Text = arr(31).Value
                _mms_12.Text = arr(32).Value
                _mms_13.Text = arr(33).Value
                _mms_14.Text = arr(34).Value
                _mms_15.Text = arr(35).Value
                _mms_16.Text = arr(36).Value
                _mms_17.Text = arr(37).Value

                _gms_0.Text = arr(38).Value
                _gms_1.Text = arr(39).Value
                _gms_2.Text = arr(40).Value
                _gms_3.Text = arr(41).Value
                _gms_4.Text = arr(42).Value
                _gms_5.Text = arr(43).Value
                _gms_6.Text = arr(44).Value
                _gms_7.Text = arr(45).Value
                _gms_8.Text = arr(46).Value
                _gms_9.Text = arr(47).Value
                _gms_10.Text = arr(48).Value
                _gms_11.Text = arr(49).Value
                _gms_12.Text = arr(50).Value
                _gms_13.Text = arr(51).Value
                _gms_14.Text = arr(52).Value
                _gms_15.Text = arr(53).Value
                _gms_16.Text = arr(54).Value
                _gms_17.Text = arr(55).Value


                _reds_0.Text = arr(56).Value
                _reds_1.Text = arr(57).Value
                _reds_2.Text = arr(58).Value
                _reds_3.Text = arr(59).Value
                _reds_4.Text = arr(60).Value
                _reds_5.Text = arr(61).Value
                _reds_6.Text = arr(62).Value
                _reds_7.Text = arr(63).Value
                _reds_8.Text = arr(64).Value
                _reds_9.Text = arr(65).Value
                _reds_10.Text = arr(66).Value
                _reds_11.Text = arr(67).Value
                _reds_12.Text = arr(68).Value
                _reds_13.Text = arr(69).Value
                _reds_14.Text = arr(70).Value
                _reds_15.Text = arr(71).Value
                _reds_16.Text = arr(72).Value
                _reds_17.Text = arr(73).Value

                txtStump.Text = arr(74).Value
                txtGauntlet.Text = arr(75).Value
                optProximalTape(0).Checked = arr(76).Value
                optProximalTape(1).Checked = arr(77).Value

                cboDistalTape.Text = arr(78).Value
                cboProximalTape.Text = arr(79).Value
                txtFlap.Text = arr(80).Value
                txtWristPleat1.Text = arr(81).Value
                txtWristPleat2.Text = arr(82).Value
                txtShoulderPleat1.Text = arr(83).Value
                txtShoulderPleat2.Text = arr(84).Value

                Length_Leave(_Length_3, Nothing)
                Length_Leave(_Length_4, Nothing)
                Length_Leave(_Length_5, Nothing)
                Length_Leave(_Length_6, Nothing)
                Length_Leave(_Length_7, Nothing)
                Length_Leave(_Length_8, Nothing)
                Length_Leave(_Length_9, Nothing)
                Length_Leave(_Length_10, Nothing)
                Length_Leave(_Length_11, Nothing)
                Length_Leave(_Length_12, Nothing)
                Length_Leave(_Length_13, Nothing)
                Length_Leave(_Length_14, Nothing)
                Length_Leave(_Length_15, Nothing)
                Length_Leave(_Length_16, Nothing)
                Length_Leave(_Length_17, Nothing)

                cboFabric_SelectedIndexChanged(Nothing, Nothing)
                Calculate_Click(Nothing, Nothing)
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
                txtAge1.Text = totalYears.ToString()
            End If

        Catch ex As Exception

        End Try
    End Function

    Private Function saveInfoToDWG()
        Try
            Dim _sClass As New SurroundingClass()
            If (_sClass.GetXrecord() IsNot Nothing) Then
                _sClass.RemoveXrecord()
            End If

            Dim resbuf As New ResultBuffer

            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_0.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_1.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_2.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_3.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_4.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_5.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_6.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_7.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_8.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_9.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_10.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_11.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_12.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_13.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_14.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_15.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_16.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_17.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboFabric.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtMM.Text))

            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_0.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_1.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_2.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_3.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_4.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_5.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_6.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_7.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_8.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_9.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_10.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_11.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_12.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_13.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_14.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_15.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_16.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_17.Text))

            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_0.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_1.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_2.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_3.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_4.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_5.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_6.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_7.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_8.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_9.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_10.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_11.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_12.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_13.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_14.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_15.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_16.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_17.Text))

            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_0.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_1.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_2.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_3.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_4.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_5.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_6.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_7.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_8.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_9.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_10.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_11.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_12.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_13.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_14.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_15.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_16.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_17.Text))

            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtStump.Text))
            Dim sGauntlet As String
            If chkGauntlets.Checked = True Then
                sGauntlet = "1"
            Else
                sGauntlet = "0"
            End If
            If chkEnclosed.Checked = True Then
                sGauntlet = sGauntlet & " 1"
            Else
                sGauntlet = sGauntlet & " 0"
            End If
            If chkDetachable.Checked = True Then
                sGauntlet = sGauntlet & " 1"
            Else
                sGauntlet = sGauntlet & " 0"
            End If
            If chkNoThumb.Checked = True Then
                sGauntlet = sGauntlet & " 1"
            Else
                sGauntlet = sGauntlet & " 0"
            End If
            sGauntlet = sGauntlet & " " & Str(cboWrist.SelectedIndex + 1) & " " & Str(cboPalm.SelectedIndex + 1) & " " & txtThumb.Text & " " & txtCircum.Text & " " & txtPalmWrist.Text & " " & txtGauntletExtension.Text
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), sGauntlet))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optProximalTape(0).Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optProximalTape(1).Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboDistalTape.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), cboProximalTape.Text))
            Dim sFlapDetail As String = ""
            If optProximalTape(1).Checked = True Then
                Dim sWaistCir As String = txtWaistCir.Text
                If txtWaistCir.Text = "" Then
                    sWaistCir = "0"
                End If
                ''Changed for #168 in the issue list
                ''sFlapDetail = txtStrapLength.Text & " " & txtCustFlapLength.Text & " " & sWaistCir & " " & txtFrontStrapLength.Text & " " & Str(cboFlaps.SelectedIndex)
                Dim sStrapLength As String = txtStrapLength.Text
                If txtStrapLength.Text = "" Then
                    sStrapLength = "0"
                End If
                Dim sCustFlapLength As String = txtCustFlapLength.Text
                If txtCustFlapLength.Text = "" Then
                    sCustFlapLength = "0"
                End If
                Dim sFrontStrapLength As String = txtFrontStrapLength.Text
                If txtFrontStrapLength.Text = "" Then
                    sFrontStrapLength = "0"
                End If
                sFlapDetail = sStrapLength & " " & sCustFlapLength & " " & sWaistCir & " " & sFrontStrapLength & " " & Str(cboFlaps.SelectedIndex)
            End If
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), sFlapDetail))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtWristPleat1.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtWristPleat2.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtShoulderPleat1.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), txtShoulderPleat2.Text))


            _sClass.SetXrecord(resbuf)

        Catch ex As Exception

        End Try
    End Function

    Private Sub DrawGauntlet(ByRef Profile As ARMDIA1.curve, ByRef numtapes As Object)
        Dim Ydiff As Object
        'UPGRADE_WARNING: Arrays in structure GauntDetachEnd may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        'UPGRADE_WARNING: Arrays in structure Gauntend may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Dim Gauntend, GauntDetachEnd As ARMDIA1.curve
        Dim xyCen, xyPt1, xyTHole1, xyMidPoint As ARMDIA1.XY
        Dim xyThumbConstruct1, xyThumbConstruct2 As ARMDIA1.XY
        Dim lenty, hole, GauntSize, nPalm, nTextHt As Object
        Dim num, nlen, nAdj2, nOpp2, nOpp1, nAdj1, nAng1, midylen, Ext As Object
        Dim xyThumbhole, xyFin, xyBegin, xyTHole2, xyPt2 As ARMDIA1.XY
        Dim xyGend1, xyGend2 As ARMDIA1.XY
        Dim tps, ii, nn, fNum As Object
        Dim xyPt, xyPleat As ARMDIA1.XY
        Dim xyElbow3, xyElbow1, xyElbow2, xyElbow4 As ARMDIA1.XY
        Dim nAng2, Pleatsgn, actfabric, pl, nghk As Object
        Dim xyP As ARMDIA1.XY
        Dim xyThumb7, xyThumb5, xyThumb3, xyThumb1, xyThumbStart, xyThumb2, xyThumb4, xyThumb6, xyThumb8 As ARMDIA1.XY
        Dim thmAng1, thm2, thm1, thm3, thmang2 As Object
        Dim xyThumbNotes As ARMDIA1.XY

        'Calaculate Gauntlet extension
        GauntSize = 0
        If CDbl(txtAge.Text) > 14 And txtSex.Text = "Male" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object GauntSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            GauntSize = 1.125
        ElseIf CDbl(txtAge.Text) > 14 And txtSex.Text = "Female" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object GauntSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            GauntSize = 0.625
        ElseIf CDbl(txtAge.Text) < 15 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object GauntSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            GauntSize = 0.5
        End If
        'Use given extension if one is given
        'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object GauntSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(txtGauntletExtension.Text) > 0 Then GauntSize = cmtoinches(txtinchflag.Text, txtGauntletExtension)

        'UPGRADE_WARNING: Couldn't resolve default property of object num. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        num = (Val(txtLastTape.Text) - Val(txtFirstTape.Text)) + 2

        'Draw gauntlet extension
        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
        Gauntend.n = 4
        Dim nX(4), nY(4) As Double
        'Gauntend.X(1) = Profile.X(1)
        'Gauntend.y(1) = 0
        ''UPGRADE_WARNING: Couldn't resolve default property of object GauntSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Gauntend.X(2) = Profile.X(1) - GauntSize
        'Gauntend.y(2) = 0
        ''UPGRADE_WARNING: Couldn't resolve default property of object GauntSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Gauntend.X(3) = Profile.X(1) - GauntSize
        'Gauntend.y(3) = Profile.y(1)
        'Gauntend.X(4) = Profile.X(1)
        'Gauntend.y(4) = Profile.y(1)

        nX(1) = Profile.X(1)
        nY(1) = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object GauntSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nX(2) = Profile.X(1) - GauntSize
        nY(2) = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object GauntSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nX(3) = Profile.X(1) - GauntSize
        nY(3) = Profile.y(1)
        nX(4) = Profile.X(1)
        nY(4) = Profile.y(1)
        Gauntend.X = nX
        Gauntend.y = nY

        ARMDIA1.PR_DrawPoly(Gauntend)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "GauntEnd") 'sFileNo, sSide from FN_Open


        ' start of added lines
        ARMDIA1.PR_MakeXY(xyPt1, Gauntend.X(2), 0)

        ARMDIA1.PR_MakeXY(xyPt2, Gauntend.X(1), 0)

        ARMDIA1.PR_DrawLine(xyPt1, xyPt2) 'Fold Line
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "FoldLine") 'sFileNo, sSide from FN_Open


        ARMDIA1.PR_SetLayer("Notes")
        PR_DrawLineOffset(xyPt1, xyPt2, 0.6875) 'Inner tram line
        PR_DrawLineOffset(xyPt1, xyPt2, 0.1875) 'seam line at 3/16ths

        ' end of added lines


        ' Thumb opening
        'UPGRADE_WARNING: Couldn't resolve default property of object nPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nPalm = Val(txtPalmNo.Text) - Val(txtFirstTape.Text) + 1
        'UPGRADE_WARNING: Couldn't resolve default property of object nPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Ydiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Ydiff = Profile.y(nPalm + 1) - Profile.y(nPalm)
        'UPGRADE_WARNING: Couldn't resolve default property of object Ydiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If System.Math.Abs(Ydiff) > 0 Then 'Avoids overflow
            'UPGRADE_WARNING: Couldn't resolve default property of object Ydiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xyThumbhole.y = Profile.y(nPalm) + (Ydiff / 2)
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object nPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xyThumbhole.y = Profile.y(nPalm)
        End If
        'Draw thumb hole construction line
        ARMDIA1.PR_SetLayer("Construct")
        xyThumbConstruct1.y = xyThumbhole.y
        'UPGRADE_WARNING: Couldn't resolve default property of object nPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xyThumbConstruct1.X = Profile.X(nPalm) + (Val(txtPalmWristDist.Text) / 2)
        ARMDIA1.PR_MakeXY(xyThumbConstruct2, xyThumbConstruct1.X, 0.1875)
        ARMDIA1.PR_DrawLine(xyThumbConstruct1, xyThumbConstruct2)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHoleConstruct") 'g_sID from FN_Open

        PR_AddMarker(xyThumbConstruct1)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHoleConstruct") 'g_sID from FN_Open


        ' Lower line
        xyThumbhole.y = ((xyThumbhole.y - 0.1875) / 2) + 0.375
        'UPGRADE_WARNING: Couldn't resolve default property of object nPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xyThumbhole.X = Profile.X(nPalm)

        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
        ARMDIA1.PR_MakeXY(xyTHole1, xyThumbhole.X + 0.375, xyThumbhole.y - 0.1875)
        ARMDIA1.PR_MakeXY(xyTHole2, xyThumbhole.X + CDbl(txtThumbDist.Text) - 0.1875, xyThumbhole.y - 0.1875)
        ARMDIA1.PR_DrawLine(xyTHole1, xyTHole2)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHole") 'g_sID from FN_Open

        ' Upper Line
        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
        ARMDIA1.PR_MakeXY(xyTHole1, xyThumbhole.X + 0.375, xyThumbhole.y + 0.1875)
        ARMDIA1.PR_MakeXY(xyTHole2, xyThumbhole.X + CDbl(txtThumbDist.Text) - 0.1875, xyThumbhole.y + 0.1875)
        ARMDIA1.PR_DrawLine(xyTHole1, xyTHole2)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHole") 'g_sID from FN_Open

        ' Thumb Reminder in layer notes
        ARMDIA1.PR_SetLayer("Notes")
        ARMDIA1.PR_MakeXY(xyThumbNotes, xyThumbhole.X + 0.375, xyThumbhole.y + 0.5575)
        If chkNoThumb.CheckState = 0 Then
            ARMDIA1.PR_DrawText("Thumb Attached", xyThumbNotes, 0.125, 0)
        Else
            ARMDIA1.PR_DrawText("No Thumb", xyThumbNotes, 0.125, 0)
        End If

        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)

        ' Right Curve
        ARMDIA1.PR_MakeXY(xyCen, xyThumbhole.X + CDbl(txtThumbDist.Text) - 0.1875, xyThumbhole.y)
        ARMDIA1.PR_MakeXY(xyBegin, xyCen.X, xyCen.y + 0.1875)
        ARMDIA1.PR_MakeXY(xyFin, xyCen.X, xyCen.y - 0.1875)
        PR_RightThumbHole(xyCen, xyFin, xyBegin)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHole") 'g_sID from FN_Open

        ' Left top Curve
        ARMDIA1.PR_MakeXY(xyBegin, xyThumbhole.X + 0.375, xyThumbhole.y + 0.1875)
        ARMDIA1.PR_MakeXY(xyFin, xyThumbhole.X, xyThumbhole.y)
        'UPGRADE_WARNING: Couldn't resolve default property of object lenty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lenty = ARMDIA1.FN_CalcLength(xyBegin, xyFin)
        'UPGRADE_WARNING: Couldn't resolve default property of object lenty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        midylen = lenty / 2
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAng1 = ARMDIA1.FN_CalcAngle(xyFin, xyBegin)
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAng1 = nAng1 * (ARMDIA1.PI / 180)
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAdj1 = System.Math.Cos(nAng1) * midylen
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nOpp1 = System.Math.Sin(nAng1) * midylen
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMDIA1.PR_MakeXY(xyMidPoint, xyFin.X + nAdj1, xyFin.y + nOpp1)
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nlen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nlen = System.Math.Sqrt(0.25 - (midylen * midylen))
        'UPGRADE_WARNING: Couldn't resolve default property of object nlen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAdj2 = System.Math.Cos(nAng1) * nlen
        'UPGRADE_WARNING: Couldn't resolve default property of object nlen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nOpp2 = System.Math.Sin(nAng1) * nlen
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMDIA1.PR_MakeXY(xyCen, xyMidPoint.X + nOpp2, xyMidPoint.y - nAdj2)
        ARMDIA1.PR_DrawArc(xyCen, xyBegin, xyFin)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHole") 'g_sID from FN_Open

        ' Left Bottom Curve
        ARMDIA1.PR_MakeXY(xyBegin, xyThumbhole.X + 0.375, xyThumbhole.y - 0.1875)
        ARMDIA1.PR_MakeXY(xyFin, xyThumbhole.X, xyThumbhole.y)
        'UPGRADE_WARNING: Couldn't resolve default property of object lenty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lenty = ARMDIA1.FN_CalcLength(xyBegin, xyFin)
        'UPGRADE_WARNING: Couldn't resolve default property of object lenty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        midylen = lenty / 2
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAng1 = ARMDIA1.FN_CalcAngle(xyFin, xyBegin)
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAng1 = nAng1 * (ARMDIA1.PI / 180)
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAdj1 = System.Math.Cos(nAng1) * midylen
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nOpp1 = System.Math.Sin(nAng1) * midylen
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMDIA1.PR_MakeXY(xyMidPoint, xyFin.X + nAdj1, xyFin.y + nOpp1)
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nlen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nlen = System.Math.Sqrt(0.25 - (midylen * midylen))
        'UPGRADE_WARNING: Couldn't resolve default property of object nlen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAdj2 = System.Math.Cos(nAng1) * nlen
        'UPGRADE_WARNING: Couldn't resolve default property of object nlen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nOpp2 = System.Math.Sin(nAng1) * nlen
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMDIA1.PR_MakeXY(xyCen, xyMidPoint.X - nOpp2, xyMidPoint.y + nAdj2)
        ARMDIA1.PR_DrawArc(xyCen, xyFin, xyBegin)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHole") 'g_sID from FN_Open

        'Cutting line to Distal EOS
        'UPGRADE_WARNING: Couldn't resolve default property of object GauntSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMDIA1.PR_MakeXY(xyBegin, xyThumbhole.X - GauntSize, xyThumbhole.y)
        ARMDIA1.PR_DrawLine(xyBegin, xyThumbhole)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHole") 'g_sID from FN_Open

        If chkNoThumb.CheckState = 0 Then
            'Draw Thumb
            DrawThumb(xyThumbhole)
        End If
        If chkDetachable.CheckState = 1 Then
        End If

    End Sub
    Private Sub DrawRightGauntlet(ByRef Profile As ARMDIA1.curve, ByRef numtapes As Object)
        Dim Ydiff As Object
        Dim Gauntend, GauntDetachEnd As ARMDIA1.curve
        Dim xyCen, xyPt1, xyTHole1, xyMidPoint As ARMDIA1.XY
        Dim xyThumbConstruct1, xyThumbConstruct2 As ARMDIA1.XY
        Dim lenty, hole, GauntSize, nPalm, nTextHt As Object
        Dim num, nlen, nAdj2, nOpp2, nOpp1, nAdj1, nAng1, midylen, Ext As Object
        Dim xyThumbhole, xyFin, xyBegin, xyTHole2, xyPt2 As ARMDIA1.XY
        Dim xyGend1, xyGend2 As ARMDIA1.XY
        Dim tps, ii, nn, fNum As Object
        Dim xyPt, xyPleat As ARMDIA1.XY
        Dim xyElbow3, xyElbow1, xyElbow2, xyElbow4 As ARMDIA1.XY
        Dim nAng2, Pleatsgn, actfabric, pl, nghk As Object
        Dim xyP As ARMDIA1.XY
        Dim xyThumb7, xyThumb5, xyThumb3, xyThumb1, xyThumbStart, xyThumb2, xyThumb4, xyThumb6, xyThumb8 As ARMDIA1.XY
        Dim thmAng1, thm2, thm1, thm3, thmang2 As Object
        Dim xyThumbNotes As ARMDIA1.XY

        'Calaculate Gauntlet extension
        GauntSize = 0
        If CDbl(txtAge.Text) > 14 And txtSex.Text = "Male" Then
            GauntSize = 1.125
        ElseIf CDbl(txtAge.Text) > 14 And txtSex.Text = "Female" Then
            GauntSize = 0.625
        ElseIf CDbl(txtAge.Text) < 15 Then
            GauntSize = 0.5
        End If
        'Use given extension if one is given
        If Val(RighttxtGauntletExtension.Text) > 0 Then GauntSize = cmtoinches(Righttxtinchflag.Text, RighttxtGauntletExtension)
        num = (Val(strLastTape) - Val(strFirstTape)) + 2

        'Draw gauntlet extension
        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
        Gauntend.n = 4
        Dim nX(4), nY(4) As Double

        nX(1) = Profile.X(1)
        nY(1) = 0
        nX(2) = Profile.X(1) - GauntSize
        nY(2) = 0
        nX(3) = Profile.X(1) - GauntSize
        nY(3) = Profile.y(1)
        nX(4) = Profile.X(1)
        nY(4) = Profile.y(1)
        Gauntend.X = nX
        Gauntend.y = nY
        ARMDIA1.PR_DrawPoly(Gauntend)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "GauntEnd") 'sFileNo, sSide from FN_Open

        ' start of added lines
        ARMDIA1.PR_MakeXY(xyPt1, Gauntend.X(2), 0)
        ARMDIA1.PR_MakeXY(xyPt2, Gauntend.X(1), 0)
        ARMDIA1.PR_DrawLine(xyPt1, xyPt2) 'Fold Line
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "FoldLine") 'sFileNo, sSide from FN_Open

        ARMDIA1.PR_SetLayer("Notes")
        PR_DrawLineOffset(xyPt1, xyPt2, 0.6875) 'Inner tram line
        PR_DrawLineOffset(xyPt1, xyPt2, 0.1875) 'seam line at 3/16ths
        ' end of added lines

        ' Thumb opening
        nPalm = Val(strPalmNo) - Val(strFirstTape) + 1
        Ydiff = Profile.y(nPalm + 1) - Profile.y(nPalm)
        If System.Math.Abs(Ydiff) > 0 Then 'Avoids overflow
            xyThumbhole.y = Profile.y(nPalm) + (Ydiff / 2)
        Else
            xyThumbhole.y = Profile.y(nPalm)
        End If
        'Draw thumb hole construction line
        ARMDIA1.PR_SetLayer("Construct")
        xyThumbConstruct1.y = xyThumbhole.y
        ''--------------xyThumbConstruct1.X = Profile.X(nPalm) + (Val(txtPalmWristDist.Text) / 2)
        xyThumbConstruct1.X = Profile.X(nPalm) + (cmtoinches(Righttxtinchflag.Text, RighttxtPalmWrist) / 2)
        ARMDIA1.PR_MakeXY(xyThumbConstruct2, xyThumbConstruct1.X, 0.1875)
        ARMDIA1.PR_DrawLine(xyThumbConstruct1, xyThumbConstruct2)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHoleConstruct") 'g_sID from FN_Open

        PR_AddMarker(xyThumbConstruct1)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHoleConstruct") 'g_sID from FN_Open


        ' Lower line
        xyThumbhole.y = ((xyThumbhole.y - 0.1875) / 2) + 0.375
        xyThumbhole.X = Profile.X(nPalm)

        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
        ARMDIA1.PR_MakeXY(xyTHole1, xyThumbhole.X + 0.375, xyThumbhole.y - 0.1875)
        ARMDIA1.PR_MakeXY(xyTHole2, xyThumbhole.X + CDbl(strThumbDist) - 0.1875, xyThumbhole.y - 0.1875)
        ARMDIA1.PR_DrawLine(xyTHole1, xyTHole2)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHole") 'g_sID from FN_Open

        ' Upper Line
        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
        ARMDIA1.PR_MakeXY(xyTHole1, xyThumbhole.X + 0.375, xyThumbhole.y + 0.1875)
        ARMDIA1.PR_MakeXY(xyTHole2, xyThumbhole.X + CDbl(strThumbDist) - 0.1875, xyThumbhole.y + 0.1875)
        ARMDIA1.PR_DrawLine(xyTHole1, xyTHole2)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHole") 'g_sID from FN_Open

        ' Thumb Reminder in layer notes
        ARMDIA1.PR_SetLayer("Notes")
        ARMDIA1.PR_MakeXY(xyThumbNotes, xyThumbhole.X + 0.375, xyThumbhole.y + 0.5575)
        If RightchkNoThumb.CheckState = 0 Then
            ARMDIA1.PR_DrawText("Thumb Attached", xyThumbNotes, 0.125, 0)
        Else
            ARMDIA1.PR_DrawText("No Thumb", xyThumbNotes, 0.125, 0)
        End If

        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)

        ' Right Curve
        ARMDIA1.PR_MakeXY(xyCen, xyThumbhole.X + CDbl(strThumbDist) - 0.1875, xyThumbhole.y)
        ARMDIA1.PR_MakeXY(xyBegin, xyCen.X, xyCen.y + 0.1875)
        ARMDIA1.PR_MakeXY(xyFin, xyCen.X, xyCen.y - 0.1875)
        PR_RightThumbHole(xyCen, xyFin, xyBegin)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHole") 'g_sID from FN_Open

        ' Left top Curve
        ARMDIA1.PR_MakeXY(xyBegin, xyThumbhole.X + 0.375, xyThumbhole.y + 0.1875)
        ARMDIA1.PR_MakeXY(xyFin, xyThumbhole.X, xyThumbhole.y)
        lenty = ARMDIA1.FN_CalcLength(xyBegin, xyFin)
        midylen = lenty / 2
        nAng1 = ARMDIA1.FN_CalcAngle(xyFin, xyBegin)
        nAng1 = nAng1 * (ARMDIA1.PI / 180)
        nAdj1 = System.Math.Cos(nAng1) * midylen
        nOpp1 = System.Math.Sin(nAng1) * midylen
        ARMDIA1.PR_MakeXY(xyMidPoint, xyFin.X + nAdj1, xyFin.y + nOpp1)
        nlen = System.Math.Sqrt(0.25 - (midylen * midylen))
        nAdj2 = System.Math.Cos(nAng1) * nlen
        nOpp2 = System.Math.Sin(nAng1) * nlen
        ARMDIA1.PR_MakeXY(xyCen, xyMidPoint.X + nOpp2, xyMidPoint.y - nAdj2)
        ARMDIA1.PR_DrawArc(xyCen, xyBegin, xyFin)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHole") 'g_sID from FN_Open

        ' Left Bottom Curve
        ARMDIA1.PR_MakeXY(xyBegin, xyThumbhole.X + 0.375, xyThumbhole.y - 0.1875)
        ARMDIA1.PR_MakeXY(xyFin, xyThumbhole.X, xyThumbhole.y)
        lenty = ARMDIA1.FN_CalcLength(xyBegin, xyFin)
        midylen = lenty / 2
        nAng1 = ARMDIA1.FN_CalcAngle(xyFin, xyBegin)
        nAng1 = nAng1 * (ARMDIA1.PI / 180)
        nAdj1 = System.Math.Cos(nAng1) * midylen
        nOpp1 = System.Math.Sin(nAng1) * midylen
        ARMDIA1.PR_MakeXY(xyMidPoint, xyFin.X + nAdj1, xyFin.y + nOpp1)
        nlen = System.Math.Sqrt(0.25 - (midylen * midylen))
        nAdj2 = System.Math.Cos(nAng1) * nlen
        nOpp2 = System.Math.Sin(nAng1) * nlen
        ARMDIA1.PR_MakeXY(xyCen, xyMidPoint.X - nOpp2, xyMidPoint.y + nAdj2)
        ARMDIA1.PR_DrawArc(xyCen, xyFin, xyBegin)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHole") 'g_sID from FN_Open

        'Cutting line to Distal EOS
        ARMDIA1.PR_MakeXY(xyBegin, xyThumbhole.X - GauntSize, xyThumbhole.y)
        ARMDIA1.PR_DrawLine(xyBegin, xyThumbhole)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ThumbHole") 'g_sID from FN_Open

        If RightchkNoThumb.CheckState = 0 Then
            'Draw Thumb
            DrawRightThumb(xyThumbhole)
        End If
        If RightchkDetachable.CheckState = 1 Then
        End If
    End Sub
    Private Sub DrawRightThumb(ByRef xyThumbhole As ARMDIA1.XY)
        Dim nOpp2 As Object
        Dim nAdj2 As Object
        Dim nAng2 As Object
        Dim nlen As Object
        Dim nOpp1 As Object
        Dim nAdj1 As Object
        Dim nAng1 As Object
        Dim nHyp As Object
        Dim nAng As Object
        Dim midylen As Object
        Dim lenty As Object
        Dim xyThumb7, xyThumb5, xyThumb3, xyThumb1, xyThumbStart, xyThumb2, xyThumb4, xyThumb6, xyThumb8 As ARMDIA1.XY
        Dim xyThumbBase3, xyThumbBase1, xyThumbBase2, xyThumbBase4 As ARMDIA1.XY
        Dim thmang2, thm3, thm1, thm2, thmAng1, nTextHt As Object
        Dim xyCen, xyBegin, xyFin, xyMidPoint, xyPt As ARMDIA1.XY
        Dim ThumbBase As ARMDIA1.curve
        Dim ThumbBaseBit, ThumbBaseLen As Object
        Dim nThumbAngle, nGauntletPlateAngle As Double
        Dim nSpacing As Double
        'Draw Thumb
        ARMDIA1.PR_MakeXY(xyThumbStart, xyThumbhole.X, xyThumbhole.y + 3.1)
        ARMDIA1.PR_MakeXY(xyThumb1, xyThumbStart.X - 0.3125, xyThumbStart.y)
        ARMDIA1.PR_MakeXY(xyThumb2, xyThumbStart.X, xyThumb1.y + 0.25)

        'Calculate Thumb Angle
        nGauntletPlateAngle = 90 - 37
        ARMDIA1.PR_CalcPolar(xyThumb2, nGauntletPlateAngle, Val(strTCircum), xyPt)
        ARMDIA1.PR_MakeXY(xyThumb5, xyThumbStart.X + CDbl(strThumbDist), xyThumbStart.y)
        nThumbAngle = ARMDIA1.FN_CalcAngle(xyThumb5, xyPt)

        'Thumb length adjustments for enclosed tip
        If RightchkEnclosed.CheckState = 1 Then
            'Adjust by age
            If Val(txtAge.Text) > 14 Then
                strThumbLen = CStr(CDbl(strThumbLen) - 0.25)
            Else
                strThumbLen = CStr(CDbl(strThumbLen) - 0.125)
            End If
            'Adjust w.r.t radius
            strThumbLen = CStr(CDbl(strThumbLen) - (Val(strTCircum) / 2))
            If CDbl(strThumbLen) < 0 Then strThumbLen = CStr(0.125)
        End If

        ARMDIA1.PR_CalcPolar(xyThumb2, nThumbAngle, Val(strThumbLen), xyThumb3)
        ARMDIA1.PR_CalcPolar(xyThumb3, nThumbAngle - 90, Val(strTCircum), xyThumb4)

        thmAng1 = ARMDIA1.FN_CalcAngle(xyThumb5, xyThumb4)
        thmAng1 = 180 - thmAng1
        thmAng1 = thmAng1 * (ARMDIA1.PI / 180)
        thm1 = System.Math.Cos(thmAng1) * 0.1875
        thm2 = System.Math.Sin(thmAng1) * 0.1875
        ARMDIA1.PR_MakeXY(xyThumb6, xyThumb5.X + thm1, xyThumb5.y - thm2)
        ARMDIA1.PR_MakeXY(xyThumb7, xyThumb5.X - 0.1875, xyThumbStart.y - 0.375)
        ARMDIA1.PR_MakeXY(xyThumb8, xyThumbStart.X + 0.375, xyThumbStart.y - 0.375)
        ARMDIA1.PR_DrawLine(xyThumb2, xyThumb3)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbLength") 'sFileNo, sSide from FN_Open

        ' thumb tip
        If RightchkEnclosed.CheckState = 0 Then
            ARMDIA1.PR_DrawLine(xyThumb3, xyThumb4)
            ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbCircum") 'sFileNo, sSide from FN_Open
        ElseIf RightchkEnclosed.CheckState = 1 Then
            ThumbTip(xyThumb3, xyThumb4)
            ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbCircum") 'sFileNo, sSide from FN_Open
        End If
        ARMDIA1.PR_DrawLine(xyThumb4, xyThumb5)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbBack") 'sFileNo, sSide from FN_Open
        ARMDIA1.PR_DrawLine(xyThumb5, xyThumb6)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbBack2") 'sFileNo, sSide from FN_Open

        'Draw base of thumb
        ARMDIA1.PR_DrawLine(xyThumb7, xyThumb8)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumBase") 'sFileNo, sSide from FN_Open

        'Proximal bottom curve of thumb
        lenty = ARMDIA1.FN_CalcLength(xyThumb7, xyThumb6)
        midylen = lenty / 2
        nAng = ARMDIA1.FN_CalcAngle(xyThumb7, xyThumb6)
        nAng = (90 - nAng) * (ARMDIA1.PI / 180)
        nHyp = midylen / System.Math.Cos(nAng)
        ARMDIA1.PR_MakeXY(xyCen, xyThumb7.X, xyThumb7.y + nHyp)
        ARMDIA1.PR_DrawArc(xyCen, xyThumb7, xyThumb6)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbRightCurve") 'sFileNo, sSide from FN_Open

        'Distal Bottom thumb Curve
        lenty = ARMDIA1.FN_CalcLength(xyThumb1, xyThumb8)
        midylen = lenty / 2
        nAng = ARMDIA1.FN_CalcAngle(xyThumb8, xyThumb1)
        nAng = (nAng - 90) * (ARMDIA1.PI / 180)
        nHyp = midylen / System.Math.Cos(nAng)
        ARMDIA1.PR_MakeXY(xyCen, xyThumb8.X, xyThumb8.y + nHyp)
        ARMDIA1.PR_DrawArc(xyCen, xyThumb1, xyThumb8)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbLeftCurve") 'sFileNo, sSide from FN_Open

        'J Curve
        ARMDIA1.PR_MakeXY(xyBegin, xyThumb2.X, xyThumb2.y)
        ARMDIA1.PR_MakeXY(xyFin, xyThumb1.X, xyThumb1.y)
        lenty = ARMDIA1.FN_CalcLength(xyFin, xyBegin)
        midylen = lenty / 2
        nAng1 = ARMDIA1.FN_CalcAngle(xyFin, xyBegin)
        nAng1 = nAng1 * (ARMDIA1.PI / 180)
        nAdj1 = System.Math.Cos(nAng1) * midylen
        nOpp1 = System.Math.Sin(nAng1) * midylen
        ARMDIA1.PR_MakeXY(xyMidPoint, xyFin.X + nAdj1, xyFin.y + nOpp1)
        nlen = System.Math.Sqrt(0.1225 - (midylen * midylen))
        nAng2 = ARMDIA1.FN_CalcAngle(xyFin, xyBegin)
        nAng2 = 90 - nAng2
        nAng2 = nAng2 * (ARMDIA1.PI / 180)
        nAdj2 = System.Math.Cos(nAng2) * nlen
        nOpp2 = System.Math.Sin(nAng2) * nlen
        ARMDIA1.PR_MakeXY(xyCen, xyMidPoint.X - nAdj2, xyMidPoint.y + nOpp2)
        ARMDIA1.PR_DrawArc(xyCen, xyFin, xyBegin)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbJCurve") 'sFileNo, sSide from FN_Open

        ARMDIA1.PR_SetLayer("Notes")
        nTextHt = 0.125
        ARMDIA1.PR_MakeXY(xyPt, xyThumbStart.X, xyThumbStart.y + 0.375)

        'Main details in Green on layer notes
        Dim sWorkOrder, sText As String

        ARMDIA1.PR_SetLayer("Notes")
        ARMDIA1.PR_SetTextData(1, 32, -1, -1, -1) 'Horiz-Left, Vertical-Bottom

        ARMDIA1.PR_SetLayer("Notes")
        If txtWorkOrder.Text = "" Then
            sWorkOrder = "-"
        Else
            sWorkOrder = txtWorkOrder.Text
        End If
        sText = txtSleeve.Text & Chr(10) & txtPatientName.Text & Chr(10) & sWorkOrder
        ''ARMDIA1.PR_DrawText(sText, xyPt, 0.125, 0)
        ARMDIA1.PR_DrawMText(sText, xyPt, False)

    End Sub

    Private Sub DrawThumb(ByRef xyThumbhole As ARMDIA1.XY)
        Dim nOpp2 As Object
        Dim nAdj2 As Object
        Dim nAng2 As Object
        Dim nlen As Object
        Dim nOpp1 As Object
        Dim nAdj1 As Object
        Dim nAng1 As Object
        Dim nHyp As Object
        Dim nAng As Object
        Dim midylen As Object
        Dim lenty As Object
        Dim xyThumb7, xyThumb5, xyThumb3, xyThumb1, xyThumbStart, xyThumb2, xyThumb4, xyThumb6, xyThumb8 As ARMDIA1.XY
        Dim xyThumbBase3, xyThumbBase1, xyThumbBase2, xyThumbBase4 As ARMDIA1.XY
        Dim thmang2, thm3, thm1, thm2, thmAng1, nTextHt As Object
        Dim xyCen, xyBegin, xyFin, xyMidPoint, xyPt As ARMDIA1.XY
        'UPGRADE_WARNING: Arrays in structure ThumbBase may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Dim ThumbBase As ARMDIA1.curve
        Dim ThumbBaseBit, ThumbBaseLen As Object
        Dim nThumbAngle, nGauntletPlateAngle As Double
        Dim nSpacing As Double
        'Draw Thumb
        ARMDIA1.PR_MakeXY(xyThumbStart, xyThumbhole.X, xyThumbhole.y + 3.1)
        ARMDIA1.PR_MakeXY(xyThumb1, xyThumbStart.X - 0.3125, xyThumbStart.y)
        ARMDIA1.PR_MakeXY(xyThumb2, xyThumbStart.X, xyThumb1.y + 0.25)

        'Calculate Thumb Angle
        nGauntletPlateAngle = 90 - 37
        ARMDIA1.PR_CalcPolar(xyThumb2, nGauntletPlateAngle, Val(txtTCircum.Text), xyPt)
        ARMDIA1.PR_MakeXY(xyThumb5, xyThumbStart.X + CDbl(txtThumbDist.Text), xyThumbStart.y)
        nThumbAngle = ARMDIA1.FN_CalcAngle(xyThumb5, xyPt)

        'Thumb length adjustments for enclosed tip
        If chkEnclosed.CheckState = 1 Then
            'Adjust by age
            If Val(txtAge.Text) > 14 Then
                txtThumLen.Text = CStr(CDbl(txtThumLen.Text) - 0.25)
            Else
                txtThumLen.Text = CStr(CDbl(txtThumLen.Text) - 0.125)
            End If
            'Adjust w.r.t radius
            txtThumLen.Text = CStr(CDbl(txtThumLen.Text) - (Val(txtTCircum.Text) / 2))
            If CDbl(txtThumLen.Text) < 0 Then txtThumLen.Text = CStr(0.125)
        End If

        ARMDIA1.PR_CalcPolar(xyThumb2, nThumbAngle, Val(txtThumLen.Text), xyThumb3)
        ARMDIA1.PR_CalcPolar(xyThumb3, nThumbAngle - 90, Val(txtTCircum.Text), xyThumb4)

        'UPGRADE_WARNING: Couldn't resolve default property of object thmAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        thmAng1 = ARMDIA1.FN_CalcAngle(xyThumb5, xyThumb4)
        'UPGRADE_WARNING: Couldn't resolve default property of object thmAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        thmAng1 = 180 - thmAng1
        'UPGRADE_WARNING: Couldn't resolve default property of object thmAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        thmAng1 = thmAng1 * (ARMDIA1.PI / 180)
        'UPGRADE_WARNING: Couldn't resolve default property of object thmAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object thm1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        thm1 = System.Math.Cos(thmAng1) * 0.1875
        'UPGRADE_WARNING: Couldn't resolve default property of object thmAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object thm2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        thm2 = System.Math.Sin(thmAng1) * 0.1875
        'UPGRADE_WARNING: Couldn't resolve default property of object thm2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object thm1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMDIA1.PR_MakeXY(xyThumb6, xyThumb5.X + thm1, xyThumb5.y - thm2)
        ARMDIA1.PR_MakeXY(xyThumb7, xyThumb5.X - 0.1875, xyThumbStart.y - 0.375)
        ARMDIA1.PR_MakeXY(xyThumb8, xyThumbStart.X + 0.375, xyThumbStart.y - 0.375)
        ARMDIA1.PR_DrawLine(xyThumb2, xyThumb3)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbLength") 'sFileNo, sSide from FN_Open

        ' thumb tip
        If chkEnclosed.CheckState = 0 Then
            ARMDIA1.PR_DrawLine(xyThumb3, xyThumb4)
            ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbCircum") 'sFileNo, sSide from FN_Open
        ElseIf chkEnclosed.CheckState = 1 Then
            ThumbTip(xyThumb3, xyThumb4)
            ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbCircum") 'sFileNo, sSide from FN_Open
        End If
        ARMDIA1.PR_DrawLine(xyThumb4, xyThumb5)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbBack") 'sFileNo, sSide from FN_Open
        ARMDIA1.PR_DrawLine(xyThumb5, xyThumb6)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbBack2") 'sFileNo, sSide from FN_Open

        'Draw base of thumb
        ARMDIA1.PR_DrawLine(xyThumb7, xyThumb8)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumBase") 'sFileNo, sSide from FN_Open

        'Proximal bottom curve of thumb
        'UPGRADE_WARNING: Couldn't resolve default property of object lenty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lenty = ARMDIA1.FN_CalcLength(xyThumb7, xyThumb6)
        'UPGRADE_WARNING: Couldn't resolve default property of object lenty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        midylen = lenty / 2
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAng = ARMDIA1.FN_CalcAngle(xyThumb7, xyThumb6)
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAng = (90 - nAng) * (ARMDIA1.PI / 180)
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nHyp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nHyp = midylen / System.Math.Cos(nAng)
        'UPGRADE_WARNING: Couldn't resolve default property of object nHyp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMDIA1.PR_MakeXY(xyCen, xyThumb7.X, xyThumb7.y + nHyp)
        ARMDIA1.PR_DrawArc(xyCen, xyThumb7, xyThumb6)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbRightCurve") 'sFileNo, sSide from FN_Open

        'Distal Bottom thumb Curve
        'UPGRADE_WARNING: Couldn't resolve default property of object lenty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lenty = ARMDIA1.FN_CalcLength(xyThumb1, xyThumb8)
        'UPGRADE_WARNING: Couldn't resolve default property of object lenty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        midylen = lenty / 2
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAng = ARMDIA1.FN_CalcAngle(xyThumb8, xyThumb1)
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAng = (nAng - 90) * (ARMDIA1.PI / 180)
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nHyp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nHyp = midylen / System.Math.Cos(nAng)
        'UPGRADE_WARNING: Couldn't resolve default property of object nHyp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMDIA1.PR_MakeXY(xyCen, xyThumb8.X, xyThumb8.y + nHyp)
        ARMDIA1.PR_DrawArc(xyCen, xyThumb1, xyThumb8)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbLeftCurve") 'sFileNo, sSide from FN_Open

        'J Curve
        ARMDIA1.PR_MakeXY(xyBegin, xyThumb2.X, xyThumb2.y)
        ARMDIA1.PR_MakeXY(xyFin, xyThumb1.X, xyThumb1.y)
        'UPGRADE_WARNING: Couldn't resolve default property of object lenty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lenty = ARMDIA1.FN_CalcLength(xyFin, xyBegin)
        'UPGRADE_WARNING: Couldn't resolve default property of object lenty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        midylen = lenty / 2
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAng1 = ARMDIA1.FN_CalcAngle(xyFin, xyBegin)
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAng1 = nAng1 * (ARMDIA1.PI / 180)
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAdj1 = System.Math.Cos(nAng1) * midylen
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nOpp1 = System.Math.Sin(nAng1) * midylen
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMDIA1.PR_MakeXY(xyMidPoint, xyFin.X + nAdj1, xyFin.y + nOpp1)
        'UPGRADE_WARNING: Couldn't resolve default property of object midylen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nlen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nlen = System.Math.Sqrt(0.1225 - (midylen * midylen))
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAng2 = ARMDIA1.FN_CalcAngle(xyFin, xyBegin)
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAng2 = 90 - nAng2
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAng2 = nAng2 * (ARMDIA1.PI / 180)
        'UPGRADE_WARNING: Couldn't resolve default property of object nlen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAdj2 = System.Math.Cos(nAng2) * nlen
        'UPGRADE_WARNING: Couldn't resolve default property of object nlen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nOpp2 = System.Math.Sin(nAng2) * nlen
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMDIA1.PR_MakeXY(xyCen, xyMidPoint.X - nAdj2, xyMidPoint.y + nOpp2)
        ARMDIA1.PR_DrawArc(xyCen, xyFin, xyBegin)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "ThumbJCurve") 'sFileNo, sSide from FN_Open

        ARMDIA1.PR_SetLayer("Notes")
        'UPGRADE_WARNING: Couldn't resolve default property of object nTextHt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nTextHt = 0.125
        ARMDIA1.PR_MakeXY(xyPt, xyThumbStart.X, xyThumbStart.y + 0.375)

        'Main details in Green on layer notes
        Dim sWorkOrder, sText As String

        ARMDIA1.PR_SetLayer("Notes")
        ARMDIA1.PR_SetTextData(1, 32, -1, -1, -1) 'Horiz-Left, Vertical-Bottom

        ARMDIA1.PR_SetLayer("Notes")
        If txtWorkOrder.Text = "" Then
            sWorkOrder = "-"
        Else
            sWorkOrder = txtWorkOrder.Text
        End If
        sText = txtSleeve.Text & Chr(10) & txtPatientName.Text & Chr(10) & sWorkOrder
        ''ARMDIA1.PR_DrawText(sText, xyPt, 0.125, 0)
        ARMDIA1.PR_DrawMText(sText, xyPt, False)


    End Sub

    Private Sub Elastic(ByRef Profile As ARMDIA1.curve)
        'UPGRADE_WARNING: Arrays in structure elastictop may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Dim elastictop As New ARMDIA1.curve
        Dim xyEpt2, xyEpt1, xyEpt3 As New ARMDIA1.XY

        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMDIA1.g_nCurrTextAngle = 90
        ARMDIA1.PR_SetTextData(2, 16, -1, -1, -1) 'Horiz center, Vertical center

        elastictop.n = 3
        Dim tempDataX(3) As Double
        Dim tempDataY(3) As Double
        tempDataX(1) = Profile.X(Profile.n)
        tempDataY(1) = Profile.y(Profile.n) / 2
        tempDataY(2) = Profile.y(Profile.n) / 2
        tempDataY(3) = Profile.y(Profile.n) + 0.375

        'Proximal Elastic and Cutback
        If Val(txtAge.Text) > 10 And Val(txtLastTape.Text) = 11 Then
            tempDataX(2) = Profile.X(Profile.n) - 0.75
            tempDataX(3) = Profile.X(Profile.n) - 0.75
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
            elastictop.X = tempDataX
            elastictop.y = tempDataY
            '-----------ARMDIA1.PR_DrawPoly(elastictop)
        ElseIf Val(txtAge.Text) <= 10 And Val(txtLastTape.Text) = 11 Then
            'Proximal Elastic @ Elbow
            tempDataX(2) = Profile.X(Profile.n) - 0.375
            tempDataX(3) = Profile.X(Profile.n) - 0.375
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
            elastictop.X = tempDataX
            elastictop.y = tempDataY
            ARMDIA1.PR_DrawPoly(elastictop)
            ARMDIA1.PR_MakeXY(xyEpt1, Profile.X(Profile.n) - 0.75, Profile.y(Profile.n) / 2)
            ARMDIA1.PR_SetLayer("Notes")
            ''--------ARMDIA1.PR_DrawText("1/2'' Elastic", xyEpt1, 0.125, 0)
            ARMDIA1.PR_DrawTextCenter("1/2'' Elastic", xyEpt1, 0.125, (90 * (ARMDIA1.PI / 180)))
        ElseIf cboFlaps.Text = "" And Val(txtAge.Text) <= 10 And Val(txtLastTape.Text) <> 11 Then
            'Proximal Elastic on plain arms for children under 10
            ARMDIA1.PR_MakeXY(xyEpt1, Profile.X(Profile.n) - 0.25, Profile.y(Profile.n) / 2)
            ARMDIA1.PR_SetLayer("Notes")
            ''-------ARMDIA1.PR_DrawText("1/2'' Elastic", xyEpt1, 0.125, (90 * (ARMDIA1.PI / 180)))
            ARMDIA1.PR_DrawTextCenter("1/2'' Elastic", xyEpt1, 0.125, (90 * (ARMDIA1.PI / 180)))
        End If

        elastictop.X = tempDataX
        elastictop.y = tempDataY

        'Distal Elastic and Cutback
        elastictop.X(1) = Profile.X(1)
        elastictop.y(1) = Profile.y(1) / 2
        elastictop.y(2) = Profile.y(1) / 2
        elastictop.X(3) = Profile.X(1) + 0.375
        elastictop.y(3) = Profile.y(1) + 0.375
        ARMDIA1.PR_MakeXY(xyEpt1, Profile.X(1) + 0.75, Profile.y(1) / 2)

        If Val(txtFirstTape.Text) = 11 And Val(txtAge.Text) <= 10 Then
            elastictop.X(2) = Profile.X(1) + 0.375
            elastictop.X(3) = Profile.X(1) + 0.375
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
            ARMDIA1.PR_DrawPoly(elastictop)
            ARMDIA1.PR_SetLayer("Notes")
            ''----------ARMDIA1.PR_DrawText("1/2'' Elastic", xyEpt1, 0.125, 0)
            ARMDIA1.PR_DrawTextCenter("1/2'' Elastic", xyEpt1, 0.125, (90 * (ARMDIA1.PI / 180)))

        ElseIf Val(txtFirstTape.Text) = 11 And Val(txtAge.Text) > 10 Then
            elastictop.X(2) = Profile.X(1) + 0.75
            elastictop.X(3) = Profile.X(1) + 0.75
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
            ARMDIA1.PR_DrawPoly(elastictop)

        ElseIf Val(txtFirstTape.Text) > 11 And Val(txtAge.Text) <= 10 Then
            'For children with sleeves above elbow then use 1/2" elastic
            ARMDIA1.PR_SetLayer("Notes")
            ''---------ARMDIA1.PR_DrawText("1/2'' Elastic", xyEpt1, 0.125, 0)
            ARMDIA1.PR_DrawTextCenter("1/2'' Elastic", xyEpt1, 0.125, (90 * (ARMDIA1.PI / 180)))
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object g_nCurrTextAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ARMDIA1.g_nCurrTextAngle = 0


    End Sub

    Private Function FN_CheckValue(ByRef TextBox As System.Windows.Forms.Control, ByRef sMessage As String) As Short
        ' Check for numeric values
        'Dim sChar As String, sText As String
        'Dim iLen As Integer, nn As Integer
        'sText = TextBox.Text
        'iLen = Len(sText)
        'FN_CheckValue = True
        'For nn = 1 To iLen
        '    sChar = Mid$(sText, nn, 1)
        '    If Asc(sChar) > 57 Or Asc(sChar) < 46 Or Asc(sChar) = 47 Then
        '        MsgBox "Invalid " & sMessage & " Length has been entered", 48, "Arm Details"
        '        TextBox.SetFocus
        '        FN_CheckValue = False
        '        Exit For
        '    End If
        'Next nn

        FN_CheckValue = True
        If Not IsNumeric(TextBox.Text) And Len(TextBox.Text) > 0 Then
            MsgBox("Invalid " & sMessage & " Length has been entered", 48, "Arm Details")
            TextBox.Focus()
            FN_CheckValue = False
        End If

    End Function


    Private Function FN_PleatDistances(ByRef flag As Object, ByRef cm As Object) As Object
        Dim temprem, remainder, inch, temp, rnder As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object flag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If CType(flag, TextBox).Text = "cm" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object inch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            inch = Val(cm) / 2.54
            'UPGRADE_WARNING: Couldn't resolve default property of object inch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FN_PleatDistances. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FN_PleatDistances = inch
            'UPGRADE_WARNING: Couldn't resolve default property of object flag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf CType(flag, TextBox).Text = "inches" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            temp = Int(CDbl(cm))
            'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object cm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object remainder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            remainder = cm - temp
            If remainder > 0.7 Then
                MsgBox("Invalid Measurements, Inches and Eights only", 48, "Arm Details")
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_PleatDistances. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FN_PleatDistances = "-1"
                Exit Function
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object remainder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                remainder = remainder * 10
                'UPGRADE_WARNING: Couldn't resolve default property of object remainder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                remainder = remainder * 0.125
                'UPGRADE_WARNING: Couldn't resolve default property of object remainder. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object temp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_PleatDistances. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FN_PleatDistances = temp + remainder
            End If
        End If
    End Function



    Private Function fnInchesToText(ByRef nInches As Double) As String
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
            sString = " 0"
        End If
        If iEighths <> 0 Then
            Select Case iEighths
                Case 2, 6
                    sString = sString & "-" & LTrim(Str(iEighths / 2)) & "/4"
                Case 4
                    sString = sString & "-" & "1/2"
                Case Else
                    sString = sString & "-" & LTrim(Str(iEighths)) & "/8"
            End Select
        Else
            sString = sString & "   "
        End If

        'Return formatted string
        fnInchesToText = sString

    End Function

    'UPGRADE_ISSUE: Form event Form.LinkClose was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
    Private Sub Form_LinkClose()
        Dim Index As Object
        Dim w As Object
        Dim g_sFrontStrapLength As Object
        Dim sData As Object
        Dim k As Object
        Dim def2, rtmp, X, i, sTmp, def1, cmp As Object
        Dim nValue As Double
        Dim nAge As Short
        Dim sFlap, sDiag As String
        'Disable timer
        'Stop the timer used to ensure that the Dialogue dies
        'if the DRAFIX macro fails to establish a DDE Link
        Timer1.Enabled = False
        ''Commented on 23-06-2018
        ''If txtUidMPD.Text = "" Then
        ''    MsgBox("Can't find Patient Details", 16, "Arm Details Dialogue")
        ''    'End
        ''    Return
        ''End If

        ''cboFabric.Text = txtFabric.Text
        ''ARMDIA1.g_sUnits = txtinchflag.Text

        ''' set form caption
        ''If txtSleeve.Text = "Left" Then
        ''    Me.Text = "ARM Details - Left"
        ''Else
        ''    Me.Text = "ARM Details - Right"
        ''End If

        '''Pressure
        ''cboPressure.Text = txtMM.Text
        ' if pressure empty set from diagnosis
        nAge = Val(txtAge.Text)
        sDiag = UCase(Mid(txtDiagnosis.Text, 1, 6))
        If txtMM.Text = "" Then
            'Arterial Insufficiency
            If sDiag = "ARTERI" Then
                cboPressure.Text = "15mm"
                'Assist fluid dynamics
            ElseIf sDiag = "ASSIST" Then
                cboPressure.Text = "20mm"
                'Blood Clots
            ElseIf sDiag = "BLOOD " Then
                cboPressure.Text = "20mm"
                'Burns
            ElseIf sDiag = "BURNS " Or sDiag = "BURNS" Then
                cboPressure.Text = "15mm"
                'Cancer
            ElseIf sDiag = "CANCER" Then
                cboPressure.Text = "20mm"
                'Cardial Vascular Arrest
            ElseIf sDiag = "CARDIA" Then
                cboPressure.Text = "15mm"
                'Carpal Tunnel Syndrome
            ElseIf sDiag = "CARPAL" Then
                cboPressure.Text = "20mm"
                'Cellulitis
            ElseIf sDiag = "CELLUL" Then
                cboPressure.Text = "20mm"
                'Chronic Venous Insufficiency
            ElseIf sDiag = "CHRONI" Then
                cboPressure.Text = "20mm"
                'Heart Condition
            ElseIf sDiag = "HEART " Then
                cboPressure.Text = "15mm"
                'Hemangioma
            ElseIf sDiag = "HEMANG" Then
                cboPressure.Text = "15mm"
                'Lymphedema, Lymphedema 1+, Lymphedema 2+
            ElseIf sDiag = "LYMPHE" Then
                If InStr(txtDiagnosis.Text, "2") > 0 Then
                    If nAge <= 70 Then
                        cboPressure.Text = "25mm"
                    Else
                        cboPressure.Text = "20mm"
                    End If
                Else
                    If nAge <= 70 Then
                        cboPressure.Text = "20mm"
                    Else
                        cboPressure.Text = "15mm"
                    End If
                End If
                'Night Wear
            ElseIf sDiag = "NIGHT " Then
                cboPressure.Text = "15mm"
                'Post Fracture
            ElseIf sDiag = "POST F" Then
                cboPressure.Text = "20mm"
                'Post Mastectomy
            ElseIf sDiag = "POST M" Then
                cboPressure.Text = "20mm"
                'Postphlebetic Syndrome
            ElseIf sDiag = "POSTPH" Then
                cboPressure.Text = "20mm"
                'Renal Disease (Kidney)
            ElseIf sDiag = "RENAL " Then
                cboPressure.Text = "20mm"
                'Skin Graft
            ElseIf sDiag = "SKIN G" Then
                cboPressure.Text = "15mm"
                'Stroke
            ElseIf sDiag = "STROKE" Then
                cboPressure.Text = "15mm"
                'Tendonitis
            ElseIf sDiag = "TENDON" Then
                cboPressure.Text = "15mm"
                'Thrombophlebitis
            ElseIf sDiag = "THROMB" Then
                cboPressure.Text = "20mm"
                'Trauma
            ElseIf sDiag = "TRAUMA" Then
                cboPressure.Text = "20mm"
                'Varicose Veins
            ElseIf sDiag = "VARICO" Then
                cboPressure.Text = "20mm"
                'Request for 30 m/m, 40 m/m and 50 m/m
            ElseIf sDiag = "REQUES" Then
                If InStr(txtDiagnosis.Text, "30") > 0 Then
                    cboPressure.Text = "15mm"
                ElseIf InStr(txtDiagnosis.Text, "40") > 0 Then
                    cboPressure.Text = "20mm"
                ElseIf InStr(txtDiagnosis.Text, "50") > 0 Then
                    cboPressure.Text = "25mm"
                End If
                'Vein Removal
            ElseIf sDiag = "VEIN R" Then
                cboPressure.Text = "20mm"
            End If
        End If

        'Pleats, break up multiple fields
        'NB the order
        nValue = FN_GetNumber(txtWristPleat2.Text, 1)
        If nValue > 0 Then txtWristPleat2.Text = CStr(nValue) Else txtWristPleat2.Text = ""
        nValue = FN_GetNumber(txtWristPleat1.Text, 1)
        If nValue > 0 Then txtWristPleat1.Text = CStr(nValue) Else txtWristPleat1.Text = ""
        ' Strip_Bad_Char txtShoulderPleat1
        'NB the order
        nValue = FN_GetNumber(txtShoulderPleat2.Text, 1)
        If nValue > 0 Then txtShoulderPleat2.Text = CStr(nValue) Else txtShoulderPleat2.Text = ""
        nValue = FN_GetNumber(txtShoulderPleat1.Text, 1)
        If nValue > 0 Then txtShoulderPleat1.Text = CStr(nValue) Else txtShoulderPleat1.Text = ""

        'Stump
        ' Strip_Bad_Char txtStump
        If txtStump.Text = "True" Then
            chkStump.Checked = 1
            chkStump_CheckedChanged(chkStump, New System.EventArgs())
        Else
            chkStump.Checked = 0
        End If

        LowLength(txtTapeLent)
        LowWeight(txtWeight)
        LowRed(txtReduction)
        LowTapeMM(txtTapeMM)
        For i = 0 To 17
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            k = i + 1
            'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object ChangeChecker(k). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ChangeChecker(k) = Length(i).Text
        Next i


        'Gauntlet
        'Break up gauntlet multiple field
        If FN_GetNumber(txtGauntlet.Text, 1) = 1 Then txtGauntletFlag.Text = "True"
        If FN_GetNumber(txtGauntlet.Text, 1) = 0 Then txtGauntletFlag.Text = "False"

        If txtGauntletFlag.Text <> "True" Or txtGauntletFlag.Text <> "False" Then
            g_sOld = "True"
        Else
            g_sOld = "False"
        End If

        If txtGauntletFlag.Text = "True" Then

            'Flags
            If FN_GetNumber(txtGauntlet.Text, 2) Then txtEnclosed.Text = "True"
            If FN_GetNumber(txtGauntlet.Text, 3) Then txtDetach.Text = "True"
            If FN_GetNumber(txtGauntlet.Text, 4) Then txtNoThumb.Text = "True"
            'Data
            txtWristNo.Text = CStr(FN_GetNumber(txtGauntlet.Text, 5))
            txtPalmNo.Text = CStr(FN_GetNumber(txtGauntlet.Text, 6))
            nValue = FN_GetNumber(txtGauntlet.Text, 7)
            If nValue > 0 Then txtThumLen.Text = CStr(nValue)
            txtTCircum.Text = CStr(FN_GetNumber(txtGauntlet.Text, 8))
            txtPalmWristDist.Text = CStr(FN_GetNumber(txtGauntlet.Text, 9))
            nValue = FN_GetNumber(txtGauntlet.Text, 10)
            If nValue > 0 Then txtGauntletExtension.Text = CStr(nValue)
            'Setup rest as per normal
            cboWrist.SelectedIndex = Val(txtWristNo.Text) - 1
            cboPalm.SelectedIndex = Val(txtPalmNo.Text) - 1
            txtCircum.Enabled = True
            lblCircum.Enabled = True
            OldGauntlet(txtPalmNo.Text, txtWristNo.Text, txtPalmWristDist.Text, txtTCircum.Text, txtThumLen.Text)
            chkGauntlets.Checked = True
            lblPalm.Enabled = True
            lblWrist.Enabled = True
            lblThumb.Enabled = True
            cboPalm.Enabled = True
            cboWrist.Enabled = True
            txtThumb.Enabled = True
            chkEnclosed.Enabled = True
            If txtEnclosed.Text = "True" Then
                chkEnclosed.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            chkDetachable.Enabled = True
            If txtDetach.Text = "True" Then
                chkDetachable.CheckState = System.Windows.Forms.CheckState.Checked
                chkDetachable_CheckStateChanged(chkDetachable, New System.EventArgs())
            End If
            chkNoThumb.Enabled = True
            If txtNoThumb.Text = "True" Then
                chkNoThumb.CheckState = System.Windows.Forms.CheckState.Checked
                txtThumb.Text = ""
                txtCircum.Text = ""
                txtCircum.Enabled = False
                lblCircum.Enabled = False
            End If
            lblPalmWrist.Enabled = True
            txtPalmWrist.Enabled = True
        ElseIf txtStump.Text <> "True" Then
            chkNone.Checked = 1
            chkNone_CheckedChanged(chkNone, New System.EventArgs())
        End If

        'Flap, break up flap multiple field
        'In this case the flap is 35 char long and the strap length is from char 36
        If txtType.Text = "SLEEVE" Then cboFlaps.Items.Add("Vest")
        If txtType.Text = "SLEEVE-BODY" Then cboFlaps.Items.Add("Body")

        If (txtFlap.Text = "") Or Mid(txtFlap.Text, 1, 4) = "None" Or txtType.Text = "SLEEVE" Or txtType.Text = "SLEEVE-BODY" Then
            'For new vests raglan set default to vest
            If txtType.Text = "SLEEVE" Then
                cboFlaps.Text = "Vest"
                optProximalTape(1).Checked = True
                optProximalTape_CheckedChanged(optProximalTape.Item(1), New System.EventArgs())
                cboFlaps_SelectedIndexChanged(cboFlaps, New System.EventArgs())
            ElseIf txtType.Text = "SLEEVE-BODY" Then
                cboFlaps.Text = "Body"
                optProximalTape(1).Checked = True
                optProximalTape_CheckedChanged(optProximalTape.Item(1), New System.EventArgs())
                cboFlaps_SelectedIndexChanged(cboFlaps, New System.EventArgs())
            Else
                txtFlap.Text = ""
                optProximalTape(0).Checked = True
                optProximalTape_CheckedChanged(optProximalTape.Item(0), New System.EventArgs())
            End If
        Else
            optProximalTape(1).Checked = True
            optProximalTape_CheckedChanged(optProximalTape.Item(1), New System.EventArgs())
            'UPGRADE_WARNING: Couldn't resolve default property of object sData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ''---------------
            'sData = Mid(txtFlap.Text, 36) 'NB order is very important
            'txtFlap.Text = VB.Left(txtFlap.Text, 35)
            'cboFlaps.Text = txtFlap.Text
            ''--------------------
            sData = txtFlap.Text
            cboFlaps.SelectedIndex = FN_GetNumber(sData, 5)
            'UPGRADE_WARNING: Couldn't resolve default property of object sData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nValue = FN_GetNumber(sData, 1) 'Strap length
            If nValue > 0 Then
                txtStrapLength.Text = CStr(nValue)
                txtStrap.Text = CStr(nValue)
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                labStrap.Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtStrap))
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object sData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nValue = FN_GetNumber(sData, 2) 'Custom flap length
            If nValue > 0 Then
                'Enable custom flap
                txtCustFlapLength.Enabled = True
                labCustFlapLength.Enabled = True
                txtCustFlapLength.Text = CStr(nValue)
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                labCustFlapLength.Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtCustFlapLength))
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object sData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nValue = FN_GetNumber(sData, 3) 'Waist Circumference (Style-D only)
            If nValue > 0 Then
                'Enable custom flap
                txtWaistCir.Enabled = True
                lblWaistCir.Enabled = True
                labWaistCir.Enabled = True
                txtWaistCir.Text = CStr(nValue)
                'UPGRADE_WARNING: Couldn't resolve default property of object g_sWaistCir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                g_sWaistCir = nValue
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                labWaistCir.Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtWaistCir))
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object sData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nValue = FN_GetNumber(sData, 4) 'Front Strap
            If nValue > 0 Then
                'Enable custom flap
                txtFrontStrapLength.Enabled = True
                lblFrontStrapLength.Enabled = True
                labFrontStrapLength.Enabled = True
                txtFrontStrapLength.Text = CStr(nValue)
                'UPGRADE_WARNING: Couldn't resolve default property of object g_sFrontStrapLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                g_sFrontStrapLength = nValue
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                labFrontStrapLength.Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtFrontStrapLength))
            End If

            cboFlaps_SelectedIndexChanged(cboFlaps, New System.EventArgs())
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object EmptySleeve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        EmptySleeve = True
        For w = 0 To 17
            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            X = w + 1
            If Length(w).Text <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object EmptySleeve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                EmptySleeve = False
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Measurements(X) = cmtoinches(txtinchflag.Text, Length(w))
            End If
        Next w

        Dim inchbit As Double
        For i = 0 To 17
            If Val(Length(i).Text) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                inchbit = cmtoinches(txtinchflag.Text, Length(i))
                If inchbit = -1 Then
                    Length(i).Focus()
                    Exit Sub
                End If
                InchText(i).Text = ARMEDDIA1.fnInchesToText(inchbit)
            End If
        Next

        Dim pwinch As Double
        Dim circinch, thminch As Object
        If Val(txtPalmWrist.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            pwinch = cmtoinches(txtinchflag.Text, txtPalmWrist)
            PWDist.Text = ARMEDDIA1.fnInchesToText(pwinch)
        End If

        If Val(txtCircum.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            pwinch = cmtoinches(txtinchflag.Text, txtCircum)
            TCirc.Text = ARMEDDIA1.fnInchesToText(pwinch)
        End If

        If Val(txtThumb.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            pwinch = cmtoinches(txtinchflag.Text, txtThumb)
            TLen.Text = ARMEDDIA1.fnInchesToText(pwinch)
        End If

        If Val(txtGauntletExtension.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            labGauntletExtension.Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtGauntletExtension))
        End If


        PR_GetStartEndTapes()
        Dim Gram As Object
        For Index = 0 To 17
            If Length(Index).Text <> "" And cboFabric.Text <> "" And cboPressure.Text <> "" And mms(Index).Text <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                inchbit = cmtoinches(txtinchflag.Text, Length(Index))
                If inchbit = -1 Then
                    Length(Index).Focus()
                    Exit Sub
                End If
                InchText(Index).Text = ARMEDDIA1.fnInchesToText(inchbit)
                'UPGRADE_WARNING: Couldn't resolve default property of object Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                w = Index
                'UPGRADE_WARNING: Couldn't resolve default property of object Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                X = Index + 1
                PR_GetFabric()
                '    Gram = Fix(Val(mms(Index)) * Val(Measurements(Index + 1)))
                '    gms(Index) = round(Gram)
                '    reds(Index) = GetReduction(g_sModulus, Gram, g_sFabnam)
                PR_WristRed(w, X, g_sModulus)
                PR_PalmRed(w, X, g_sModulus)
                PR_LastTapeRed(w, X, g_sModulus)
                PR_SetRresults()
                PR_SetGlobals()
            End If
        Next Index

        'Set up first and last tapes depending on ID
        'nValue = Val(Mid(txtID.Text, 1, 2))
        'If nValue > 0 Then
        '    cboDistalTape.SelectedIndex = nValue
        'Else
        '    cboDistalTape.SelectedIndex = 0
        'End If
        If cboDistalTape.SelectedIndex < 0 Then
            cboDistalTape.SelectedIndex = 0
        End If
        cboDistalTape_SelectedIndexChanged(cboDistalTape, New System.EventArgs())

        'nValue = Val(Mid(txtID.Text, 3, 2))
        'If nValue > 0 Then
        '    cboProximalTape.SelectedIndex = (18 - nValue) + 1
        'ElseIf Mid(txtID.Text, 3, 2) <> "GT" Then  'IE Not for Detachable Gauntlet
        '    cboProximalTape.SelectedIndex = 0
        'End If
        If cboProximalTape.SelectedIndex < 0 Then
            cboProximalTape.SelectedIndex = 0
        End If
        cboProximalTape_SelectedIndexChanged(cboProximalTape, New System.EventArgs())

        Show()
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    'UPGRADE_ISSUE: Form event Form.LinkExecute was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
    Private Sub Form_LinkExecute(ByRef CmdStr As String, ByRef Cancel As Short)
        If CmdStr = "Cancel" Then
            Cancel = 0
            Return
        End If
    End Sub

    Private Sub armdia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            Dim ii As Object
            Dim fileNum As Object
            Dim txtACwo As Object

            Hide()
            ' Maintain while loading DDE data
            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor ' change pointer to hourglass
            ' Reset in form_linkclose

            ' Position to center of screen
            Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
            Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)

            ARMDIA1.MainForm = Me
            g_sAlreadyLoad = False
            g_bIsClose = False
            ARMDIA1.g_Modified = False
            g_sLeftArmValues = ""
            g_sRightArmValues = ""
            Dim i As Object
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            i = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Do While i < 18
                reds(i).Text = ""
                gms(i).Text = ""
                mms(i).Text = ""
                Length(i).Text = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                i = i + 1
            Loop
            txtVestRaglan.Text = ""
            txtVestID.Text = ""
            txtPatientName.Text = ""
            txtFileNo.Text = ""
            txtinchflag.Text = ""
            cboFabric.Text = ""
            cboPressure.Text = ""
            txtDiagnosis.Text = ""
            cboFlaps.Text = ""
            txtStrapLength.Text = ""
            g_sFlapLength = ""
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sFlapChk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sFlapChk = ""
            txtTapeLent.Text = ""
            g_sPressureChange = "False"
            g_sFabricChange = "False"
            txtWeight.Text = ""
            txtTapeMM.Text = ""
            txtReduction.Text = ""
            txtThumLen.Text = ""
            txtThumbDist.Text = ""
            txtTCircum.Text = ""
            txtDetach.Text = ""
            txtPalmNo.Text = ""
            txtWristNo.Text = ""
            txtNoThumb.Text = ""
            txtPalmWristDist.Text = ""
            txtStump.Text = ""
            txtGauntletExtension.Text = ""
            txtFrontStrapLength.Text = ""
            txtFlap.Text = ""
            txtCustFlapLength.Text = ""
            txtFabric.Text = ""
            txtSleeve.Text = ""
            txtMM.Text = ""
            txtType.Text = ""
            txtFirstTape.Text = ""
            txtSecondTape.Text = ""
            txtSecondLastTape.Text = ""
            txtLastTape.Text = ""
            txtSex.Text = ""
            txtAge.Text = ""
            txtModulus.Text = ""
            txtWristPleat1.Text = ""
            txtWristPleat2.Text = ""
            txtShoulderPleat1.Text = ""
            txtShoulderPleat2.Text = ""
            chkNone.Checked = True
            chkStump.Checked = False
            frmGauntlet.Enabled = 0
            chkGauntlets.Checked = False
            chkDetachable.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkEnclosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            chkNoThumb.CheckState = System.Windows.Forms.CheckState.Unchecked
            txtPalmWrist.Text = ""
            txtCircum.Text = ""
            txtThumb.Text = ""
            txtWaistCir.Text = ""
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            i = 0
            txtUidMPD.Text = ""
            txtUidAC.Text = ""
            'UPGRADE_WARNING: Couldn't resolve default property of object txtACwo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtACwo = ""
            txtUidArm.Text = ""

            'initalise globals
            ARMDIA1.g_sFileNo = ""
            ARMDIA1.g_sSide = ""
            ARMDIA1.g_sPatient = ""
            ARMDIA1.g_sFabric = ""
            g_sFabnam = ""
            g_sWpleat1 = ""
            g_sWpleat2 = ""
            g_sSpleat1 = ""
            g_sSpleat2 = ""
            g_sFlapStrap = ""
            g_sFlap = ""
            g_sMM = ""
            g_sGaunt = ""
            g_sDetGaunt = ""
            g_sNoThumb = ""
            g_sPalmNo = ""
            g_sWristNo = ""
            g_sPalmWristDist = ""
            g_sThumbCircum = ""
            g_sThumbLength = ""
            g_sEnclosedThumb = ""
            g_sSecondLastTape = ""
            g_sSecondTape = ""
            g_sFirstTape = ""
            g_sLastTape = ""
            g_sModulus = ""
            g_sStump = ""
            g_sTapeLengths = ""
            g_sTapeMMs = ""
            g_sGrams = ""
            g_sFlapLength = ""
            g_sReduction = ""
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sAmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sAmm = ""
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sBmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sBmm = ""
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sCmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sCmm = ""
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sDmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sDmm = ""

            ARMDIA1.g_iStyleLastTape = -1
            ARMDIA1.g_iStyleFirstTape = -1
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sWaistCir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sWaistCir = ""

            ARMDIA1.g_sPathJOBST = fnPathJOBST()

            Dim textline As Object
            'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            fileNum = FreeFile()
            'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Dim sSettingsPath As String = fnGetSettingsPath("LookupTables")
            FileOpen(fileNum, sSettingsPath & "\Pressure.dat", Microsoft.VisualBasic.OpenMode.Input)
            ''FileOpen(fileNum, "C:\ACAD2018_SMITH-NEPHEW\Lookup Tables\Pressure.dat", OpenMode.Input)
            'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Do While Not EOF(fileNum)
                'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                textline = LineInput(fileNum)
                'UPGRADE_WARNING: Couldn't resolve default property of object textline. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                cboPressure.Items.Add(textline)
            Loop
            FileClose()


            'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            fileNum = FreeFile()
            'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(fileNum, sSettingsPath & "\fabric.dat", Microsoft.VisualBasic.OpenMode.Input)
            'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Do While Not EOF(fileNum)
                'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                textline = LineInput(fileNum)
                'UPGRADE_WARNING: Couldn't resolve default property of object textline. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                cboFabric.Items.Add(textline)
            Loop
            FileClose()

            'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            fileNum = FreeFile()
            'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(fileNum, sSettingsPath & "\flaps.dat", Microsoft.VisualBasic.OpenMode.Input)
            'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Do While Not EOF(fileNum)
                'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                textline = LineInput(fileNum)
                'UPGRADE_WARNING: Couldn't resolve default property of object textline. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                cboFlaps.Items.Add(textline)
            Loop
            FileClose()

            '
            'Load Palm Tapes and Wrist
            '
            Dim g_sTextList As String = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"

            For ii = 0 To 9
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                cboWrist.Items.Add(LTrim(Mid(g_sTextList, (ii * 3) + 1, 3)))
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                cboPalm.Items.Add(LTrim(Mid(g_sTextList, (ii * 3) + 1, 3)))
            Next ii

            cboDistalTape.Items.Add("1St")
            For ii = 0 To 17
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                cboDistalTape.Items.Add(LTrim(Mid(g_sTextList, (ii * 3) + 1, 3)))
            Next ii

            cboProximalTape.Items.Add("Last")
            For ii = 17 To 0 Step -1
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                cboProximalTape.Items.Add(LTrim(Mid(g_sTextList, (ii * 3) + 1, 3)))
            Next ii

            '   ------------------------------------------------
            '   Routine to read in the Bobinnet conversion Chart
            '   ------------------------------------------------
            Dim j, offset As Object
            Dim Modulus, Reduction As String
            'Static bob(1 To 19, 1 To 23) As Variant
            'UPGRADE_WARNING: Lower bound of array indi was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            Dim indi(19) As Object
            'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            fileNum = FreeFile()
            'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(fileNum, sSettingsPath & "\bobinnet.dat", Microsoft.VisualBasic.OpenMode.Input)
            ' For i = 1 To 19
            i = 1
            Do While Not EOF(fileNum)
                If (i <= 19) Then
                    textline = LineInput(fileNum)
                    Modulus = textline.ToString().Substring(0, 5).Trim()
                    Reduction = textline.ToString().Substring(6).Trim()
                    ''    'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ''    Input(fileNum, Modulus)
                    ''    Input(fileNum, Reduction)
                    ''    'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ''    'UPGRADE_WARNING: Couldn't resolve default property of object indi(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    indi(i) = Modulus
                    ''    'UPGRADE_WARNING: Couldn't resolve default property of object offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    offset = 1
                    Dim sString() As String = Reduction.Replace("  ", " ").Trim().Split(" ")
                    Dim colReduction As List(Of String) = New List(Of String)(sString)
                    j = 1
                    For Each tempa As String In colReduction
                        If j <= 23 Then
                            If (tempa.Length > 1) Then
                                Bob(i, j) = tempa
                                j += 1
                            End If
                        End If
                    Next
                    'For j = 1 To 23
                    'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Bob(i, j). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'Bob(i, j) = Mid(Reduction, offset, 4)
                    'UPGRADE_WARNING: Couldn't resolve default property of object offset. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'offset = offset + 4
                    'Next j
                    i += 1
                End If
            Loop ' Next i
            FileClose()
            '   ------------------------------------------------
            '            End Of Bobinnet Routine
            '   ------------------------------------------------

            '   ------------------------------------------------
            '   Routine to read in the Powernet conversion Chart
            '   ------------------------------------------------
            Dim q, P, offset2 As Object
            Dim modulus2, Reduction2 As String
            '    ReDim pow(1 To 19, 1 To 23) As Variant
            'UPGRADE_WARNING: Lower bound of array pindi was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
            Dim pindi(19) As Object
            'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            fileNum = FreeFile()
            'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FileOpen(fileNum, sSettingsPath & "\powernet.dat", Microsoft.VisualBasic.OpenMode.Input)

            P = 1
            Do While Not EOF(fileNum) ''For P = 1 To 19
                textline = LineInput(fileNum)
                modulus2 = textline.ToString().Substring(0, 5).Trim()
                Reduction2 = textline.ToString().Substring(6).Trim()
                ''    'UPGRADE_WARNING: Couldn't resolve default property of object fileNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ''    Input(fileNum, modulus2)
                ''    Input(fileNum, Reduction2)
                ''    'UPGRADE_WARNING: Couldn't resolve default property of object P. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ''    'UPGRADE_WARNING: Couldn't resolve default property of object pindi(P). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                pindi(P) = modulus2
                ''    'UPGRADE_WARNING: Couldn't resolve default property of object offset2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                offset2 = 1

                Dim sString() As String = Reduction2.Replace("  ", " ").Trim().Split(" ")
                Dim colReduction2 As List(Of String) = New List(Of String)(sString)
                q = 1
                For Each tempa As String In colReduction2
                    If (tempa.Length > 1) Then
                        Pow(P, q) = tempa
                        q += 1
                    End If
                Next

                'For q = 1 To 23
                '    'UPGRADE_WARNING: Couldn't resolve default property of object q. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                '    'UPGRADE_WARNING: Couldn't resolve default property of object P. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                '    'UPGRADE_WARNING: Couldn't resolve default property of object offset2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                '    'UPGRADE_WARNING: Couldn't resolve default property of object Pow(P, q). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                '    Pow(P, q) = Mid(Reduction2, offset2, 4)
                '    'UPGRADE_WARNING: Couldn't resolve default property of object offset2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                '    offset2 = offset2 + 4
                'Next q
                P += 1
            Loop            ''Next P
            FileClose()
            '   ------------------------------------------------
            '            End Of Powernet Routine
            '   ------------------------------------------------

            'Start a timer
            'The Timer is disabled in LinkClose
            'If after 10 seconds the timer event will "End" the programme
            'This ensures that the dialogue dies in event of a failure
            'on the drafix macro side
            Timer1.Interval = 10000 'Approx 10 Seconds
            Timer1.Enabled = True

            Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
            Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
            Dim obj As New BlockCreation.BlockCreation
            Dim blkId As ObjectId = New ObjectId()
            blkId = obj.LoadBlockInstance()
            If (blkId.IsNull()) Then
                MsgBox("Can't find Patient Details", 48, "Arm Details Dialog")
                g_bIsClose = True
                Me.Close()
                Exit Sub
            End If

            obj.BindAttributes(blkId, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units)
            If fileNo = "" Then
                MsgBox("Please enter Patient Details", 48, "Arm Details Dialog")
                g_bIsClose = True
                Me.Close()
                Exit Sub
            End If
            txtDiagnosis.Text = diagnosis
            txtFileNo.Text = fileNo
            txtPatientName.Text = patient
            txtinchflag.Text = units
            txtAge.Text = age
            txtSex.Text = sex
            txtWorkOrder.Text = workOrder

            ''txtDiagnosis1.Text = diagnosis
            txtAge1.Text = age
            txtSex1.Text = sex
            txtWorkOrder1.Text = workOrder
            txtDesigner.Text = tempEng
            txtTempDate.Text = tempDate

            cboFabric.Text = txtFabric.Text
            ARMDIA1.g_sUnits = txtinchflag.Text
            '27-06-2018
            'If (rdoLeft.Checked = True) Then
            '    txtSleeve.Text = "Left"
            'Else
            '    txtSleeve.Text = "Right"
            'End If
            txtSleeve.Text = "Left"
            readDWGInfo()
            Form_LinkClose()
            'Pressure
            If txtMM.Text.Equals("") = False Then
                cboPressure.Text = txtMM.Text
            End If
            If cboFabric.Text.Equals("") = False Then
                PR_GetFabric()
            End If
            g_sLeftArmValues = FN_LeftArmValuesString()
            g_sRightArmValues = FN_RightArmValuesString()
        Catch ex As Exception
            g_bIsClose = True
            Me.Close()
        End Try
    End Sub

    Private Sub Gauntlet(ByRef NoThumb As Object, ByRef Palm As Object, ByRef Wrist As Object, ByRef PalmWrist As Object, ByRef thumlength As Object, ByRef Sex As Object, ByRef age As Object, ByRef Enclosed As Object, ByRef Circum As Object, ByRef Detach As Object, ByRef txtinchflag As Object)
        Dim tempCircum As Object
        Dim txtThumblen As Object
        Dim thum As Object
        Dim thumdist As Object
        Dim temppalm As Object
        Dim tempwrist As Object

        txtGauntletFlag.Text = "True"
        'UPGRADE_WARNING: Couldn't resolve default property of object Wrist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object tempwrist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tempwrist = Wrist
        'UPGRADE_WARNING: Couldn't resolve default property of object Palm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object temppalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        temppalm = Palm
        'UPGRADE_WARNING: Couldn't resolve default property of object PalmWrist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object thumdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        thumdist = PalmWrist
        'UPGRADE_WARNING: Couldn't resolve default property of object thumdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If thumdist = "" Then txtPalmWrist.Text = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object thumdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        thumdist = cmtoinches(txtinchflag, txtPalmWrist)

        'No Thumb
        If chkNoThumb.CheckState = 1 Then
            txtNoThumb.Text = "True"
        Else
            txtNoThumb.Text = "False"
        End If

        ' Check values entered
        'UPGRADE_WARNING: Couldn't resolve default property of object Palm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Palm = "" Then
            MsgBox("Palm tape number has Not been entered.", 48, "Gauntlet Details")
            Exit Sub
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object Wrist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Wrist = "" Then
            MsgBox("Wrist tape number has Not been entered.", 48, "Gauntlet Details")
            Exit Sub
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object PalmWrist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If PalmWrist = "" Then
            MsgBox("Palm To Wrist Distance has Not been entered.", 48, "Gauntlet Details")
            Exit Sub
        End If
        ' Thumbhole Dist
        If thumdist > 2.5 Then
            txtThumbDist.Text = CStr(2.5)
            'UPGRADE_WARNING: Couldn't resolve default property of object thumdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtPalmWristDist.Text = thumdist
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object thumdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtThumbDist.Text = thumdist
            'UPGRADE_WARNING: Couldn't resolve default property of object thumdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtPalmWristDist.Text = thumdist
        End If
        ' End Thumbhole Dist

        'Thumb length
        'UPGRADE_WARNING: Couldn't resolve default property of object thumlength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(thumlength) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object thumlength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object thum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            thum = thumlength
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object thum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            thum = cmtoinches(txtinchflag, txtThumb)
            'UPGRADE_WARNING: Couldn't resolve default property of object thum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtThumLen.Text = thum
        Else : txtThumLen.Text = CStr(0)
            'UPGRADE_WARNING: Couldn't resolve default property of object thum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            thum = 0
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object age. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object thum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If thum = 0 And Val(age) > 14 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object thum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            thum = 0.75
            'UPGRADE_WARNING: Couldn't resolve default property of object thum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtThumLen.Text = thum
            'UPGRADE_WARNING: Couldn't resolve default property of object age. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf CDbl(txtThumLen.Text) = 0 And Val(age) <= 14 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object thum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            thum = 0.5
            'UPGRADE_WARNING: Couldn't resolve default property of object thum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtThumLen.Text = thum
        End If
        'End Thumb length

        'Enclosed Thumb
        'UPGRADE_WARNING: Couldn't resolve default property of object age. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Enclosed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Enclosed = 1 And Val(age) > 14 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object thumlength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If thumlength = "" Then
                MsgBox("Thumb Length has Not been entered.", 48, "Gauntlet Details")
                Exit Sub
            End If
            txtEnclosed.Text = "True"
            'UPGRADE_WARNING: Couldn't resolve default property of object thumlength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object txtThumblen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtThumblen = Val(thumlength) - 0.25
            'UPGRADE_WARNING: Couldn't resolve default property of object age. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Enclosed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf Enclosed = 1 And Val(age) < 15 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object thumlength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If thumlength = 0 Then
                MsgBox("Thumb Length has Not been entered.", 48, "Gauntlet Details")
                Exit Sub
            End If
            txtEnclosed.Text = "True"
            'UPGRADE_WARNING: Couldn't resolve default property of object thumlength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object txtThumblen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtThumblen = Val(thumlength) - 0.125
            'UPGRADE_WARNING: Couldn't resolve default property of object Enclosed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf Enclosed = 0 Then
            txtEnclosed.Text = "False"
        End If
        'End Enclosed Thumb

        ' Thumb Circumference
        'UPGRADE_WARNING: Couldn't resolve default property of object Circum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object NoThumb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If NoThumb = 0 And Val(Circum) = 0 Then
            MsgBox("Thumb Circumference has Not been entered.", 48, "Gauntlet Details")
            Exit Sub
            'UPGRADE_WARNING: Couldn't resolve default property of object Circum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object NoThumb. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf NoThumb = 0 And Val(Circum) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object tempCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            tempCircum = cmtoinches(txtinchflag, txtCircum)
            'UPGRADE_WARNING: Couldn't resolve default property of object tempCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            tempCircum = Val(tempCircum) / 2
            'UPGRADE_WARNING: Couldn't resolve default property of object tempCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtTCircum.Text = tempCircum
        End If
        ' End Thumb Circumference

        ' Detachable
        'UPGRADE_WARNING: Couldn't resolve default property of object Detach. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Detach = 1 Then
            txtDetach.Text = "True"
        Else
            txtDetach.Text = "False"
        End If
        ' End Detachable

        'UPGRADE_WARNING: Couldn't resolve default property of object Wrist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Palm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(Palm) > Val(Wrist) Then
            MsgBox("Palm Tape No Should be less than Wrist Tape No. Enter New values And re-calculate. ", 48, "Gauntlet Details")
            Exit Sub
        End If
        ' check palm and wrist tapes

    End Sub

    Private Function GetReduction(ByRef Modul As Object, ByRef Gram As Object, ByRef fabnam As Object) As Object
        Dim tempred As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If VB.Left(fabnam, 3) = "Pow" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object PowReduction(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object tempred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            tempred = PowReduction(Modul, Gram)
            'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf VB.Left(fabnam, 3) = "Bob" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object BobReduction(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object tempred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            tempred = BobReduction(Modul, Gram)
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object GetReduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetReduction = tempred
    End Function

    Private Sub lblPleat_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblPleat.DoubleClick
        Dim Index As Short = lblPleat.GetIndex(eventSender)
        'Procedure to disable the pleat fields
        'to enable the User to select only the relevent
        'data to be drawn
        'Saves having to delele and re-enter to get variations
        'on pleats
        If System.Drawing.ColorTranslator.ToOle(lblPleat(Index).ForeColor) = &HC0C0C0 Then
            'Enable pleat if previously disabled
            lblPleat(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
            Select Case Index
                Case 0
                    txtWristPleat1.Enabled = True
                Case 1
                    txtWristPleat2.Enabled = True
                Case 3
                    txtShoulderPleat1.Enabled = True
                Case 2
                    txtShoulderPleat2.Enabled = True
            End Select
        Else
            'Disable pleat if previously enabled
            lblPleat(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            Select Case Index
                Case 0
                    txtWristPleat1.Enabled = False
                Case 1
                    txtWristPleat2.Enabled = False
                Case 3
                    txtShoulderPleat1.Enabled = False
                Case 2
                    txtShoulderPleat2.Enabled = False
            End Select
        End If

    End Sub

    Private Sub Length_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Length.Enter
        Dim Index As Short = Length.GetIndex(eventSender)
        select_text_in_Box1(Length(Index))
        ''Length(Index).Focus()
        Length(Index).SelectAll()
    End Sub

    Private Sub Length_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Length.Leave
        Dim Index As Short = Length.GetIndex(eventSender)

        'UPGRADE_WARNING: Couldn't resolve default property of object ChangeChecker(Index + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Dim inchbit As Double
        If (ChangeChecker(Index + 1) <> Length(Index).Text) Then

            If Me.Visible = True Then ARMDIA1.g_Modified = True

            ' Delete tape values
            If Length(Index).Text = "" And mms(Index).Text <> "" And gms(Index).Text <> "" And reds(Index).Text <> "" And InchText(Index).Text <> "" Then
                mms(Index).Text = ""
                gms(Index).Text = ""
                reds(Index).Text = ""
                InchText(Index).Text = ""
                PR_GetStartEndTapes()
                'Reset the first and last tape indicators
                If cboProximalTape.SelectedIndex = 0 Then cboProximalTape_SelectedIndexChanged(cboProximalTape, New System.EventArgs())
                If cboDistalTape.SelectedIndex = 0 Then cboDistalTape_SelectedIndexChanged(cboDistalTape, New System.EventArgs())

            End If


            ' Check for numeric values
            If Not FN_CheckValue(Length(Index), "Tape") Then Exit Sub


            ' check for too large a number
            If (Val(Length(Index).Text) > 99 And txtinchflag.Text = "cm") Or (Val(Length(Index).Text) > 50 And txtinchflag.Text = "inches") Then
                MsgBox("Tape length entered Is too Large", 48, "Arm Details")
                Length(Index).Focus()
                Exit Sub
            End If

            ' check for too small a number
            ' tapelen from above
            If Val(Length(Index).Text) < 1 And Len(Length(Index).Text) > 0 Then
                MsgBox("Tape length entered Is too Small", 48, "Arm Details")
                Length(Index).Focus()
                Exit Sub
            End If


            If Val(Length(Index).Text) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                inchbit = cmtoinches(txtinchflag.Text, Length(Index))
                If inchbit = -1 Then
                    Length(Index).Focus()
                    Exit Sub
                End If
                InchText(Index).Text = ARMEDDIA1.fnInchesToText(inchbit)
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(Index + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Measurements(Index + 1) = cmtoinches(txtinchflag.Text, Length(Index))

                'Exit if no fabric or no pressure has been selected
                'We also supress calculations if this is an empty sleeve
                'the flag EmptySleeve is set in Form "LinkClose" & Command "Calculate"
                'I.E. don't do any calculations
                'UPGRADE_WARNING: Couldn't resolve default property of object EmptySleeve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If cboFabric.Text = "" Or cboPressure.Text = "" Or EmptySleeve = True Then Exit Sub

                'Recalculate etc
                PR_GetPressure()
                PR_GetFabric()
                PR_CheckValues()
                PR_GetGauntletDetails()
                PR_GetStartEndTapes()
                PR_CalculateMMSGmsReds(Index)
                PR_ReSetMMs(Index, g_sModulus)
                PR_SetRresults()
                PR_SetGlobals()


                'check for extra measurements in Gauntlet
                PR_GauntletHoleCheck()


                'If a tape has been added then reset the tape indicators
                If cboProximalTape.SelectedIndex = 0 Then cboProximalTape_SelectedIndexChanged(cboProximalTape, New System.EventArgs())
                If cboDistalTape.SelectedIndex = 0 Then cboDistalTape_SelectedIndexChanged(cboDistalTape, New System.EventArgs())

            End If
            'Reset ChangeChecker to revised value GG.17.Mar.95
            'UPGRADE_WARNING: Couldn't resolve default property of object ChangeChecker(Index + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ChangeChecker(Index + 1) = Length(Index).Text
        End If
    End Sub

    Private Sub LowLength(ByRef tapes As Object)
        Dim i, nValue As Object
        For i = 0 To 17
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object tapes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nValue = Val(Mid(DirectCast(tapes, TextBox).Text, (i * 3) + 1, 3)) / 10
            'UPGRADE_WARNING: Couldn't resolve default property of object nValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If nValue > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object nValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Length(i).Text = nValue
            End If
        Next i
    End Sub

    Private Sub LowRed(ByRef tapes As Object)
        Dim i, nValue As Object
        For i = 0 To 17
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object tapes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nValue = Val(Mid(DirectCast(tapes, TextBox).Text, (i * 3) + 1, 3))
            'UPGRADE_WARNING: Couldn't resolve default property of object nValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If nValue > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object nValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                reds(i).Text = nValue
            End If
        Next i
    End Sub

    Private Sub LowTapeMM(ByRef tapes As Object)
        Dim i, nValue As Object
        For i = 0 To 17
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object tapes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nValue = Val(Mid(DirectCast(tapes, TextBox).Text, (i * 3) + 1, 3))
            'UPGRADE_WARNING: Couldn't resolve default property of object nValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If nValue > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object nValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mms(i).Text = nValue
            End If
        Next i
    End Sub

    Private Sub LowWeight(ByRef tapes As Object)
        Dim i, nValue As Object
        For i = 0 To 17
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object tapes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object nValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nValue = Val(Mid(DirectCast(tapes, TextBox).Text, (i * 3) + 1, 3))
            'UPGRADE_WARNING: Couldn't resolve default property of object nValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If nValue > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object nValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                gms(i).Text = nValue
            End If
        Next i
    End Sub

    Private Function MissingVal(ByRef Before As Object, ByRef Present As Object, ByRef after As Object) As Object
        '   *** Function to Check For Missing Measurements
        'UPGRADE_WARNING: Couldn't resolve default property of object after. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Before. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Present. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Present = "" And Before <> "" And after <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object MissingVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            MissingVal = 1
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object MissingVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            MissingVal = 0
        End If
    End Function

    Private Sub mms_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mms.Enter
        Dim Index As Short = mms.GetIndex(eventSender)
        select_text_in_Box1(mms(Index))
        'UPGRADE_WARNING: Couldn't resolve default property of object MMsChecker(Index + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        MMsChecker(Index + 1) = mms(Index).Text
    End Sub

    Private Sub mms_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mms.Leave
        Dim Index As Short = mms.GetIndex(eventSender)
        Dim Backgram As Object

        If mms(Index).Text <> "" And Length(Index).Text = "" And gms(Index).Text = "" Then
            mms(Index).Text = ""
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object MMsChecker(Index + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Dim inchbit As Double
        Dim X, w, Gram As Object
        If MMsChecker(Index + 1) <> mms(Index).Text Then

            If Me.Visible = True Then ARMDIA1.g_Modified = True

            If Length(Index).Text <> "" And cboFabric.Text <> "" And cboPressure.Text <> "" And gms(Index).Text <> "" Then

                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                inchbit = cmtoinches(txtinchflag.Text, Length(Index))
                If inchbit = -1 Then
                    Length(Index).Focus()
                    Exit Sub
                End If
                InchText(Index).Text = ARMEDDIA1.fnInchesToText(inchbit)
                PR_GetFabric()
                PR_GetStartEndTapes()
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                w = Index
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                X = Index + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Gram = Fix(Val(mms(Index).Text) * Val(Measurements(Index + 1)))
                'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                gms(Index).Text = CStr(ARMDIA1.round(Gram))
                'UPGRADE_WARNING: Couldn't resolve default property of object GetReduction(g_sModulus, Gram, g_sFabnam). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GetReduction(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                reds(Index).Text = GetReduction(g_sModulus, Gram, g_sFabnam)

                'Check for Grams Being Too Big
                If Val(reds(Index).Text) = 0 And CDbl(gms(Index).Text) > 205 Then
                    reds(Index).Text = CStr(32)
                    If g_sFabnam = "Pow" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object Pow(txtModulus, 23). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Backgram = Pow(CInt(txtModulus.Text), 23)
                    ElseIf g_sFabnam = "Bob" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object Bob(txtModulus, 23). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Backgram = Bob(CInt(txtModulus.Text), 23)
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    gms(Index).Text = Backgram
                    'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    mms(Index).Text = CStr(ARMDIA1.round(Backgram / Val(Measurements(Index + 1))))
                End If

                'Check for grams being too little
                If Val(reds(Index).Text) <= 10 Then
                    reds(Index).Text = CStr(10)
                    If g_sFabnam = "Pow" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object Pow(txtModulus, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Backgram = Pow(CInt(txtModulus.Text), 1)
                    ElseIf g_sFabnam = "Bob" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object Bob(txtModulus, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Backgram = Bob(CInt(txtModulus.Text), 1)
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    gms(Index).Text = Backgram
                    'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    mms(Index).Text = CStr(ARMDIA1.round(Backgram / Val(Measurements(Index + 1))))
                End If

                PR_WristRed(w, X, g_sModulus)
                PR_PalmRed(w, X, g_sModulus)
                PR_LastTapeRed(w, X, g_sModulus)
                PR_Stump(w, X)
                PR_CheckFlaps(w)
                PR_ReSetMMs(Index, g_sModulus)
                PR_SetRresults()
                PR_SetGlobals()
            End If
        End If
    End Sub

    Private Function NoVals() As Object
        '   *** Function to check if no values have been entered ***
        Dim contents, i, total As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        i = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object contents. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        contents = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object total. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        total = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While i < 18
            'UPGRADE_WARNING: Couldn't resolve default property of object contents. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            contents = Val(Length(i).Text)
            'UPGRADE_WARNING: Couldn't resolve default property of object contents. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object total. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            total = total + contents
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            i = i + 1
        Loop
        'UPGRADE_WARNING: Couldn't resolve default property of object total. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object NoVals. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        NoVals = total
        '    *** End Check values have been entered ***
    End Function

    Private Sub OldGauntlet(ByRef txtPalmNo As Object, ByRef txtWristNo As Object, ByRef txtPalmWristDist As Object, ByRef txtTCircum As Object, ByRef txtThumLen As Object)
        Dim txtWrist As Object
        Dim txtPalm As Object
        Dim differ As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object txtPalmNo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object txtWristNo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object differ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        differ = Val(txtWristNo) - Val(txtPalmNo)
        'UPGRADE_WARNING: Couldn't resolve default property of object differ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If differ > 1 Then
            g_sHoleCheck = "True"
            'UPGRADE_WARNING: Couldn't resolve default property of object differ. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf differ = 1 Then
            g_sHoleCheck = "False"
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object txtPalmNo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object txtPalm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtPalm = (Val(txtPalmNo) * 1.5) + (-7.5)
        'UPGRADE_WARNING: Couldn't resolve default property of object txtWristNo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object txtWrist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtWrist = (Val(txtWristNo) * 1.5) + (-7.5)
        'UPGRADE_WARNING: Couldn't resolve default property of object txtPalmWristDist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtPalmWrist.Text = txtPalmWristDist
        'UPGRADE_WARNING: Couldn't resolve default property of object txtTCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtCircum.Text = txtTCircum
        'UPGRADE_WARNING: Couldn't resolve default property of object txtThumLen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        txtThumb.Text = txtThumLen
    End Sub

    'UPGRADE_WARNING: Event optProximalTape.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optProximalTape_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optProximalTape.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optProximalTape.GetIndex(eventSender)
            'only two choices
            '
            If Index = 0 Then
                'Disable flaps
                frmFlaps.Enabled = False
                'cboFlaps.Text = "none"
                cboFlaps.SelectedIndex = -1
                cboFlaps_SelectedIndexChanged(cboFlaps, New System.EventArgs()) 'Force change of pressure
                cboFlaps.Enabled = False
                txtStrap.Text = ""
                txtFlap.Text = ""

                lblStyle.Enabled = False
                txtStrapLength.Text = ""
                txtStrapLength.Enabled = False
                labStrap.Text = ""
                lblStrapLength.Enabled = False

                txtFrontStrapLength.Text = ""
                txtFrontStrapLength.Enabled = False
                labFrontStrapLength.Text = ""
                lblFrontStrapLength.Enabled = False

                txtCustFlapLength.Text = ""
                txtCustFlapLength.Enabled = False
                labCustFlapLength.Text = ""
                lblCustFlapLength.Enabled = False

                txtWaistCir.Text = ""
                txtWaistCir.Enabled = False
                labWaistCir.Text = ""
                lblWaistCir.Enabled = False

                'Enable proximal tape position
                cboProximalTape.Enabled = True

            Else
                'Reset proximal tape position to last
                cboProximalTape.SelectedIndex = 0
                cboProximalTape_SelectedIndexChanged(cboProximalTape, New System.EventArgs())
                cboProximalTape.Enabled = False

                'Enable flaps
                frmFlaps.Enabled = True
                frmFlaps.Enabled = True
                cboFlaps.Enabled = True
                lblStyle.Enabled = True

                txtStrapLength.Enabled = True
                lblStrapLength.Enabled = True

                txtCustFlapLength.Enabled = True
                lblCustFlapLength.Enabled = True

                txtFrontStrapLength.Enabled = True
                lblFrontStrapLength.Enabled = True

                'Can't have a detachable gauntlet with a flap
                If chkDetachable.CheckState = 1 Then chkDetachable.CheckState = System.Windows.Forms.CheckState.Unchecked

                '        txtWaistCir.Enabled = True
                '        lblWaistCir.Enabled = True
            End If

            If Me.Visible = True Then ARMDIA1.g_Modified = True

        End If
    End Sub

    Private Function PadGram(ByRef GramTxt As Object) As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Dim Gram As String
        If GramTxt.GetType Is GetType(TextBox) Then
            Gram = DirectCast(GramTxt, TextBox).Text
        ElseIf GramTxt.GetType Is GetType(Label) Then
            Gram = DirectCast(GramTxt, Label).Text
        Else
            Gram = GramTxt
        End If

        If Gram <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Val(Gram) < 10 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PadGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PadGram = "00" & Gram
                'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf Val(Gram) < 100 And Val(Gram) > 9 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PadGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PadGram = "0" & Gram
                'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf Val(Gram) > 99 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PadGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PadGram = Gram
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf Gram = "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object PadGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PadGram = "000"
        End If

    End Function

    Private Function PadMM(ByRef mmTxt As Object) As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object mm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Dim mm As String
        If mmTxt.GetType() Is GetType(TextBox) Then
            mm = DirectCast(mmTxt, TextBox).Text
        Else
            mm = mmTxt
        End If
        If mm <> "" Then
            If Len(mm) = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object mm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PadMM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PadMM = "00" & mm
            ElseIf Len(mm) = 2 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object mm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PadMM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PadMM = "0" & mm
            ElseIf Len(mm) = 3 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object mm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PadMM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PadMM = mm
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object mm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf mm = "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object PadMM. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PadMM = "000"
        End If
    End Function

    Private Function PadRed(ByRef redTxt As Object) As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object red. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Dim red As String
        If redTxt.GetType() Is GetType(TextBox) Then
            red = DirectCast(redTxt, TextBox).Text
        ElseIf redTxt.GetType() Is GetType(Label) Then
            red = DirectCast(redTxt, Label).Text
        Else
            red = redTxt
        End If

        If red <> "" Then
            If Len(red) = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object red. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PadRed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PadRed = "00" & red
            ElseIf Len(red) = 2 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object red. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PadRed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PadRed = "0" & red
            ElseIf Len(red) = 3 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object red. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PadRed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PadRed = red
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object red. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf red = "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object PadRed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PadRed = "000"
        End If
    End Function

    Private Function PadTape(ByRef measure As Object) As Object
        Dim bit, pos, dig, Result As Object
        If Len(measure) = 4 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object measure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            pos = InStr(measure, ".")
            'UPGRADE_WARNING: Couldn't resolve default property of object pos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If pos <> 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dig. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dig = Int(CDbl(measure))
                'UPGRADE_WARNING: Couldn't resolve default property of object measure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object bit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                bit = VB.Right(measure, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object bit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dig. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Result = dig & bit
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object PadTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PadTape = Result
        ElseIf Len(measure) = 3 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object measure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            pos = InStr(measure, ".")
            'UPGRADE_WARNING: Couldn't resolve default property of object pos. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If pos <> 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object dig. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dig = Int(CDbl(measure))
                'UPGRADE_WARNING: Couldn't resolve default property of object measure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object bit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                bit = VB.Right(measure, 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object bit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object dig. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Result = "0" & dig & bit
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object PadTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PadTape = Result
        ElseIf Len(measure) = 2 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object measure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Result = measure
            'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object PadTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PadTape = Result & "0"
        ElseIf Len(measure) = 1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object measure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Result = measure
            'UPGRADE_WARNING: Couldn't resolve default property of object Result. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object PadTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PadTape = "0" & Result & "0"
            'UPGRADE_WARNING: Couldn't resolve default property of object measure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf measure = "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object PadTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PadTape = "000"
        End If
    End Function

    Private Function PowReduction(ByRef Modul As Object, ByRef lblGram As Object) As Object
        Dim lblred As Object
        Dim j As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        one = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object two. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        two = 2
        'UPGRADE_WARNING: Couldn't resolve default property of object nine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nine = 9
        'UPGRADE_WARNING: Couldn't resolve default property of object eight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        eight = 8
        'UPGRADE_WARNING: Couldn't resolve default property of object ten. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ten = 10
        'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        j = one
        'UPGRADE_WARNING: Couldn't resolve default property of object lblGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Pow(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (Val(Pow(Modul, 1)) = Val(lblGram)) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object ten. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            lblred = ten
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object two. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        j = two
        'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While j <= 22
            'UPGRADE_WARNING: Couldn't resolve default property of object lblGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Pow(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (Val(Pow(Modul, j)) = Val(lblGram)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object nine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                lblred = Val(j + nine)
                'UPGRADE_WARNING: Couldn't resolve default property of object lblGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Pow(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (Val(lblGram) - Val(Pow(Modul, j - one))) <= (Val(Pow(Modul, j)) - Val(lblGram)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object eight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                lblred = Val(j + eight)
                Exit Do
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            j = j + one
        Loop
        'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object PowReduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PowReduction = lblred
    End Function

    Private Sub PR_AddMarker(ByRef xyInsert As ARMDIA1.XY)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''-----------------PrintLine(ARMDIA1.fNum, "hEnt = AddEntity(" & ARMDIA1.QQ & "marker" & ARMDIA1.QCQ & "xmarker" & ARMDIA1.QC)
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''--------------------PrintLine(ARMDIA1.fNum, "xyStart.x+" & Str(xyInsert.X) & ARMDIA1.CC & "xyStart.y+" & Str(xyInsert.y) & ARMDIA1.CC)
        'UPGRADE_WARNING: Couldn't resolve default property of object CC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''--------------------PrintLine(ARMDIA1.fNum, 0.1 & ARMDIA1.CC & 0.1 & ARMDIA1.CC & 0 & ");")
    End Sub

    Private Sub PR_CalculateMMSGmsReds(ByRef Index As Short)
        Dim w, X As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        w = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While X < 19
            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Do While w < 18

                PR_InitaliseEmpty(w) 'Zero any empty fields
                PR_InitaliseArrays(w, X)
                If (Length(Index).Text <> "" And mms(Index).Text = "") Or (Length(Index).Text <> "" And g_sPressureChange = "True") Then
                    PR_SetMMs(w, g_sAmm, g_sBmm, g_sCmm, g_sDmm)
                End If
                PR_Reductions(w, X)
                PR_WristRed(w, X, g_sModulus)
                PR_PalmRed(w, X, g_sModulus)
                PR_Stump(w, X)
                PR_LastCircum(w, X)
                PR_LastTapeRed(w, X, g_sModulus)
                PR_CheckFlaps(w)
                PR_Detachable(w, X, g_sModulus, g_sFabnam)
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                w = w + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                X = X + 1
            Loop
        Loop

    End Sub

    Private Sub PR_CheckFlaps(ByRef w As Object)
        If Length(w).Text <> "" Then
            ' Check Flaps
            txtFlap.Text = cboFlaps.Text
            g_sFlap = txtFlap.Text
            If txtFlap.Text <> "" And txtStrapLength.Text = "" Then
                If txtinchflag.Text = "cm" Then
                    txtStrap.Text = "61"
                Else
                    txtStrap.Text = "24"
                End If
                'labStrap = fnInchesToText(cmtoinches(txtinchflag, txtStrap))
            End If
        End If
        ' End Check Flaps
    End Sub

    Private Sub PR_CheckValues()
        ''   *** Check values have been entered ***
        Dim contents, i, total As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        i = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object contents. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        contents = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object total. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        total = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While i < 18
            'UPGRADE_WARNING: Couldn't resolve default property of object contents. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            contents = Val(Length(i).Text)
            'UPGRADE_WARNING: Couldn't resolve default property of object contents. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object total. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            total = total + contents
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            i = i + 1
        Loop
        'UPGRADE_WARNING: Couldn't resolve default property of object total. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If total = 0 Then
            MsgBox("No tape measurements have been entered!", 48, "Arm Details")
            Exit Sub
        End If
        '   *** End Check values have been entered ***

    End Sub

    Private Sub PR_Detachable(ByRef w As Object, ByRef X As Object, ByRef Modul As Object, ByRef fabnam As Object)
        Dim Backgram As Object
        If Length(w).Text <> "" Then
            'CheckDetachable Gauntlet
            If chkDetachable.CheckState = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If X = Val(txtWristNo.Text) And X = Val(txtSecondLastTape.Text) Then
                    reds(w).Text = CStr(10)
                    'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If VB.Left(fabnam, 3) = "Pow" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modul, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Backgram = Pow(Modul, 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If Val(gms(w).Text) <> Backgram Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            gms(w).Text = CStr(ARMDIA1.round(Backgram))
                            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                        End If
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modul, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Backgram = Bob(Modul, 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If Val(gms(w).Text) <> Backgram Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            gms(w).Text = CStr(ARMDIA1.round(Backgram))
                            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                        End If
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf X = Val(txtWristNo.Text) + 1 And X = Val(txtLastTape.Text) Then
                    'If length of wrist is same as wrist+1 then assume detachable gauntlet
                    'originally based on 2 tapes
                    'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If Length(w).Text = Length(w - 1).Text Then
                        reds(w).Text = CStr(5)
                        gms(w).Text = ""
                        mms(w).Text = ""
                    Else
                        reds(w).Text = CStr(10)
                        'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If VB.Left(fabnam, 3) = "Pow" Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modul, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Backgram = Pow(Modul, 1)
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If Val(gms(w).Text) <> Backgram Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                gms(w).Text = CStr(ARMDIA1.round(Backgram))
                                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                            End If
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modul, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Backgram = Bob(Modul, 1)
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If Val(gms(w).Text) <> Backgram Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                gms(w).Text = CStr(ARMDIA1.round(Backgram))
                                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                            End If
                        End If
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf X = Val(txtWristNo.Text) And X = Val(txtLastTape.Text) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If Length(w + 1).Text = "" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Length(w + 1).Text = Length(w).Text
                        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(X + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Measurements(X + 1) = Measurements(X)
                        'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        InchText(w + 1).Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, Length(w)))
                        reds(w).Text = CStr(10)
                        'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If VB.Left(fabnam, 3) = "Pow" Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modul, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Backgram = Pow(Modul, 1)
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If Val(gms(w).Text) <> Backgram Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                gms(w).Text = CStr(ARMDIA1.round(Backgram))
                                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                            End If
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modul, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Backgram = Bob(Modul, 1)
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If Val(gms(w).Text) <> Backgram Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                gms(w).Text = CStr(ARMDIA1.round(Backgram))
                                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                            End If
                        End If
                    Else
                        reds(w).Text = CStr(5)
                        gms(w).Text = ""
                        mms(w).Text = ""
                    End If
                End If
            End If
        End If


    End Sub

    Public Sub PR_DrawShoulderFlaps(ByRef Profile As ARMDIA1.curve, ByRef numtapes As Object)
        Dim Strap As Object
        Dim RagAng5 As Object
        Dim RagAng4 As Object
        Dim PhiExtra As Object
        Dim ii As Object
        Dim nNotchToTangent As Object
        Dim nTemplateAngle As Object
        Dim nTemplateRadius As Object
        'UPGRADE_WARNING: Arrays in structure RaglanFlap may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        'UPGRADE_WARNING: Arrays in structure ShoulderFlap may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
        Dim ShoulderFlap, RaglanFlap As ARMDIA1.curve
        Dim LastTapeValue As Object
        Dim xyPt1, xyPt2 As ARMDIA1.XY
        Dim xyTopFlap4, xyTopFlap2, xyTopFlap1, xyTopFlap3, xyTopFlap5 As ARMDIA1.XY
        Dim xyTopFlap9, xyTopFlap7, xyTopFlap6, xyTopFlap8, xyTopFlap10 As ARMDIA1.XY
        Dim xyBotFlap4, xyBotFlap2, xyBotFlap1, xyBotFlap3, xyBotFlap5 As ARMDIA1.XY
        Dim xyPotA, xyFlap4, xyFlap2, xyFlap1, xyFlap3, xyFlap5, xyPotB As ARMDIA1.XY
        Dim Alpha, BottomRadius, Delta, Omega, Beta, Phi, Theta, PI, TopArcIncrement As Object
        Dim TopRadius, LittleBit, opp, h1, x1, y1, adj, hyp, Change, FlapMarkerAngle As Object
        Dim xyFlapText1 As ARMDIA1.XY
        Dim xyFlapText2, xyCentre, xyMid, xyFlapMarker, xyStrapText As ARMDIA1.XY
        Dim FlapLength, FlapMarkerLength As Object
        Dim TopArcAngle, TopArcRadius, TopArcLength, ArcLength, CircleCircum, BottomArcLength, BottomArcRadius, BottomArcAngle As Object
        Dim xfablen, TempMarker, xfabby, xfabdist As Object
        Dim Arrow As String
        Dim xyRaglan6, xyRaglan4, xyRaglan2, xyRaglan1, xyRaglan3, xyRaglan5, xyRaglan7 As ARMDIA1.XY
        Dim TopRagLanAngle, RaglanBottom, RagAng2, Rag4, Rag2, Rag1, Rag3, RagAng1, RagAng3, RaglanTip, InnerRaglanTip As Object
        Dim nNotchOffset As Double
        'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        BottomArcAngle = 47.5
        Arrow = "closed arrow"
        'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PI = 4 * System.Math.Atan(1)
        'If txtFlap <> "" And txtFlap <> "Vest Raglan" Then
        If txtFlap.Text <> "" And txtFlap.Text <> "Vest" And txtFlap.Text <> "Body" Then
            ' Check for Custom Flap Length
            If Val(txtCustFlapLength.Text) > 0 Then
                g_sinchflag = txtinchflag.Text
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                g_sFlapLength = cmtoinches(txtinchflag.Text, txtCustFlapLength)
                'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                LastTapeValue = g_sFlapLength
            Else
                'Standard flap length
                'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                LastTapeValue = Val(txtLastTape.Text) - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                LastTapeValue = cmtoinches(txtinchflag.Text, Length(LastTapeValue))
                'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                LastTapeValue = (LastTapeValue / 3.14) * 0.92
            End If


            ' Calculate Bottom part of curve with 5 Points
            If (Profile.y(Profile.n)) >= 5.5 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                BottomRadius = LastTapeValue - 0.625
                'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RaglanBottom = 3.5
                'TopRagLanAngle = 20
                'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                TopRagLanAngle = 17
                'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RaglanTip = 1.9375
                'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                InnerRaglanTip = 2.173
                nNotchOffset = 0.875 '14/16 ths Raglan Template s287
            ElseIf ((Profile.y(Profile.n)) < 5.5) And ((Profile.y(Profile.n)) > 3) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                BottomRadius = 2.8
                'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RaglanBottom = 2.5
                'TopRagLanAngle = 20
                'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                TopRagLanAngle = 17
                'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RaglanTip = 1.05
                'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                InnerRaglanTip = 1.22
                nNotchOffset = 0.75 '12/16 ths Raglan Template s289
            ElseIf ((Profile.y(Profile.n)) <= 3) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                BottomRadius = 2
                'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RaglanBottom = 2.0625
                'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RaglanTip = 0.625
                'TopRagLanAngle = 15
                'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                TopRagLanAngle = 15
                'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                InnerRaglanTip = 0.727
                nNotchOffset = 0.6875 '11/16 ths, Raglan Template s290
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object nTemplateRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nTemplateRadius = 3.5
            'UPGRADE_WARNING: Couldn't resolve default property of object nTemplateAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nTemplateAngle = 40
            'UPGRADE_WARNING: Couldn't resolve default property of object nNotchToTangent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            nNotchToTangent = 2.125

            'Calc length of bottom Arc
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            BottomArcRadius = BottomRadius
            'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object CircleCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            CircleCircum = PI * (2 * (BottomArcRadius))
            'UPGRADE_WARNING: Couldn't resolve default property of object CircleCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            BottomArcLength = (BottomArcAngle / 360) * CircleCircum

            'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ShoulderFlap.n = 10
            Dim X(10), Y(10) As Double
            X(1) = Profile.X(Profile.n) + LastTapeValue
            Y(1) = 0
            'ShoulderFlap.X(1) = Profile.X(Profile.n) + LastTapeValue
            'ShoulderFlap.y(1) = 0
            ShoulderFlap.X = X
            ShoulderFlap.y = Y
            For ii = 1 To 4
                'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Theta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Theta = (ii * 11.875 * PI) / 180
                'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Theta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                adj = System.Math.Cos(Theta) * BottomRadius
                'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Theta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                opp = System.Math.Sin(Theta) * BottomRadius
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object LittleBit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                LittleBit = BottomRadius - adj
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object LittleBit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ShoulderFlap.X(ii + 1) = Profile.X(Profile.n) + LastTapeValue - LittleBit
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ShoulderFlap.y(ii + 1) = 0 + opp
            Next ii
            ARMDIA1.PR_MakeXY(xyBotFlap5, ShoulderFlap.X(5), ShoulderFlap.y(5))

            'Calculate 6 angles and points for top part of curve
            ARMDIA1.PR_MakeXY(xyTopFlap1, Profile.X(Profile.n), Profile.y(Profile.n))
            ShoulderFlap.X(10) = xyTopFlap1.X
            ShoulderFlap.y(10) = xyTopFlap1.y

            'midpoint
            'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            h1 = ARMDIA1.FN_CalcLength(xyTopFlap1, xyBotFlap5)
            'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            h1 = h1 / 2
            'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Omega = ARMDIA1.FN_CalcAngle(xyBotFlap5, xyTopFlap1)
            'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Omega = 180 - Omega
            'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Omega = (Omega * PI) / 180
            'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            opp = System.Math.Abs(h1 * System.Math.Sin(Omega))
            'UPGRADE_WARNING: Couldn't resolve default property of object Omega. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            adj = System.Math.Abs(h1 * System.Math.Cos(Omega))
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.PR_MakeXY(xyMid, xyBotFlap5.X - adj, xyBotFlap5.y + opp)

            'Centre of Top Arc
            'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Delta = ARMDIA1.FN_CalcAngle(xyBotFlap5, xyTopFlap1)
            'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Delta = Delta - 90
            'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Delta = (Delta * PI) / 180
            'UPGRADE_WARNING: Couldn't resolve default property of object h1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object hyp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            hyp = 4 * h1
            'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object hyp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            opp = System.Math.Abs(hyp * System.Math.Sin(Delta))
            'UPGRADE_WARNING: Couldn't resolve default property of object Delta. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object hyp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            adj = System.Math.Abs(hyp * System.Math.Cos(Delta))
            'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.PR_MakeXY(xyCentre, xyMid.X + adj, xyMid.y + opp)
            'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TopRadius = ARMDIA1.FN_CalcLength(xyCentre, xyTopFlap1)
            'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TopArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TopArcRadius = TopRadius
            'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Phi = ARMDIA1.FN_CalcAngle(xyCentre, xyTopFlap1)
            'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Phi = Phi - 180
            'UPGRADE_WARNING: Couldn't resolve default property of object Alpha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Alpha = ARMDIA1.FN_CalcAngle(xyCentre, xyBotFlap5)
            'UPGRADE_WARNING: Couldn't resolve default property of object Alpha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Alpha = Alpha - 180
            'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Alpha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TopArcAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TopArcAngle = Alpha - Phi
            'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Alpha. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TopArcIncrement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TopArcIncrement = (Alpha - Phi) / 5

            'Calc Flap Length for position of marker
            'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object CircleCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            CircleCircum = PI * (2 * (TopArcRadius))
            'UPGRADE_WARNING: Couldn't resolve default property of object CircleCircum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TopArcAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TopArcLength = (TopArcAngle / 360) * CircleCircum
            'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FlapLength = TopArcLength + BottomArcLength
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            FlapMarkerLength = FlapLength / 2.7
            'Minimum length from market to seam is 2.5"
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If FlapLength - FlapMarkerLength < 2.5 Then FlapMarkerLength = FlapLength - 2.5
            'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If FlapMarkerLength < TopArcLength Then
                'UPGRADE_WARNING: Couldn't resolve default property of object TopArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FlapMarkerAngle = (FlapMarkerLength * 360) / (PI * (2 * TopArcRadius))
                'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FlapMarkerAngle = FlapMarkerAngle + Phi
                'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
                'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                opp = System.Math.Abs(TopRadius * System.Math.Sin(FlapMarkerAngle))
                'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                adj = System.Math.Abs(TopRadius * System.Math.Cos(FlapMarkerAngle))
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ARMDIA1.PR_MakeXY(xyFlapMarker, xyCentre.X - adj, xyCentre.y - opp)
                'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf FlapMarkerLength > TopArcLength Then
                'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object TempMarker. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                TempMarker = FlapMarkerLength - TopArcLength
                'UPGRADE_WARNING: Couldn't resolve default property of object TempMarker. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                TempMarker = BottomArcLength - TempMarker
                'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object TempMarker. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FlapMarkerAngle = (TempMarker * 360) / (PI * (2 * BottomRadius))
                'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
                'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                opp = System.Math.Abs(BottomArcRadius * System.Math.Sin(FlapMarkerAngle))
                'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object BottomArcRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                adj = System.Math.Abs(BottomArcRadius * System.Math.Cos(FlapMarkerAngle))
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object LittleBit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                LittleBit = BottomRadius - adj
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object LittleBit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ARMDIA1.PR_MakeXY(xyFlapMarker, Profile.X(Profile.n) + LastTapeValue - LittleBit, 0 + opp)
                'UPGRADE_WARNING: Couldn't resolve default property of object TopArcLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FlapMarkerLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf FlapMarkerLength = TopArcLength Then
                ARMDIA1.PR_MakeXY(xyFlapMarker, xyBotFlap5.X, xyBotFlap5.y)
            End If
            ARMDIA1.PR_SetLayer("Notes")
            'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'PrintLine(ARMDIA1.fNum, "hEnt = AddEntity(" & ARMDIA1.QQ & "marker" & ARMDIA1.QCQ & "closed arrow" & ARMDIA1.QC & "xyStart.x +" & Str(xyFlapMarker.X) & ARMDIA1.CC & "xyStart.y +" & Str(xyFlapMarker.y) & ",0.25,0.1,225);")
            PR_DrawClosedArrow(xyFlapMarker, 225)
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)

            For ii = 1 To 4
                'UPGRADE_WARNING: Couldn't resolve default property of object TopArcIncrement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Phi. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PhiExtra. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PhiExtra = Phi + (ii * TopArcIncrement)
                'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object PhiExtra. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                PhiExtra = (PhiExtra * PI) / 180
                'UPGRADE_WARNING: Couldn't resolve default property of object PhiExtra. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                opp = System.Math.Abs(TopRadius * System.Math.Sin(PhiExtra))
                'UPGRADE_WARNING: Couldn't resolve default property of object PhiExtra. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object TopRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                adj = System.Math.Abs(TopRadius * System.Math.Cos(PhiExtra))
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ShoulderFlap.X(10 - ii) = xyCentre.X - adj
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ShoulderFlap.y(10 - ii) = xyCentre.y - opp
            Next ii


            If VB.Left(txtFlap.Text, 6) = "Raglan" Then
                'Copy standard shoulder from profile down
                'only go one vetex past the FlapMarker
                RaglanFlap.n = 0
                Dim XX(10), YY(10) As Double
                XX(1) = 0
                YY(1) = 0
                RaglanFlap.X = XX
                RaglanFlap.y = YY
                For ii = 10 To 1 Step -1
                    RaglanFlap.n = RaglanFlap.n + 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    RaglanFlap.X(RaglanFlap.n) = ShoulderFlap.X(ii)
                    'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    RaglanFlap.y(RaglanFlap.n) = ShoulderFlap.y(ii)
                    'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If ShoulderFlap.X(ii) > xyFlapMarker.X Then Exit For
                Next ii
                ARMDIA1.PR_DrawFitted(RaglanFlap)
            Else
                ShoulderFlap.n = 10
                ARMDIA1.PR_DrawFitted(ShoulderFlap)
            End If

            'Draw Raglan
            If VB.Left(txtFlap.Text, 6) = "Raglan" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Rag1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Rag1 = System.Math.Sqrt((BottomRadius * BottomRadius) - (nNotchOffset * nNotchOffset))
                'UPGRADE_WARNING: Couldn't resolve default property of object Rag1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object BottomRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Rag1 = BottomRadius - Rag1
                'UPGRADE_WARNING: Couldn't resolve default property of object Rag1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ARMDIA1.PR_MakeXY(xyRaglan1, Profile.X(Profile.n) + LastTapeValue - Rag1, nNotchOffset)
                'UPGRADE_WARNING: Couldn't resolve default property of object nNotchToTangent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ARMDIA1.PR_MakeXY(xyRaglan2, xyRaglan1.X - nNotchToTangent, 0)
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng1 = ARMDIA1.FN_CalcAngle(xyRaglan2, xyRaglan1)
                'UPGRADE_WARNING: Couldn't resolve default property of object Rag2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Rag2 = ARMDIA1.FN_CalcLength(xyRaglan2, xyRaglan1)
                'UPGRADE_WARNING: Couldn't resolve default property of object Rag2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Rag2 = Rag2 / 2
                'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng1 = (RagAng1 * PI) / 180
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Rag2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                opp = System.Math.Abs(Rag2 * System.Math.Sin(RagAng1))
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Rag2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                adj = System.Math.Abs(Rag2 * System.Math.Cos(RagAng1))
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ARMDIA1.PR_MakeXY(xyRaglan3, xyRaglan2.X + adj, xyRaglan2.y + opp)

                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng1 = ARMDIA1.FN_CalcAngle(xyRaglan3, xyRaglan1)
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng1 = 90 - RagAng1
                'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng1 = (RagAng1 * PI) / 180
                'UPGRADE_WARNING: Couldn't resolve default property of object Rag3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Rag3 = ARMDIA1.FN_CalcLength(xyRaglan3, xyRaglan1)
                'UPGRADE_WARNING: Couldn't resolve default property of object Rag3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object nTemplateRadius. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Rag4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Rag4 = System.Math.Sqrt((nTemplateRadius * nTemplateRadius) - (Rag3 * Rag3))
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Rag4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                opp = System.Math.Abs(Rag4 * System.Math.Sin(RagAng1))
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Rag4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                adj = System.Math.Abs(Rag4 * System.Math.Cos(RagAng1))
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ARMDIA1.PR_MakeXY(xyRaglan4, xyRaglan3.X - adj, xyRaglan3.y + opp)
                ARMDIA1.PR_DrawArc(xyRaglan4, xyRaglan2, xyRaglan1)

                'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object nTemplateAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng2 = (nTemplateAngle * PI) / 180
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                opp = System.Math.Abs(RaglanBottom * System.Math.Sin(RagAng2))
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object RaglanBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                adj = System.Math.Abs(RaglanBottom * System.Math.Cos(RagAng2))
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ARMDIA1.PR_MakeXY(xyRaglan5, xyRaglan1.X + adj, xyRaglan1.y + opp)
                ARMDIA1.PR_DrawLine(xyRaglan1, xyRaglan5)
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng3 = ARMDIA1.FN_CalcAngle(xyRaglan5, xyFlapMarker)

                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If RagAng3 > 180 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    RagAng3 = RagAng3 - 180
                    'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object RagAng4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    RagAng4 = 180 - 115 - TopRagLanAngle - RagAng3
                    'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf RagAng3 <= 180 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    RagAng3 = RagAng3 - 90
                    'UPGRADE_WARNING: Couldn't resolve default property of object RagAng3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object RagAng4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    RagAng4 = 90 - RagAng3
                    'UPGRADE_WARNING: Couldn't resolve default property of object RagAng4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object TopRagLanAngle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    RagAng4 = (180 - TopRagLanAngle - 115) + RagAng4
                End If

                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng5 = RagAng4 - 13
                'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng4 = (RagAng4 * PI) / 180
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                opp = System.Math.Abs(RaglanTip * System.Math.Sin(RagAng4))
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng4. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object RaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                adj = System.Math.Abs(RaglanTip * System.Math.Cos(RagAng4))
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ARMDIA1.PR_MakeXY(xyRaglan6, xyRaglan5.X - adj, xyRaglan5.y + opp)

                'Draw Front line of tip
                'UPGRADE_WARNING: Couldn't resolve default property of object PI. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RagAng5 = (RagAng5 * PI) / 180
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                opp = System.Math.Abs(InnerRaglanTip * System.Math.Sin(RagAng5))
                'UPGRADE_WARNING: Couldn't resolve default property of object RagAng5. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object InnerRaglanTip. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                adj = System.Math.Abs(InnerRaglanTip * System.Math.Cos(RagAng5))
                'UPGRADE_WARNING: Couldn't resolve default property of object opp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object adj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ARMDIA1.PR_MakeXY(xyRaglan7, xyRaglan5.X - adj, xyRaglan5.y + opp)

                ARMDIA1.PR_DrawLine(xyRaglan5, xyRaglan6)
                ARMDIA1.PR_DrawLine(xyRaglan6, xyFlapMarker)
                ARMDIA1.PR_SetLayer("Notes")
                ARMDIA1.PR_DrawLine(xyRaglan5, xyRaglan7)

                '    PR_AddMarker xyRaglan1, "xyRaglan1"
                '    PR_AddMarker xyRaglan2, "xyRaglan2"
                '    PR_AddMarker xyRaglan3, "xyRaglan3"
                '    PR_AddMarker xyRaglan4, "xyRaglan4"
                '    PR_AddMarker xyRaglan5, "xyRaglan5"
                '    PR_AddMarker xyRaglan6, "xyRaglan6"
                '    PR_AddMarker xyRaglan7, "xyRaglan7"
                '    PR_AddMarker xyRaglan6, "xyRaglan5"
                '    PR_AddMarker xyFlapMarker, "xyFlapMarker"

            End If

            ARMDIA1.PR_SetLayer("Notes")
            ARMDIA1.PR_MakeXY(xyStrapText, xyFlapMarker.X - 1.5, xyFlapMarker.y)
            Strap = "STRAP = " & Trim(ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtStrap))) & Chr(34)
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                Strap = "STRAP = " & Trim(fnDisplayStrapToCM(cmtoinches(txtinchflag.Text, txtStrap)))
            End If
            If Val(txtFrontStrapLength.Text) > 0 Then
                ''-------------Strap = "BACK  STRAP = " & Trim(ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtStrap))) & Chr(34)
                ''------------Strap = Strap & Chr(10) & "FRONT STRAP = " & Trim(ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtFrontStrapLength))) & Chr(34)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    Strap = "BACK  STRAP = " & Trim(fnDisplayStrapToCM(cmtoinches(txtinchflag.Text, txtStrap))) & "CM"
                    Strap = Strap & Chr(10) & "FRONT STRAP = " & Trim(fnDisplayStrapToCM(cmtoinches(txtinchflag.Text, txtFrontStrapLength))) & "CM"
                Else
                    Strap = "BACK  STRAP = " & Trim(ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtStrap))) & Chr(34)
                    Strap = Strap & Chr(10) & "FRONT STRAP = " & Trim(ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtFrontStrapLength))) & Chr(34)
                End If

            End If
            If Val(txtWaistCir.Text) > 0 Then
                'Display Waist Circumference for D-Style flaps only
                ''-----------------Strap = Strap & Chr(10) & "WAIST = " & Trim(ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtWaistCir))) & Chr(34)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    Strap = Strap & Chr(10) & "WAIST = " & Trim(fnDisplayStrapToCM(cmtoinches(txtinchflag.Text, txtWaistCir)))
                Else
                    Strap = Strap & Chr(10) & "WAIST = " & Trim(ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtWaistCir))) & Chr(34)
                End If
            End If
            ''ARMDIA1.PR_DrawText(Strap, xyStrapText, 0.125, 0)
            ARMDIA1.PR_DrawMText(Strap, xyStrapText, True)

            'UPGRADE_WARNING: Couldn't resolve default property of object xfabby. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xfabby = txtFlap.Text
            'UPGRADE_WARNING: Couldn't resolve default property of object xfablen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xfablen = Len(xfabby)
            'UPGRADE_WARNING: Couldn't resolve default property of object xfablen. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object xfabdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xfabdist = (xfablen * 0.075) + 0.1
            'UPGRADE_WARNING: Couldn't resolve default property of object xfabdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.PR_MakeXY(xyFlapText1, xyFlapMarker.X - xfabdist, xyFlapMarker.y - 0.75)
            ''ARMDIA1.PR_DrawText(txtFlap.Text, xyFlapText1, 0.125, 0)
            ARMDIA1.PR_DrawMText(txtFlap.Text, xyFlapText1, True)

            ARMDIA1.PR_MakeXY(xyPt1, Profile.X(Profile.n), 0)
            'UPGRADE_WARNING: Couldn't resolve default property of object LastTapeValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ARMDIA1.PR_MakeXY(xyPt2, Profile.X(Profile.n) + LastTapeValue, 0)

            If VB.Left(txtFlap.Text, 6) = "Raglan" Then
                ARMDIA1.PR_MakeXY(xyPotA, xyPt2.X - 0.2875, xyPt2.y)
                ARMDIA1.PR_MakeXY(xyPotB, xyPt2.X - 1.225, xyPt2.y)
                PR_DrawLineOffset(xyPt1, xyPotA, 0.6875) 'Inner tram line
                PR_DrawLineOffset(xyPt1, xyPotB, 0.1875) 'seam line at 3/16ths
            Else
                ARMDIA1.PR_MakeXY(xyPotA, xyPt2.X - 0.09375, xyPt2.y)
                ARMDIA1.PR_MakeXY(xyPotB, xyPt2.X - 0.03125, xyPt2.y)
                PR_DrawLineOffset(xyPt1, xyPotA, 0.6875) 'Inner tram line
                PR_DrawLineOffset(xyPt1, xyPotB, 0.1875) 'seam line at 3/16ths
            End If

            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)

            If VB.Left(txtFlap.Text, 6) = "Raglan" Then
                ARMDIA1.PR_MakeXY(xyPotA, xyPt2.X - 1.75, xyPt2.y)
                ARMDIA1.PR_DrawLine(xyPt1, xyRaglan2)
            Else
                ARMDIA1.PR_DrawLine(xyPt1, xyPt2) 'Fold Line
            End If
            ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "FoldLineFlap") 'sFileNo, sSide from FN_Open

        End If
    End Sub

    Private Sub PR_GauntletHoleCheck()

        Dim Gauntletpwdiff, count As Object
        If txtGauntletFlag.Text = "True" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Gauntletpwdiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Gauntletpwdiff = (Val(txtWristNo.Text) - Val(txtPalmNo.Text))
            'UPGRADE_WARNING: Couldn't resolve default property of object Gauntletpwdiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Gauntletpwdiff = Gauntletpwdiff - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object Gauntletpwdiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Gauntletpwdiff = 1 Then
                g_sHoleCheck = "True"
                Length(txtPalmNo.Text).Text = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(txtPalmNo + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Measurements(CDbl(txtPalmNo.Text) + 1) = ""
                gms(txtPalmNo.Text).Text = ""
                reds(txtPalmNo.Text).Text = ""
                mms(txtPalmNo.Text).Text = ""
                InchText(txtPalmNo.Text).Text = ""
            Else
                g_sHoleCheck = "False"
            End If
        End If

    End Sub

    Private Sub PR_GetFabric()
        '   ------------------------
        '   *** Get Fabric array ***
        '   ------------------------
        Dim Fabtype, fabnumber, fabnam As Object
        Dim Modul, fabcount As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        fabnam = VB.Left(cboFabric.Text, 3)
        'UPGRADE_WARNING: Couldn't resolve default property of object Fabtype. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Fabtype = Mid(cboFabric.Text, 5, 3)
        'UPGRADE_WARNING: Couldn't resolve default property of object Fabtype. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(Fabtype) > 340 Or Val(Fabtype) < 160 Then
            MsgBox("Fabric Modulus Not Recognised", 48, "Arm Details")
            Exit Sub
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (fabnam <> "Pow") And (fabnam <> "Bob") Then
            MsgBox("Fabric Type Not Recognised, Should be either Bobbinette or Powernet", 48, "Arm Details")
            Exit Sub
        End If
        g_sFabnam = cboFabric.Text
        txtFabric.Text = cboFabric.Text
        'UPGRADE_WARNING: Couldn't resolve default property of object fabnumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        fabnumber = 160
        'UPGRADE_WARNING: Couldn't resolve default property of object fabcount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        fabcount = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object fabcount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While fabcount < 20
            'UPGRADE_WARNING: Couldn't resolve default property of object fabnumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Fabtype. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Val(Fabtype) = fabnumber Then
                'UPGRADE_WARNING: Couldn't resolve default property of object fabcount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Modul = fabcount
                'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                g_sModulus = Modul
                'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtModulus.Text = Modul
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object fabcount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            fabcount = fabcount + 1
            'UPGRADE_WARNING: Couldn't resolve default property of object fabnumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            fabnumber = fabnumber + 10
        Loop
        '   *** End Get Fabric ***
    End Sub

    Private Sub PR_GetGauntletDetails()
        '       *** Start Gauntlet Details ***

        If chkGauntlets.Checked = True Then
            txtWristNo.Text = CStr(cboWrist.SelectedIndex + 1)
            If Val(txtWristNo.Text) = 0 Then txtWristNo.Text = ""

            txtPalmNo.Text = CStr(cboPalm.SelectedIndex + 1)
            If Val(txtPalmNo.Text) = 0 Then txtPalmNo.Text = ""

            Call Gauntlet(chkNoThumb.Checked, txtPalmNo.Text, txtWristNo.Text, txtPalmWrist.Text, txtThumb.Text,
                          txtSex.Text, txtAge.Text, chkEnclosed.Checked, txtCircum.Text, chkDetachable.Checked, txtinchflag.Text)
        ElseIf chkGauntlets.Checked = False Then
            txtGauntletFlag.Text = "False"
        End If

        ''       *** End Gauntlet Details ***

    End Sub

    Private Sub PR_GetPressure()
        '   --------------------
        '   *** get Pressure ***
        '   --------------------
        If cboPressure.Text = "15mm" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sAmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sAmm = 12
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sBmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sBmm = 15
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sCmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sCmm = 10
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sDmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sDmm = 8
        ElseIf cboPressure.Text = "20mm" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sAmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sAmm = 16
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sBmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sBmm = 20
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sCmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sCmm = 13
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sDmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sDmm = 10
        ElseIf cboPressure.Text = "25mm" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sAmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sAmm = 20
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sBmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sBmm = 25
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sCmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sCmm = 17
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sDmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sDmm = 13
        End If
        txtMM.Text = cboPressure.Text
        '   --- end get pressure ---

    End Sub

    Private Sub PR_GetStartEndTapes()
        Dim g_iSecondLastTape As Object
        Dim g_iSecondTape As Object
        'Now sets start and last tape text values w.r.t style start and end
        Dim i, ii As Object
        'first  & second tapes
        'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        i = 0
        txtFirstTape.Text = ""
        'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While i < 18
            If Length(i).Text <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ARMDIA1.g_iFirstTape = i + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object g_iSecondTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                g_iSecondTape = i + 2
                Exit Do
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            i = i + 1
        Loop

        If ARMDIA1.g_iStyleFirstTape > ARMDIA1.g_iFirstTape And ARMDIA1.g_iStyleFirstTape >= 0 And cboDistalTape.SelectedIndex <> 0 Then
            'Use style tape value if it is given ( g_iStyleFirstTape = -1 => not given)
            txtFirstTape.Text = CStr(ARMDIA1.g_iStyleFirstTape)
            txtSecondTape.Text = CStr(ARMDIA1.g_iStyleFirstTape + 1)
        Else
            txtFirstTape.Text = CStr(ARMDIA1.g_iFirstTape)
            'UPGRADE_WARNING: Couldn't resolve default property of object g_iSecondTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtSecondTape.Text = g_iSecondTape
        End If

        'secondlast & last tapes
        txtLastTape.Text = ""
        'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ii = 17
        'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While ii > 0
            If Length(ii).Text <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ARMDIA1.g_iLastTape = ii + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object g_iSecondLastTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                g_iSecondLastTape = ii
                Exit Do
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ii = ii - 1
        Loop

        If ARMDIA1.g_iStyleLastTape < ARMDIA1.g_iLastTape And ARMDIA1.g_iStyleLastTape >= 0 And cboProximalTape.SelectedIndex <> 0 Then
            'Use style tape value if it is given ( g_iStyleLastTape = -1 => not given)
            txtLastTape.Text = CStr(ARMDIA1.g_iStyleLastTape)
            txtSecondLastTape.Text = CStr(ARMDIA1.g_iStyleLastTape - 1)
        Else
            txtLastTape.Text = CStr(ARMDIA1.g_iLastTape)
            'UPGRADE_WARNING: Couldn't resolve default property of object g_iSecondLastTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtSecondLastTape.Text = g_iSecondLastTape
        End If

    End Sub

    Private Sub PR_InitaliseArrays(ByRef w As Object, ByRef X As Object)
        If Length(w).Text = "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Measurements(X) = ""
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            tapewidths(X) = ""
        End If
    End Sub

    Private Sub PR_InitaliseEmpty(ByRef w As Object)
        'initalise empty fields on front end
        If Length(w).Text = "" Then
            mms(w).Text = ""
            gms(w).Text = ""
            reds(w).Text = ""
            InchText(w).Text = ""
        End If
    End Sub

    Private Sub PR_LastCircum(ByRef w As Object, ByRef X As Object)
        Dim temptapes1, temptapes2 As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If w > 9 Then

            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Length(w).Text <> "" And Length(w - 1).Text <> "" Then
                'check circum of last tape
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object temptapes1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                temptapes1 = cmtoinches(txtinchflag.Text, Length(w))

                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object temptapes2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                temptapes2 = cmtoinches(txtinchflag.Text, Length(w - 1))

                'UPGRADE_WARNING: Couldn't resolve default property of object temptapes2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object temptapes1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If (w = (Val(txtLastTape.Text) - 1)) And (Val(txtAge.Text) > 10) And (Val(temptapes1 - temptapes2) > 1.5) And (Val(txtLastTape.Text) > 11) Then

                    MsgBox("Last tape is measured over axilla. Exclude measurement", 48, "Arm Details")

                    ' Exit Sub
                    'UPGRADE_WARNING: Couldn't resolve default property of object temptapes2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object temptapes1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf (w = (Val(txtLastTape.Text) - 1)) And (Val(txtAge.Text) < 11) And (Val(temptapes1 - temptapes2) > 0.75) And (Val(txtLastTape.Text) > 11) Then

                    MsgBox("Last tape is measured over axilla. Exclude measurement", 48, "Arm Details, child")

                    ' Exit Sub

                End If

            End If
        End If

    End Sub

    Private Sub PR_LastTapeRed(ByRef w As Object, ByRef X As Object, ByRef Modulus As Object)
        Dim Backgram As Object
        If Length(w).Text <> "" Then
            'check last tape reduction
            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (w > 1) And (w = Val(txtLastTape.Text) - 1) And (cboFlaps.Text = "") Then
                If VB.Left(g_sFabnam, 3) = "Pow" Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modulus, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Backgram = Pow(Modulus, 1)
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If Val(gms(w).Text) <> Backgram Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        gms(w).Text = CStr(ARMDIA1.round(Backgram))
                        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                    End If
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modulus, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Backgram = Bob(Modulus, 1)
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If Val(gms(w).Text) <> Backgram Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        gms(w).Text = CStr(ARMDIA1.round(Backgram))
                        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                    End If
                End If
                reds(w).Text = CStr(10)

                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (w > 1) And (w = Val(txtLastTape.Text) - 1) And ((txtFlap.Text <> "") Or (VB.Left(cboFlaps.Text, 4) = "None") Or (VB.Left(cboFlaps.Text, 5) = "Style")) Then
                If VB.Left(cboFlaps.Text, 4) <> "Vest" And VB.Left(cboFlaps.Text, 4) <> "Body" Then
                    If VB.Left(g_sFabnam, 3) = "Pow" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modulus, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Backgram = Pow(Modulus, 3) '12 reduction for Flaps
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If Val(gms(w).Text) <> Backgram Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            gms(w).Text = CStr(ARMDIA1.round(Backgram))
                            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                        End If
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modulus, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Backgram = Bob(Modulus, 3) '12 reduction for Flaps
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If Val(gms(w).Text) <> Backgram Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            gms(w).Text = CStr(ARMDIA1.round(Backgram))
                            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                        End If
                    End If
                    reds(w).Text = CStr(12)
                Else
                    If VB.Left(g_sFabnam, 3) = "Pow" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modulus, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Backgram = Pow(Modulus, 1) '10 reduction for Vest & Body
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If Val(gms(w).Text) <> Backgram Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            gms(w).Text = CStr(ARMDIA1.round(Backgram))
                            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                        End If
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modulus, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Backgram = Bob(Modulus, 1) '10 reduction for Vest & Body
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If Val(gms(w).Text) <> Backgram Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            gms(w).Text = CStr(ARMDIA1.round(Backgram))
                            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                        End If
                    End If
                    reds(w).Text = CStr(10)
                End If
            End If
        End If
        ' end check last tape reduction

    End Sub

    Private Sub PR_OldDrawing()
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "Table(" & ARMDIA1.QQ & "add" & ARMDIA1.QCQ & "field" & ARMDIA1.QCQ & "Flap" & ARMDIA1.QCQ & "string" & ARMDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "Table(" & ARMDIA1.QQ & "add" & ARMDIA1.QCQ & "field" & ARMDIA1.QCQ & "Gauntlet" & ARMDIA1.QCQ & "string" & ARMDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "Table(" & ARMDIA1.QQ & "add" & ARMDIA1.QCQ & "field" & ARMDIA1.QCQ & "Stump" & ARMDIA1.QCQ & "string" & ARMDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "Table(" & ARMDIA1.QQ & "add" & ARMDIA1.QCQ & "field" & ARMDIA1.QCQ & "TapeLengths" & ARMDIA1.QCQ & "string" & ARMDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "Table(" & ARMDIA1.QQ & "add" & ARMDIA1.QCQ & "field" & ARMDIA1.QCQ & "TapeMMs" & ARMDIA1.QCQ & "string" & ARMDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "Table(" & ARMDIA1.QQ & "add" & ARMDIA1.QCQ & "field" & ARMDIA1.QCQ & "Reduction" & ARMDIA1.QCQ & "string" & ARMDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "Table(" & ARMDIA1.QQ & "add" & ARMDIA1.QCQ & "field" & ARMDIA1.QCQ & "Grams" & ARMDIA1.QCQ & "string" & ARMDIA1.QQ & ");")
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "Table(" & ARMDIA1.QQ & "add" & ARMDIA1.QCQ & "field" & ARMDIA1.QCQ & "ID" & ARMDIA1.QCQ & "string" & ARMDIA1.QQ & ");")
    End Sub

    Private Sub PR_PalmRed(ByRef w As Object, ByRef X As Object, ByRef Modulus As Object)
        Dim Backgram As Object
        If Length(w).Text <> "" Then
            'check palm reduction
            If (chkGauntlets.Checked = True And X = Val(txtPalmNo.Text)) Or (chkGauntlets.Checked = True And w = (CDbl(txtFirstTape.Text.Replace("st", "").Trim) - 1)) Then

                reds(w).Text = CStr(15)

                If VB.Left(g_sFabnam, 3) = "Pow" Then
                    Backgram = Pow(Modulus, 6)
                Else
                    Backgram = Bob(Modulus, 6)
                End If

                mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))

                gms(w).Text = CStr(ARMDIA1.round(Backgram))

            End If
        End If

    End Sub

    Private Sub PR_Pressurecalc(ByRef Index As Short)
        Dim w, X As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        w = Index
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X = Index + 1
        PR_InitaliseEmpty(w) 'Zero any empty fields
        PR_InitaliseArrays(w, X)
        If (Length(Index).Text <> "" And mms(Index).Text = "") Or (Length(Index).Text <> "" And g_sPressureChange = "True") Then
            PR_SetMMs(w, g_sAmm, g_sBmm, g_sCmm, g_sDmm)
        End If

        PR_Reductions(w, X)
        PR_WristRed(w, X, g_sModulus)
        PR_PalmRed(w, X, g_sModulus)
        PR_Stump(w, X)
        PR_LastCircum(w, X)
        PR_LastTapeRed(w, X, g_sModulus)
        PR_CheckFlaps(w)
        PR_Detachable(w, X, g_sModulus, g_sFabnam)
    End Sub

    Private Sub PR_Reductions(ByRef w As Object, ByRef X As Object)
        Dim Gram As Object
        If Length(w).Text <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Gram = Fix(Val(mms(w).Text) * Val(Measurements(X)))
            'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Gram <> 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                gms(w).Text = CStr(ARMDIA1.round(Gram))
                'UPGRADE_WARNING: Couldn't resolve default property of object GetReduction(g_sModulus, Gram, g_sFabnam). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GetReduction(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                reds(w).Text = GetReduction(g_sModulus, Gram, g_sFabnam)
            End If
        End If

    End Sub

    Private Sub PR_ReSetMMs(ByRef w As Object, ByRef Modulus As Object)
        Dim tempmms As Object
        Dim newmms As Object
        Dim Backgram As Object
        Dim givenreduction, X As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X = w + 1
        If Length(w).Text <> "" And mms(w).Text <> "" And reds(w).Text <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object givenreduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            givenreduction = Val(CStr(CDbl(reds(w).Text) - 9))
            'UPGRADE_WARNING: Couldn't resolve default property of object givenreduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If givenreduction > 32 Or givenreduction < 1 Then
                MsgBox("Reduction Value Not Recognised", 48, "Arm Details")
                Exit Sub
            End If
            If VB.Left(g_sFabnam, 3) = "Pow" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object givenreduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modulus, givenreduction). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Backgram = Pow(Modulus, givenreduction)
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Val(gms(w).Text) <> Backgram Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    gms(w).Text = CStr(ARMDIA1.round(Backgram))
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    newmms = Backgram / Measurements(X)
                    'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tempmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tempmms = ARMDIA1.round(newmms)
                    'UPGRADE_WARNING: Couldn't resolve default property of object tempmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tempmms = (newmms - tempmms) * 10
                    'UPGRADE_WARNING: Couldn't resolve default property of object tempmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If tempmms >= 5 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mms(w).Text = CStr(ARMDIA1.round(newmms) + 1)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mms(w).Text = CStr(ARMDIA1.round(newmms))
                    End If

                End If
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object givenreduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modulus, givenreduction). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Backgram = Bob(Modulus, givenreduction)
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Val(gms(w).Text) <> Backgram Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    gms(w).Text = CStr(ARMDIA1.round(Backgram))
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    newmms = Backgram / Measurements(X)
                    'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tempmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tempmms = ARMDIA1.round(newmms)
                    'UPGRADE_WARNING: Couldn't resolve default property of object tempmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tempmms = (newmms - tempmms) * 10
                    'UPGRADE_WARNING: Couldn't resolve default property of object tempmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If tempmms >= 5 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mms(w).Text = CStr(ARMDIA1.round(newmms) + 1)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mms(w).Text = CStr(ARMDIA1.round(newmms))
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub PR_SetDistances(ByRef w As Object, ByRef X As Object)

        'Length(w)
        'Measurements(X)

        'Check Distances between tapes
        Dim tdist, tpace As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object tpace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        tpace = 1.375
        If Length(w).Text <> "" Then
            'First tape
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If X = Val(txtFirstTape.Text) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                tapewidths(X) = "0"

                'secondTape
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf X = Val(txtSecondTape.Text) And chkGauntlets.Checked = False Then
                If Val(txtWristPleat1.Text) = 0 Or Not txtWristPleat1.Enabled Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tpace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = tpace
                ElseIf Val(txtWristPleat1.Text) > 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object tdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tdist = Val(txtWristPleat1.Text)
                    'UPGRADE_WARNING: Couldn't resolve default property of object FN_PleatDistances(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tdist = FN_PleatDistances(txtinchflag, tdist)
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = tapewidths(X - 1) + tdist
                End If

                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf X = Val(txtSecondTape.Text) And chkGauntlets.Checked = True Then
                If Val(Length(w).Text) = 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = ""
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf Val(Length(w).Text) > 0 And Val(txtWristNo.Text) = X Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = Val(txtPalmWristDist.Text)
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf Val(Length(w).Text) > 0 And Val(txtWristNo.Text) <> X Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tpace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = tpace
                End If

                'thirdtape
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf X = Val(CStr(CDbl(txtSecondTape.Text) + 1)) And chkGauntlets.Checked = False Then
                If Val(txtWristPleat2.Text) = 0 Or Not txtWristPleat2.Enabled Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tpace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = tapewidths(X - 1) + tpace
                ElseIf Val(txtWristPleat2.Text) > 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object tdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tdist = Val(txtWristPleat2.Text)
                    'UPGRADE_WARNING: Couldn't resolve default property of object FN_PleatDistances(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tdist = FN_PleatDistances(txtinchflag, tdist)
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = Val(tapewidths(X - 1)) + tdist
                End If

                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf X = Val(CStr(CDbl(txtSecondTape.Text) + 1)) And chkGauntlets.Checked = True Then
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If (Length(w - 1).Text = "") Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = Val(tapewidths(X - 2)) + Val(txtPalmWristDist.Text)
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf Length(w - 1).Text <> "" And Val(txtWristNo.Text) = X Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = tapewidths(X - 1) + Val(txtPalmWristDist.Text)
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf Length(w - 1).Text <> "" And Val(txtWristNo.Text) <> X Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tpace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = tapewidths(X - 1) + tpace
                    'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf Length(w - 1).Text <> "" Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = ""
                End If

                'fourthtape
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf X = Val(CStr(CDbl(txtSecondTape.Text) + 2)) And chkGauntlets.Checked = False Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tpace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                tapewidths(X) = tapewidths(X - 1) + tpace

                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf X = Val(CStr(CDbl(txtSecondTape.Text) + 2)) And chkGauntlets.Checked = True Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Length(w - 1).Text = "" And Val(txtWristNo.Text) = X Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = tapewidths(X - 2) + Val(txtPalmWristDist.Text)
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf Length(w - 1).Text <> "" And Val(txtWristNo.Text) = X Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = tapewidths(X - 1) + Val(txtPalmWristDist.Text)
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf Length(w - 1).Text <> "" And Val(txtWristNo.Text) <> X Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tpace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = tapewidths(X - 1) + tpace
                End If

                'Tapes above fourthtape
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf X > (Val(txtSecondTape.Text) + 2) And X < Val(txtSecondLastTape.Text) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tpace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                tapewidths(X) = tapewidths(X - 1) + tpace

                'Lasttape -1
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf X = Val(txtSecondLastTape.Text) And (X > 1) Then
                If Val(txtShoulderPleat2.Text) = 0 Or Not txtShoulderPleat2.Enabled Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tpace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = Val(tapewidths(X - 1)) + tpace
                ElseIf Val(txtShoulderPleat2.Text) > 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object tdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tdist = Val(txtShoulderPleat2.Text)
                    'UPGRADE_WARNING: Couldn't resolve default property of object FN_PleatDistances(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tdist = FN_PleatDistances(txtinchflag, tdist)
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = Val(tapewidths(X - 1)) + tdist
                End If

                'Last tape
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf X = Val(txtLastTape.Text) And (X > 1) Then
                If Val(txtShoulderPleat1.Text) = 0 Or Not txtShoulderPleat1.Enabled Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tpace. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = Val(tapewidths(X - 1)) + tpace
                ElseIf Val(txtShoulderPleat1.Text) > 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object tdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tdist = Val(txtShoulderPleat1.Text)
                    'UPGRADE_WARNING: Couldn't resolve default property of object FN_PleatDistances(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tdist = FN_PleatDistances(txtinchflag, tdist)
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tdist. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tapewidths(X) = Val(tapewidths(X - 1)) + tdist
                End If
            End If

            'The following statements duplicate some of the above.  They are here
            'as bug fixes as I can't understand the above code.  The above does work
            '(mostly) so I have left well enough alone.

            'Account for wrist pleats if a gauntlet given
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Val(txtWristNo.Text) + 1 = X And (Val(txtWristPleat1.Text) > 0 And txtWristPleat1.Enabled) And chkGauntlets.Checked = True Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_PleatDistances(txtinchflag, Val(txtWristPleat1)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                tapewidths(X) = Val(tapewidths(X - 1)) + FN_PleatDistances(txtinchflag, Val(txtWristPleat1.Text))
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Val(txtWristNo.Text) + 2 = X And (Val(txtWristPleat2.Text) > 0 And txtWristPleat2.Enabled) And chkGauntlets.Checked = True Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_PleatDistances(txtinchflag, Val(txtWristPleat2)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                tapewidths(X) = Val(tapewidths(X - 1)) + FN_PleatDistances(txtinchflag, Val(txtWristPleat2.Text))
            End If

            'Account for 1ST shoulder pleat
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Val(txtLastTape.Text) = X And (Val(txtShoulderPleat1.Text) > 0 And txtShoulderPleat1.Enabled) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_PleatDistances(txtinchflag, Val(txtShoulderPleat1)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                tapewidths(X) = Val(tapewidths(X - 1)) + FN_PleatDistances(txtinchflag, Val(txtShoulderPleat1.Text))
            End If
            'Account for 2nd shoulder pleat
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Val(txtSecondLastTape.Text) = X And (Val(txtShoulderPleat2.Text) > 0 And txtShoulderPleat2.Enabled) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object FN_PleatDistances(txtinchflag, Val(txtShoulderPleat2)). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object tapewidths(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                tapewidths(X) = Val(tapewidths(X - 1)) + FN_PleatDistances(txtinchflag, Val(txtShoulderPleat2.Text))
            End If


        End If
        'End Check Distances between tape
    End Sub

    Private Sub PR_SetGlobals()
        '    SetTapeLengths
        g_sTapeLengths = txtTapeLent.Text

        '    SetMMs
        g_sTapeMMs = txtTapeMM.Text
        '    SetGrams
        g_sGrams = txtWeight.Text
        '    SetReductions
        g_sReduction = txtReduction.Text
        g_sFlapStrap = txtStrap.Text
        'Revised method w.r.t. multiple styles
        If optProximalTape(0).Checked = True Then
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sFlapChk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sFlapChk = "False"
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sFlapChk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sFlapChk = "True"
        End If
        If optProximalTape(1).Checked = True Then
            'UPGRADE_WARNING: Couldn't resolve default property of object g_sFlapChk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            g_sFlapChk = "True"
        End If
        ARMDIA1.g_sPatient = txtPatientName.Text
        g_sGaunt = txtGauntletFlag.Text
        g_sNoThumb = txtNoThumb.Text
        ARMDIA1.g_sPatient = txtPatientName.Text
        g_sWpleat1 = txtWristPleat1.Text
        If g_sWpleat1 = "" Then g_sWpleat1 = "0"
        g_sWpleat2 = txtWristPleat2.Text
        g_sSpleat1 = txtShoulderPleat1.Text
        If g_sSpleat1 = "" Then g_sSpleat1 = "0"
        g_sSpleat2 = txtShoulderPleat2.Text
        g_sDetGaunt = txtDetach.Text
        g_sEnclosedThumb = txtEnclosed.Text
        g_sPalmNo = txtPalmNo.Text
        g_sWristNo = txtWristNo.Text
        g_sPalmWristDist = txtPalmWrist.Text
        g_sThumbCircum = txtCircum.Text
        g_sThumbLength = txtThumb.Text
        g_sSecondTape = txtSecondTape.Text
        g_sSecondLastTape = txtSecondLastTape.Text
        g_sMM = cboPressure.Text
        g_sFirstTape = txtFirstTape.Text
        g_sLastTape = txtLastTape.Text
        g_sStump = txtStump.Text
        ARMDIA1.g_sSide = txtSleeve.Text
        'UPGRADE_WARNING: Couldn't resolve default property of object g_sWaistCir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        g_sWaistCir = txtWaistCir.Text

    End Sub

    Public Sub PR_SetMMs(ByRef w As Object, ByRef Amm As Object, ByRef Bmm As Object, ByRef Cmm As Object, ByRef Dmm As Object)
        Dim GenMMs As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Length(w).Text <> "" And w >= CDbl(txtFirstTape.Text) - 1 And w <= CDbl(txtLastTape.Text) - 1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (w = CDbl(txtFirstTape.Text) - 1) Or (w = 9) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Amm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GenMMs = Amm
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (w > 11) And (w < (CDbl(txtSecondLastTape.Text) - 1)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Amm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GenMMs = Amm
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf ((w > CDbl(txtFirstTape.Text) - 1) And (w < 9)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Bmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GenMMs = Bmm
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (w = 11) Or (w = (CDbl(txtSecondLastTape.Text) - 1)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Cmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GenMMs = Cmm
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (w = 10) Or (w = (CDbl(txtLastTape.Text) - 1)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Dmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GenMMs = Dmm
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (txtWristNo.Text <> "" And w = Val(txtWristNo.Text) - 1) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Amm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GenMMs = Amm
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mms(w).Text = CStr(Val(GenMMs))
        Else
            mms(w).Text = ""
        End If
    End Sub

    Private Sub PR_SetRresults()
        SetTapeLengths()
        SetMMs()
        SetGrams()
        SetReductions()
    End Sub

    Private Sub PR_Stump(ByRef w As Object, ByRef X As Object)
        Dim Backgram As Object
        If Length(w).Text <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (w = (CDbl(txtFirstTape.Text) - 1)) And chkStump.Checked = True Then
                reds(w).Text = CStr(14)
                If VB.Left(g_sFabnam, 3) = "Pow" Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object Pow(txtModulus, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Backgram = Pow(CInt(txtModulus.Text), 5)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object Bob(txtModulus, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Backgram = Bob(CInt(txtModulus.Text), 5)
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                gms(w).Text = CStr(ARMDIA1.round(Backgram))
                txtStump.Text = "True"
            End If
        End If
    End Sub

    Private Sub PR_WristRed(ByRef w As Object, ByRef X As Object, ByRef Modulus As Object)
        Dim Backgram, Reduction As Object

        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Length(w).Text <> "" And X = Val(txtFirstTape.Text) And txtStump.Text <> "True" And txtGauntletFlag.Text <> "True" Then

            ' check wrist reduction
            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (w < 9) And reds(w).Text = "" Then
                reds(w).Text = CStr(14)
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf CDbl(reds(w).Text) > 14 And (w < 9) Then
                reds(w).Text = CStr(14)
                'Set elbow reduction to 5
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf CDbl(reds(w).Text) <> 5 And (w = 10) Then
                reds(w).Text = CStr(5)
                'Above elbow is also a 5
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf CDbl(reds(w).Text) <> 5 And (w > 10) Then
                reds(w).Text = CStr(5)
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (w < 9) And ((CDbl(reds(w).Text) < 10) And (CDbl(reds(w).Text) > 0)) Then
                reds(w).Text = CStr(10)
            End If

            'Set up an index into the reduction to grams mapping table
            'minimum is 1
            'UPGRADE_WARNING: Couldn't resolve default property of object Reduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Reduction = CDbl(reds(w).Text) - 9
            'UPGRADE_WARNING: Couldn't resolve default property of object Reduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Reduction < 1 Then Reduction = 1

            'From either powernet or bobinnet table get the grams and back calculate the
            'mms
            If VB.Left(g_sFabnam, 3) = "Pow" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Reduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modulus, Reduction). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Backgram = Pow(Modulus, Reduction)
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object Reduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modulus, Reduction). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Backgram = Bob(Modulus, Reduction)
            End If

            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mms(w).Text = CStr(ARMDIA1.round(Backgram / Measurements(X)))
            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            gms(w).Text = CStr(ARMDIA1.round(Backgram))

        End If

        ' end check wrist reduction
    End Sub

    Private Sub select_text_in_Box1(ByRef Text_Box_Name As System.Windows.Forms.Control)
        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelStart. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''Text_Box_Name.SelStart = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object Text_Box_Name.SelLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ''Text_Box_Name. = 4
    End Sub

    Private Sub SetGrams()
        'Dim a, b, c, d, e, f, G, h, i, j, k, l, m, n, o, P, q, r
        Dim ii As Short
        txtWeight.Text = ""
        For ii = 0 To 17
            'UPGRADE_WARNING: Couldn't resolve default property of object PadGram(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtWeight.Text = txtWeight.Text & PadGram(gms(ii))
        Next ii
        'a = PadGram(gms(0))
        'b = PadGram(gms(1))
        'c = PadGram(gms(2))
        'd = PadGram(gms(3))
        'e = PadGram(gms(4))
        'f = PadGram(gms(5))
        'G = PadGram(gms(6))
        'h = PadGram(gms(7))
        'i = PadGram(gms(8))
        'j = PadGram(gms(9))
        'k = PadGram(gms(10))
        'l = PadGram(gms(11))
        'm = PadGram(gms(12))
        'n = PadGram(gms(13))
        'o = PadGram(gms(14))
        'P = PadGram(gms(15))
        'q = PadGram(gms(16))
        'r = PadGram(gms(17))
        'txtWeight = a & b & c & d & e & f & G & h & i & j & k & l & m & n & o & P & q & r

    End Sub

    Private Sub SetMMs()
        'Dim a, b, c, d, e, f, G, h, i, j, k, l, m, n, o, P, q, r
        Dim ii As Short
        txtTapeMM.Text = ""
        For ii = 0 To 17
            'UPGRADE_WARNING: Couldn't resolve default property of object PadMM(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtTapeMM.Text = txtTapeMM.Text & PadMM(mms(ii))
        Next ii
        'a = PadMM(mms(0))
        'b = PadMM(mms(1))
        'c = PadMM(mms(2))
        'd = PadMM(mms(3))
        'e = PadMM(mms(4))
        'f = PadMM(mms(5))
        'G = PadMM(mms(6))
        'h = PadMM(mms(7))
        'i = PadMM(mms(8))
        'j = PadMM(mms(9))
        'k = PadMM(mms(10))
        'l = PadMM(mms(11))
        'm = PadMM(mms(12))
        'n = PadMM(mms(13))
        'o = PadMM(mms(14))
        'P = PadMM(mms(15))
        'q = PadMM(mms(16))
        'r = PadMM(mms(17))
        'txtTapeMM = a & b & c & d & e & f & G & h & i & j & k & l & m & n & o & P & q & r

    End Sub

    Private Sub SetReductions()
        'Dim a, b, c, d, e, f, G, h, i, j, k, l, m, n, o, P, q, r
        Dim ii As Short
        txtReduction.Text = ""
        For ii = 0 To 17
            'UPGRADE_WARNING: Couldn't resolve default property of object PadRed(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtReduction.Text = txtReduction.Text & PadRed(reds(ii))
        Next ii
        'a = PadRed(reds(0))
        'b = PadRed(reds(1))
        'c = PadRed(reds(2))
        'd = PadRed(reds(3))
        'e = PadRed(reds(4))
        'f = PadRed(reds(5))
        'G = PadRed(reds(6))
        'h = PadRed(reds(7))
        'i = PadRed(reds(8))
        'j = PadRed(reds(9))
        'k = PadRed(reds(10))
        'l = PadRed(reds(11))
        'm = PadRed(reds(12))
        'n = PadRed(reds(13))
        'o = PadRed(reds(14))
        'P = PadRed(reds(15))
        'q = PadRed(reds(16))
        'r = PadRed(reds(17))
        'txtReduction = a & b & c & d & e & f & G & h & i & j & k & l & m & n & o & P & q & r

    End Sub

    Private Function SetStump(ByRef Length As Object) As Object
        Dim Circum As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object Length. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Circum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Circum = Length
        'UPGRADE_WARNING: Couldn't resolve default property of object Circum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Circum = Circum * 0.86
        'UPGRADE_WARNING: Couldn't resolve default property of object Circum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object SetStump. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        SetStump = Circum
    End Function

    Private Sub SetTapeLengths()
        Dim ii As Short
        txtTapeLent.Text = ""
        For ii = 0 To 17
            'UPGRADE_WARNING: Couldn't resolve default property of object PadTape(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            txtTapeLent.Text = txtTapeLent.Text & PadTape(Trim(Length(ii).Text))
        Next ii

    End Sub

    Private Sub Tab_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Tab_Renamed.Click
        'Allows the user to use enter as a tab

        System.Windows.Forms.SendKeys.Send("{TAB}")

    End Sub

    Private Sub ThumbTip(ByRef xyArcStart As ARMDIA1.XY, ByRef xyArcEnd As ARMDIA1.XY)
        Dim nOpp As Object
        Dim nAdj As Object
        ' This Calculates and Draws the curve at the
        ' right of the ThumbHole on the template
        Dim xyCen As ARMDIA1.XY
        Dim nEndAng, nRad, nStartAng, nDeltaAng As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nRad = ARMDIA1.FN_CalcLength(xyArcStart, xyArcEnd)
        'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nRad = nRad / 2
        'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nStartAng = ARMDIA1.FN_CalcAngle(xyArcStart, xyArcEnd)
        'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nStartAng = nStartAng * (ARMDIA1.PI / 180)
        'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nAdj = System.Math.Cos(nStartAng) * nRad
        'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nOpp = System.Math.Sin(nStartAng) * nRad
        'UPGRADE_WARNING: Couldn't resolve default property of object nAdj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xyCen.X = xyArcStart.X + nAdj
        'UPGRADE_WARNING: Couldn't resolve default property of object nOpp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xyCen.y = xyArcStart.y + nOpp
        'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nStartAng = ARMDIA1.FN_CalcAngle(xyCen, xyArcStart)
        'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nDeltaAng = -180
        'UPGRADE_WARNING: Couldn't resolve default property of object nDeltaAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nStartAng. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object nRad. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object fNum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(ARMDIA1.fNum, "hEnt = AddEntity(" & ARMDIA1.QQ & "arc" & ARMDIA1.QC & "xyStart.x +" & Str(xyCen.X) & ARMDIA1.CC & "xyStart.y +" & Str(xyCen.y) & ARMDIA1.CC & Str(nRad) & ARMDIA1.CC & Str(nStartAng) & ARMDIA1.CC & Str(nDeltaAng) & ");")
    End Sub

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        'This is enabled on Load
        'It is disabled in Link close
        'If the event happens then it is assumed that
        'for some reason the link between DRAFIX and This dialogue
        'has failed'
        'Therefor we "End" here

        Return
    End Sub

    Private Sub txtCircum_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCircum.Enter
        select_text_in_Box1(txtCircum)
    End Sub

    Private Sub txtCircum_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCircum.Leave
        If Not FN_CheckValue(txtCircum, "Thumb Circumference") Then
            Exit Sub
        ElseIf Val(txtCircum.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TCirc.Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtCircum))
        Else
            TCirc.Text = ""
        End If
    End Sub

    Private Sub txtCustFlapLength_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustFlapLength.Enter
        select_text_in_Box1(txtCustFlapLength)
    End Sub

    Private Sub txtCustFlapLength_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustFlapLength.Leave
        If Not FN_CheckValue(txtCustFlapLength, "Custom Flap") Then
            Exit Sub
        ElseIf Val(txtCustFlapLength.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            labCustFlapLength.Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtCustFlapLength))
        Else
            labCustFlapLength.Text = ""
        End If
    End Sub

    Private Sub txtFrontStrapLength_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFrontStrapLength.Enter
        select_text_in_Box1(txtFrontStrapLength)
    End Sub

    Private Sub txtFrontStrapLength_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFrontStrapLength.Leave
        If Not FN_CheckValue(txtFrontStrapLength, "Front Strap") Then
            Exit Sub
        ElseIf Val(txtFrontStrapLength.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            labFrontStrapLength.Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtFrontStrapLength))
        Else
            labFrontStrapLength.Text = ""
        End If
    End Sub

    Private Sub txtGauntletExtension_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtGauntletExtension.Enter
        select_text_in_Box1(txtGauntletExtension)
    End Sub

    Private Sub txtGauntletExtension_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtGauntletExtension.Leave
        If Not FN_CheckValue(txtGauntletExtension, "Gauntlet Extension") Then
            Exit Sub
        ElseIf Val(txtGauntletExtension.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            labGauntletExtension.Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtGauntletExtension))
        Else
            labGauntletExtension.Text = ""
        End If
    End Sub

    Private Sub txtPalmWrist_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPalmWrist.Enter
        select_text_in_Box1(txtPalmWrist)
    End Sub

    Private Sub txtPalmWrist_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPalmWrist.Leave
        If Not FN_CheckValue(txtPalmWrist, "Palm to Wrist") Then
            Exit Sub
        ElseIf Val(txtPalmWrist.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PWDist.Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtPalmWrist))
        Else
            PWDist.Text = ""
        End If
    End Sub

    Private Sub txtShoulderPleat1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtShoulderPleat1.Enter
        select_text_in_Box1(txtShoulderPleat1)
    End Sub

    Private Sub txtShoulderPleat1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtShoulderPleat1.Leave
        If Not FN_CheckValue(txtShoulderPleat1, "First Shoulder Pleat") Then
            Exit Sub
        End If
    End Sub

    Private Sub txtShoulderPleat2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtShoulderPleat2.Enter
        select_text_in_Box1(txtShoulderPleat2)
    End Sub

    Private Sub txtShoulderPleat2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtShoulderPleat2.Leave
        If Not FN_CheckValue(txtShoulderPleat2, "Second Shoulder Pleat") Then
            Exit Sub
        End If
    End Sub

    Private Sub txtStrapLength_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtStrapLength.Enter
        select_text_in_Box1(txtStrapLength)
    End Sub

    Private Sub txtStrapLength_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtStrapLength.Leave

        If Not FN_CheckValue(txtStrapLength, "Strap") Then
            Exit Sub
        ElseIf Val(txtStrapLength.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            labStrap.Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtStrapLength))
            txtStrap.Text = txtStrapLength.Text
        Else
            labStrap.Text = ""
        End If


        Dim w, X As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X = Val(txtLastTape.Text)
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        w = X - 1

        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If X > 0 Then

            If Length(w).Text <> "" And cboFabric.Text <> "" And cboPressure.Text <> "" And gms(w).Text <> "" And reds(w).Text <> "" Then

                txtFlap.Text = cboFlaps.Text

                PR_LastTapeRed(w, X, g_sModulus)

                PR_CheckFlaps(w)

                PR_SetRresults()

                PR_SetGlobals()

            End If
        End If

    End Sub

    Private Sub rdoLeft_CheckedChanged(sender As Object, e As EventArgs)
        'If (rdoLeft.Checked = True) Then
        '    txtSleeve.Text = "Left"
        'Else
        '    txtSleeve.Text = "Right"
        'End If
    End Sub

    Private Sub txtThumb_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtThumb.Enter
        select_text_in_Box1(txtThumb)
    End Sub

    Private Sub txtThumb_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtThumb.Leave
        If Not FN_CheckValue(txtThumb, "Thumb") Then
            Exit Sub
        ElseIf Val(txtThumb.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            TLen.Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtThumb))
        Else
            TLen.Text = ""
        End If
    End Sub

    Private Sub txtWaistCir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistCir.Enter
        select_text_in_Box1(txtWaistCir)
    End Sub

    Private Sub txtWaistCir_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWaistCir.Leave
        If Not FN_CheckValue(txtWaistCir, "Style-D waist circumference") Then
            Exit Sub
        ElseIf Val(txtWaistCir.Text) > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            labWaistCir.Text = ARMEDDIA1.fnInchesToText(cmtoinches(txtinchflag.Text, txtWaistCir))
        Else
            labWaistCir.Text = ""
        End If

    End Sub

    Private Sub txtWristPleat1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWristPleat1.Enter
        select_text_in_Box1(txtWristPleat1)
    End Sub

    Private Sub txtWristPleat1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWristPleat1.Leave
        If Not FN_CheckValue(txtWristPleat1, "First Wrist Pleat") Then
            Exit Sub
        End If
    End Sub

    Private Sub txtWristPleat2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWristPleat2.Enter
        select_text_in_Box1(txtWristPleat2)
    End Sub

    Private Sub txtWristPleat2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWristPleat2.Leave
        If Not FN_CheckValue(txtWristPleat2, "Second Wrist Pleat") Then
            Exit Sub
        End If
    End Sub

    Private Sub TypeDetails(ByRef Profile As ARMDIA1.curve, ByRef numtapes As Object)
        Dim n, nTextHt As Object
        Dim xyDetails As ARMDIA1.XY

        Dim sDetails, NL As Object
        Dim sSymbol As String
        Dim nSpacing As Double

        nSpacing = 0.2

        ARMDIA1.PR_SetTextData(1, 32, -1, -1, -1) 'Horiz-Left, Vertical-Bottom

        If Profile.X(Profile.n) < 6 Then
            ARMDIA1.PR_MakeXY(xyDetails, Profile.X(Profile.n) - 2, 2.5)
        Else
            ARMDIA1.PR_MakeXY(xyDetails, Profile.X(1) + 4, 2.5)
        End If
        If VB.Left(cboFlaps.Text, 4) = "Vest" Or VB.Left(cboFlaps.Text, 4) = "Body" Then
            ARMDIA1.PR_MakeXY(xyDetails, Profile.X(Profile.n) + 1.5, 1.75)
        End If

        'Main details in Green on layer notes
        ARMDIA1.PR_SetLayer("Notes")
        Dim sText As String

        'sText = txtSleeve.Text & "\n" & txtPatientName.Text & "\n" & ARMDIA1.g_sWorkOrder & "\n" & Trim(Mid(txtFabric.Text, 4))
        'ARMDIA1.PR_DrawText(sText, xyDetails, 0.125)
        sText = txtSleeve.Text & Chr(10) & txtPatientName.Text & Chr(10) & txtWorkOrder.Text & Chr(10) & Trim(Mid(txtFabric.Text, 4))
        ARMDIA1.PR_DrawMText(sText, xyDetails, False)

        'Other patient details in black on layer construct
        ARMDIA1.PR_SetLayer("Construct")

        ARMDIA1.PR_MakeXY(xyDetails, xyDetails.X, xyDetails.y - 0.8)
        'sText = txtFileNo.Text & "\n" & txtDiagnosis.Text & "\n" & txtAge.Text & "\n" & txtSex.Text
        'ARMDIA1.PR_DrawText(sText, xyDetails, 0.125)
        sText = txtFileNo.Text & Chr(10) & txtDiagnosis.Text & Chr(10) & txtAge.Text & Chr(10) & txtSex.Text
        ARMDIA1.PR_DrawMText(sText, xyDetails, False)

    End Sub

    Private Function WristRed(ByRef lblred As Object) As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Val(lblred) > 14 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object WristRed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            WristRed = 14
            'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf Val(lblred) < 10 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object WristRed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            WristRed = 10
        End If
    End Function

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
        ARMDIA1.PR_MakeXY(ARMDIA1.xyInsertion, ptStart.X, ptStart.Y)
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        txtSleeve.Text = "Left"
        If TabControl1.SelectedIndex = 1 Then
            txtSleeve.Text = "Right"
            ARMDIA1.g_sSide = txtSleeve.Text
        End If
        ''Added for #211 in the issue list
        Label96.Text = txtSleeve.Text.ToLower() + " arm"
        If g_sAlreadyLoad = False Then
            g_sAlreadyLoad = True
            RightDialogLoad()
        End If
    End Sub

    Private Sub RightLength_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles RightLength.Leave
        Dim Index As Short = RightLength.GetIndex(eventSender)
        Dim inchbit As Double
        If (RightChangeChecker(Index + 1) <> RightLength(Index).Text) Then

            If Me.Visible = True Then ARMDIA1.g_Modified = True

            ' Delete tape values
            If RightLength(Index).Text = "" And Rightmms(Index).Text <> "" And Rightgms(Index).Text <> "" And Rightreds(Index).Text <> "" And RightInchText(Index).Text <> "" Then
                Rightmms(Index).Text = ""
                Rightgms(Index).Text = ""
                Rightreds(Index).Text = ""
                RightInchText(Index).Text = ""
                ''PR_GetStartEndTapes()
                PR_GetRightStartEndTapes()
                'Reset the first and last tape indicators
                If RightcboProximalTape.SelectedIndex = 0 Then RightcboProximalTape_SelectedIndexChanged(RightcboProximalTape, New System.EventArgs())
                If RightcboDistalTape.SelectedIndex = 0 Then RightcboDistalTape_SelectedIndexChanged(RightcboDistalTape, New System.EventArgs())

            End If
            ' Check for numeric values
            If Not FN_CheckValue(RightLength(Index), "Tape") Then Exit Sub
            ' check for too large a number
            If (Val(RightLength(Index).Text) > 99 And Righttxtinchflag.Text = "cm") Or (Val(RightLength(Index).Text) > 50 And Righttxtinchflag.Text = "inches") Then
                MsgBox("Tape length entered Is too Large", 48, "Arm Details")
                RightLength(Index).Focus()
                Exit Sub
            End If

            ' check for too small a number
            ' tapelen from above
            If Val(RightLength(Index).Text) < 1 And Len(RightLength(Index).Text) > 0 Then
                MsgBox("Tape length entered Is too Small", 48, "Arm Details")
                RightLength(Index).Focus()
                Exit Sub
            End If
            If Val(RightLength(Index).Text) > 0 Then
                inchbit = cmtoinches(Righttxtinchflag.Text, RightLength(Index))
                If inchbit = -1 Then
                    RightLength(Index).Focus()
                    Exit Sub
                End If
                RightInchText(Index).Text = ARMEDDIA1.fnInchesToText(inchbit)
                RightMeasurements(Index + 1) = cmtoinches(Righttxtinchflag.Text, RightLength(Index))

                'Exit if no fabric or no pressure has been selected
                'We also supress calculations if this is an empty sleeve
                'the flag EmptySleeve is set in Form "LinkClose" & Command "Calculate"
                'I.E. don't do any calculations
                If RightcboFabric.Text = "" Or RightcboPressure.Text = "" Or EmptySleeve = True Then Exit Sub

                'Recalculate etc
                PR_GetRightPressure()
                PR_GetRightFabric()
                PR_CheckRightValues()
                PR_GetRightGauntletDetails()
                PR_GetRightStartEndTapes()
                PR_CalculateRightMMSGmsReds(Index)
                PR_ReSetRightMMs(Index, g_sRightModulus)
                ''--------------PR_SetRresults()
                '--------PR_SetGlobals()
                'check for extra measurements in Gauntlet
                PR_RightGauntletHoleCheck()
                'If a tape has been added then reset the tape indicators
                If RightcboProximalTape.SelectedIndex = 0 Then RightcboProximalTape_SelectedIndexChanged(RightcboProximalTape, New System.EventArgs())
                If RightcboDistalTape.SelectedIndex = 0 Then RightcboDistalTape_SelectedIndexChanged(RightcboDistalTape, New System.EventArgs())

            End If
            RightChangeChecker(Index + 1) = RightLength(Index).Text
        End If
    End Sub
    Private Sub PR_RightGauntletHoleCheck()

        Dim Gauntletpwdiff, count As Object
        If RightchkGauntlets.Checked = True Then
            'UPGRADE_WARNING: Couldn't resolve default property of object Gauntletpwdiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Gauntletpwdiff = (Val(strWristNo) - Val(strPalmNo))
            'UPGRADE_WARNING: Couldn't resolve default property of object Gauntletpwdiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Gauntletpwdiff = Gauntletpwdiff - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object Gauntletpwdiff. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Gauntletpwdiff = 1 Then
                g_sRightHoleCheck = "True"
                RightLength(strPalmNo).Text = ""
                'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(txtPalmNo + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                RightMeasurements(CDbl(strPalmNo) + 1) = ""
                Rightgms(strPalmNo).Text = ""
                Rightreds(strPalmNo).Text = ""
                Rightmms(strPalmNo).Text = ""
                RightInchText(strPalmNo).Text = ""
            Else
                g_sRightHoleCheck = "False"
            End If
        End If

    End Sub
    Private Sub PR_ReSetRightMMs(ByRef w As Object, ByRef Modulus As Object)
        Dim tempmms As Object
        Dim newmms As Object
        Dim Backgram As Object
        Dim givenreduction, X As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X = w + 1
        If RightLength(w).Text <> "" And Rightmms(w).Text <> "" And Rightreds(w).Text <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object givenreduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            givenreduction = Val(CStr(CDbl(Rightreds(w).Text) - 9))
            'UPGRADE_WARNING: Couldn't resolve default property of object givenreduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If givenreduction > 32 Or givenreduction < 1 Then
                MsgBox("Reduction Value Not Recognised", 48, "Arm Details")
                Exit Sub
            End If
            If VB.Left(RightcboFabric.Text, 3) = "Pow" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object givenreduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modulus, givenreduction). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Backgram = RightPow(Modulus, givenreduction)
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Val(Rightgms(w).Text) <> Backgram Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    newmms = Backgram / RightMeasurements(X)
                    'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tempmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tempmms = ARMDIA1.round(newmms)
                    'UPGRADE_WARNING: Couldn't resolve default property of object tempmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tempmms = (newmms - tempmms) * 10
                    'UPGRADE_WARNING: Couldn't resolve default property of object tempmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If tempmms >= 5 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Rightmms(w).Text = CStr(ARMDIA1.round(newmms) + 1)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Rightmms(w).Text = CStr(ARMDIA1.round(newmms))
                    End If

                End If
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object givenreduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Modulus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modulus, givenreduction). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Backgram = RightBob(Modulus, givenreduction)
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Val(Rightgms(w).Text) <> Backgram Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    newmms = Backgram / RightMeasurements(X)
                    'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object tempmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tempmms = ARMDIA1.round(newmms)
                    'UPGRADE_WARNING: Couldn't resolve default property of object tempmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    tempmms = (newmms - tempmms) * 10
                    'UPGRADE_WARNING: Couldn't resolve default property of object tempmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If tempmms >= 5 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Rightmms(w).Text = CStr(ARMDIA1.round(newmms) + 1)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object newmms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Rightmms(w).Text = CStr(ARMDIA1.round(newmms))
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub PR_RightInitaliseEmpty(ByRef w As Object)
        'initalise empty fields on front end
        If RightLength(w).Text = "" Then
            Rightmms(w).Text = ""
            Rightgms(w).Text = ""
            Rightreds(w).Text = ""
            RightInchText(w).Text = ""
        End If
    End Sub
    Private Sub PR_RightInitaliseArrays(ByRef w As Object, ByRef X As Object)
        If RightLength(w).Text = "" Then
            RightMeasurements(X) = ""
            Righttapewidths(X) = ""
        End If
    End Sub
    Public Sub PR_SetRightMMs(ByRef w As Object, ByRef Amm As Object, ByRef Bmm As Object, ByRef Cmm As Object, ByRef Dmm As Object)
        Dim GenMMs As Object = ""
        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If RightLength(w).Text <> "" And w >= CDbl(strFirstTape) - 1 And w <= CDbl(strLastTape) - 1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (w = CDbl(strFirstTape) - 1) Or (w = 9) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Amm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GenMMs = Amm
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (w > 11) And (w < (CDbl(strSecondLastTape) - 1)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Amm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GenMMs = Amm
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf ((w > CDbl(strFirstTape) - 1) And (w < 9)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Bmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GenMMs = Bmm
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (w = 11) Or (w = (CDbl(strSecondLastTape) - 1)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Cmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GenMMs = Cmm
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (w = 10) Or (w = (CDbl(strLastTape) - 1)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Dmm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GenMMs = Dmm
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (strWristNo <> "" And w = Val(strWristNo) - 1) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Amm. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GenMMs = Amm
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object GenMMs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Rightmms(w).Text = CStr(Val(GenMMs))
        Else
            Rightmms(w).Text = ""
        End If
    End Sub
    Private Sub PR_RightReductions(ByRef w As Object, ByRef X As Object)
        Dim Gram As Object
        If RightLength(w).Text <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Gram = Fix(Val(Rightmms(w).Text) * Val(RightMeasurements(X)))
            'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Gram <> 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Gram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Rightgms(w).Text = CStr(ARMDIA1.round(Gram))
                'UPGRADE_WARNING: Couldn't resolve default property of object GetReduction(g_sModulus, Gram, g_sFabnam). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object GetReduction(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Rightreds(w).Text = GetRightReduction(g_sRightModulus, Gram, RightcboFabric.Text)
            End If
        End If

    End Sub
    Private Function GetRightReduction(ByRef Modul As Object, ByRef Gram As Object, ByRef fabnam As Object) As Object
        Dim tempred As Object = ""
        'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If VB.Left(fabnam, 3) = "Pow" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object PowReduction(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object tempred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            tempred = RightPowReduction(Modul, Gram)
            'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ElseIf VB.Left(fabnam, 3) = "Bob" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object BobReduction(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object tempred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            tempred = RightBobReduction(Modul, Gram)
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object tempred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object GetReduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetRightReduction = tempred
    End Function
    Private Function RightBobReduction(ByRef Modul As Object, ByRef lblGram As Object) As Object
        Dim lblred As Object = ""
        Dim j As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        one = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object two. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        two = 2
        'UPGRADE_WARNING: Couldn't resolve default property of object nine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nine = 9
        'UPGRADE_WARNING: Couldn't resolve default property of object eight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        eight = 8
        'UPGRADE_WARNING: Couldn't resolve default property of object ten. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ten = 10
        'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        j = one
        'UPGRADE_WARNING: Couldn't resolve default property of object lblGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object RightBob(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (Val(RightBob(Modul, 1)) = Val(lblGram)) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object ten. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            lblred = ten
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object two. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        j = two
        'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While j <= 22
            'UPGRADE_WARNING: Couldn't resolve default property of object lblGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object RightBob(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (Val(RightBob(Modul, j)) = Val(lblGram)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object nine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                lblred = Val(j + nine)
                'UPGRADE_WARNING: Couldn't resolve default property of object lblGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object RightBob(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (Val(lblGram) - Val(RightBob(Modul, j - one))) <= (Val(RightBob(Modul, j)) - Val(lblGram)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object eight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                lblred = Val(j + eight)
                Exit Do
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            j = j + one
        Loop
        'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object BobReduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        RightBobReduction = lblred
    End Function
    Private Function RightPowReduction(ByRef Modul As Object, ByRef lblGram As Object) As Object
        Dim lblred As Object = ""
        Dim j As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        one = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object two. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        two = 2
        'UPGRADE_WARNING: Couldn't resolve default property of object nine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        nine = 9
        'UPGRADE_WARNING: Couldn't resolve default property of object eight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        eight = 8
        'UPGRADE_WARNING: Couldn't resolve default property of object ten. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ten = 10
        'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        j = one
        'UPGRADE_WARNING: Couldn't resolve default property of object lblGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object RightPow(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (Val(RightPow(Modul, 1)) = Val(lblGram)) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object ten. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            lblred = ten
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object two. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        j = two
        'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While j <= 22
            'UPGRADE_WARNING: Couldn't resolve default property of object lblGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object RightPow(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (Val(RightPow(Modul, j)) = Val(lblGram)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object nine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                lblred = Val(j + nine)
                'UPGRADE_WARNING: Couldn't resolve default property of object lblGram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object RightPow(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf (Val(lblGram) - Val(RightPow(Modul, j - one))) <= (Val(RightPow(Modul, j)) - Val(lblGram)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object eight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                lblred = Val(j + eight)
                Exit Do
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object one. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            j = j + one
        Loop
        'UPGRADE_WARNING: Couldn't resolve default property of object lblred. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object PowReduction. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        RightPowReduction = lblred
    End Function
    Private Sub PR_RightStump(ByRef w As Object, ByRef X As Object)
        Dim Backgram As Object
        If RightLength(w).Text <> "" Then
            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If (w = (CDbl(strFirstTape) - 1)) And RightchkStump.Checked = True Then
                Rightreds(w).Text = CStr(14)
                If VB.Left(RightcboFabric.Text, 3) = "Pow" Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object Pow(txtModulus, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Backgram = RightPow(CInt(g_sRightModulus), 5)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object Bob(txtModulus, 5). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Backgram = RightBob(CInt(g_sRightModulus), 5)
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Measurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                ''------------------------txtStump.Text = "True"
            End If
        End If
    End Sub
    Private Sub PR_LastRightCircum(ByRef w As Object, ByRef X As Object)
        Dim temptapes1, temptapes2 As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If w > 9 Then

            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If RightLength(w).Text <> "" And RightLength(w - 1).Text <> "" Then
                'check circum of last tape
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object temptapes1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                temptapes1 = cmtoinches(Righttxtinchflag.Text, RightLength(w))

                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object temptapes2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                temptapes2 = cmtoinches(Righttxtinchflag.Text, RightLength(w - 1))

                'UPGRADE_WARNING: Couldn't resolve default property of object temptapes2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object temptapes1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If (w = (Val(strLastTape) - 1)) And (Val(txtAge.Text) > 10) And (Val(temptapes1 - temptapes2) > 1.5) And (Val(strLastTape) > 11) Then

                    MsgBox("Last tape is measured over axilla. Exclude measurement", 48, "Arm Details")

                    ' Exit Sub
                    'UPGRADE_WARNING: Couldn't resolve default property of object temptapes2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object temptapes1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf (w = (Val(strLastTape) - 1)) And (Val(txtAge.Text) < 11) And (Val(temptapes1 - temptapes2) > 0.75) And (Val(strLastTape) > 11) Then

                    MsgBox("Last tape is measured over axilla. Exclude measurement", 48, "Arm Details, child")
                End If
            End If
        End If
    End Sub
    Private Sub PR_RightDetachable(ByRef w As Object, ByRef X As Object, ByRef Modul As Object, ByRef fabnam As Object)
        Dim Backgram As Object
        If RightLength(w).Text <> "" Then
            'CheckDetachable Gauntlet
            If RightchkDetachable.CheckState = 1 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If X = Val(strWristNo) And X = Val(strSecondLastTape) Then
                    Rightreds(w).Text = CStr(10)
                    'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If VB.Left(fabnam, 3) = "Pow" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modul, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Backgram = RightPow(Modul, 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If Val(Rightgms(w).Text) <> Backgram Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object RightMeasurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                        End If
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modul, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        Backgram = RightBob(Modul, 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If Val(Rightgms(w).Text) <> Backgram Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object RightMeasurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                        End If
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf X = Val(strWristNo) + 1 And X = Val(strLastTape) Then
                    'If length of wrist is same as wrist+1 then assume detachable gauntlet
                    'originally based on 2 tapes
                    'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If RightLength(w).Text = RightLength(w - 1).Text Then
                        Rightreds(w).Text = CStr(5)
                        Rightgms(w).Text = ""
                        Rightmms(w).Text = ""
                    Else
                        Rightreds(w).Text = CStr(10)
                        'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If VB.Left(fabnam, 3) = "Pow" Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modul, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Backgram = RightPow(Modul, 1)
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If Val(Rightgms(w).Text) <> Backgram Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object RightMeasurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                            End If
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modul, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Backgram = RightBob(Modul, 1)
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If Val(Rightgms(w).Text) <> Backgram Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object RightMeasurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                            End If
                        End If
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ElseIf X = Val(strWristNo) And X = Val(strLastTape) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If RightLength(w + 1).Text = "" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        RightLength(w + 1).Text = RightLength(w).Text
                        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object RightMeasurements(X). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object RightMeasurements(X + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        RightMeasurements(X + 1) = RightMeasurements(X)
                        'UPGRADE_WARNING: Couldn't resolve default property of object cmtoinches(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        InchText(w + 1).Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RightLength(w)))
                        Rightreds(w).Text = CStr(10)
                        'UPGRADE_WARNING: Couldn't resolve default property of object fabnam. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If VB.Left(fabnam, 3) = "Pow" Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Pow(Modul, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Backgram = RightPow(Modul, 1)
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If Val(Rightgms(w).Text) <> Backgram Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object RightMeasurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                            End If
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object Modul. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Bob(Modul, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            Backgram = RightBob(Modul, 1)
                            'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If Val(Rightgms(w).Text) <> Backgram Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object RightMeasurements(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object Backgram. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                            End If
                        End If
                    Else
                        Rightreds(w).Text = CStr(5)
                        Rightgms(w).Text = ""
                        Rightmms(w).Text = ""
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub PR_CalculateRightMMSGmsReds(ByRef Index As Short)
        Dim w, X As Object
        w = 0
        X = 1
        Do While X < 19
            Do While w < 18
                PR_RightInitaliseEmpty(w) 'Zero any empty fields
                PR_RightInitaliseArrays(w, X)
                If (RightLength(Index).Text <> "" And Rightmms(Index).Text = "") Or (RightLength(Index).Text <> "" And g_sRightPressureChange = "True") Then
                    PR_SetRightMMs(w, g_sRightAmm, g_sRightBmm, g_sRightCmm, g_sRightDmm)
                End If
                PR_RightReductions(w, X)
                PR_RightWristRed(w, X, g_sRightModulus)
                PR_RightPalmRed(w, X, g_sRightModulus)
                PR_RightStump(w, X)
                PR_LastRightCircum(w, X)
                PR_RightLastTapeRed(w, X, g_sRightModulus)
                '-------------------PR_CheckRightFlaps(w)
                PR_RightDetachable(w, X, g_sRightModulus, RightcboFabric.Text)
                w = w + 1
                X = X + 1
            Loop
        Loop
    End Sub
    Private Sub RightGauntlet(ByRef NoThumb As Object, ByRef Palm As Object, ByRef Wrist As Object, ByRef PalmWrist As Object, ByRef thumlength As Object, ByRef Sex As Object, ByRef age As Object, ByRef Enclosed As Object, ByRef Circum As Object, ByRef Detach As Object, ByRef Righttxtinchflag As Object)
        Dim tempCircum As Object
        Dim txtThumblen As Object
        Dim thum As Object
        Dim thumdist As Object
        Dim temppalm As Object
        Dim tempwrist As Object
        tempwrist = Wrist
        temppalm = Palm
        thumdist = PalmWrist
        If thumdist = "" Then RighttxtPalmWrist.Text = 0
        thumdist = cmtoinches(Righttxtinchflag, RighttxtPalmWrist)
        ' Check values entered
        If Palm = "" Then
            MsgBox("Palm tape number has Not been entered.", 48, "Gauntlet Details")
            Exit Sub
        End If
        If Wrist = "" Then
            MsgBox("Wrist tape number has Not been entered.", 48, "Gauntlet Details")
            Exit Sub
        End If
        If PalmWrist = "" Then
            MsgBox("Palm To Wrist Distance has Not been entered.", 48, "Gauntlet Details")
            Exit Sub
        End If
        ' Thumbhole Dist
        If thumdist > 2.5 Then
            strThumbDist = CStr(2.5)
        Else
            strThumbDist = thumdist
            '-------------cmtoinches(Righttxtinchflag, RighttxtPalmWrist)
        End If
        ' End Thumbhole Dist

        'Thumb length
        If Val(thumlength) > 0 Then
            thum = thumlength
            thum = cmtoinches(Righttxtinchflag, RighttxtThumb)
            strThumbLen = thum
        Else : strThumbLen = CStr(0)
            thum = 0
        End If

        If thum = 0 And Val(age) > 14 Then
            thum = 0.75
            strThumbLen = thum
        ElseIf CDbl(strThumbLen) = 0 And Val(age) <= 14 Then
            thum = 0.5
            strThumbLen = thum
        End If
        'End Thumb length

        'Enclosed Thumb
        If Enclosed = 1 And Val(age) > 14 Then
            If thumlength = "" Then
                MsgBox("Thumb Length has Not been entered.", 48, "Gauntlet Details")
                Exit Sub
            End If
            txtThumblen = Val(thumlength) - 0.25
        ElseIf Enclosed = 1 And Val(age) < 15 Then
            If thumlength = 0 Then
                MsgBox("Thumb Length has Not been entered.", 48, "Gauntlet Details")
                Exit Sub
            End If
            txtThumblen = Val(thumlength) - 0.125
        End If
        'End Enclosed Thumb

        ' Thumb Circumference
        If NoThumb = 0 And Val(Circum) = 0 Then
            MsgBox("Thumb Circumference has Not been entered.", 48, "Gauntlet Details")
            Exit Sub
        ElseIf NoThumb = 0 And Val(Circum) > 0 Then
            tempCircum = cmtoinches(Righttxtinchflag, RighttxtCircum)
            tempCircum = Val(tempCircum) / 2
            strTCircum = tempCircum
        End If
        ' End Thumb Circumference
        If Val(Palm) > Val(Wrist) Then
            MsgBox("Palm Tape No Should be less than Wrist Tape No. Enter New values And re-calculate. ", 48, "Gauntlet Details")
            Exit Sub
        End If
        ' check palm and wrist tapes
    End Sub
    Private Sub PR_GetRightGauntletDetails()
        '       *** Start Gauntlet Details ***
        If RightchkGauntlets.Checked = True Then
            strWristNo = CStr(RightcboWrist.SelectedIndex + 1)
            If Val(strWristNo) = 0 Then strWristNo = ""

            strPalmNo = CStr(RightcboPalm.SelectedIndex + 1)
            If Val(strPalmNo) = 0 Then strPalmNo = ""

            Call RightGauntlet(RightchkNoThumb.Checked, strPalmNo, strWristNo, RighttxtPalmWrist.Text, RighttxtThumb.Text,
                          txtSex.Text, txtAge.Text, RightchkEnclosed.Checked, RighttxtCircum.Text, RightchkDetachable.Checked, Righttxtinchflag.Text)
            '---------------------------
            'ElseIf RightchkGauntlets.Checked = False Then
            '    txtGauntletFlag.Text = "False"
        End If

        ''       *** End Gauntlet Details ***

    End Sub
    Private Sub PR_CheckRightValues()
        ''   *** Check values have been entered ***
        Dim contents, i, total As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        i = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object contents. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        contents = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object total. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        total = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While i < 18
            'UPGRADE_WARNING: Couldn't resolve default property of object contents. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            contents = Val(RightLength(i).Text)
            'UPGRADE_WARNING: Couldn't resolve default property of object contents. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object total. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            total = total + contents
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            i = i + 1
        Loop
        'UPGRADE_WARNING: Couldn't resolve default property of object total. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If total = 0 Then
            MsgBox("No tape measurements have been entered!", 48, "Arm Details")
            Exit Sub
        End If
    End Sub
    Private Sub PR_GetRightPressure()
        If RightcboPressure.Text = "15mm" Then
            g_sRightAmm = 12
            g_sRightBmm = 15
            g_sRightCmm = 10
            g_sRightDmm = 8
        ElseIf RightcboPressure.Text = "20mm" Then
            g_sRightAmm = 16
            g_sRightBmm = 20
            g_sRightCmm = 13
            g_sRightDmm = 10
        ElseIf RightcboPressure.Text = "25mm" Then
            g_sRightAmm = 20
            g_sRightBmm = 25
            g_sRightCmm = 17
            g_sRightDmm = 13
        End If
        '--------------txtMM.Text = RightcboPressure.Text
    End Sub
    Private Sub RightDialogLoad()
        Try
            Dim i As Object
            i = 0
            Do While i < 18
                Rightreds(i).Text = ""
                Rightgms(i).Text = ""
                Rightmms(i).Text = ""
                RightLength(i).Text = ""
                i = i + 1
            Loop
            Righttxtinchflag.Text = ""
            RightcboFabric.Text = ""
            RightcboPressure.Text = ""
            RightcboFlaps.Text = ""
            RighttxtStrapLength.Text = ""

            RighttxtGauntletExtension.Text = ""
            RighttxtFrontStrapLength.Text = ""
            RighttxtCustFlapLength.Text = ""
            RighttxtWristPleat1.Text = ""
            RighttxtWristPleat2.Text = ""
            RighttxtShoulderPleat1.Text = ""
            RighttxtShoulderPleat2.Text = ""

            RightchkNone.Checked = True
            RightchkStump.Checked = False
            RightfrmGauntlet.Enabled = 0
            RightchkGauntlets.Checked = False
            RightchkDetachable.CheckState = System.Windows.Forms.CheckState.Unchecked
            RightchkEnclosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            RightchkNoThumb.CheckState = System.Windows.Forms.CheckState.Unchecked
            RighttxtPalmWrist.Text = ""
            RighttxtCircum.Text = ""
            RighttxtThumb.Text = ""
            RighttxtWaistCir.Text = ""

            g_sRightPressureChange = "False"
            strWristNo = ""
            strPalmNo = ""
            strThumbDist = ""
            strThumbLen = ""
            strTCircum = ""
            ARMDIA1.g_sPathJOBST = fnPathJOBST()
            Dim ii As Object
            Dim fileNum As Object
            Dim textline As Object
            fileNum = FreeFile()
            Dim sSettingsPath As String = fnGetSettingsPath("LookupTables")
            FileOpen(fileNum, sSettingsPath & "\Pressure.dat", Microsoft.VisualBasic.OpenMode.Input)
            Do While Not EOF(fileNum)
                textline = LineInput(fileNum)
                RightcboPressure.Items.Add(textline)
            Loop
            FileClose()
            fileNum = FreeFile()
            FileOpen(fileNum, sSettingsPath & "\fabric.dat", Microsoft.VisualBasic.OpenMode.Input)
            Do While Not EOF(fileNum)
                textline = LineInput(fileNum)
                RightcboFabric.Items.Add(textline)
            Loop
            FileClose()

            fileNum = FreeFile()
            FileOpen(fileNum, sSettingsPath & "\flaps.dat", Microsoft.VisualBasic.OpenMode.Input)
            Do While Not EOF(fileNum)
                textline = LineInput(fileNum)
                RightcboFlaps.Items.Add(textline)
            Loop
            FileClose()

            'Load Palm Tapes and Wrist
            Dim g_sTextList As String = " -6-4½ -3-1½  0 1½  3 4½  6 7½  910½ 1213½ 1516½ 1819½ 2122½ 2425½ 2728½ 3031½ 3334½ 36"
            For ii = 0 To 9
                RightcboWrist.Items.Add(LTrim(Mid(g_sTextList, (ii * 3) + 1, 3)))
                RightcboPalm.Items.Add(LTrim(Mid(g_sTextList, (ii * 3) + 1, 3)))
            Next ii

            RightcboDistalTape.Items.Add("1St")
            For ii = 0 To 17
                RightcboDistalTape.Items.Add(LTrim(Mid(g_sTextList, (ii * 3) + 1, 3)))
            Next ii

            RightcboProximalTape.Items.Add("Last")
            For ii = 17 To 0 Step -1
                RightcboProximalTape.Items.Add(LTrim(Mid(g_sTextList, (ii * 3) + 1, 3)))
            Next ii

            Dim j, offset As Object
            Dim Modulus, Reduction As String
            Dim indi(19) As Object
            fileNum = FreeFile()
            FileOpen(fileNum, sSettingsPath & "\bobinnet.dat", Microsoft.VisualBasic.OpenMode.Input)
            i = 1
            Do While Not EOF(fileNum)
                If (i <= 19) Then
                    textline = LineInput(fileNum)
                    Modulus = textline.ToString().Substring(0, 5).Trim()
                    Reduction = textline.ToString().Substring(6).Trim()
                    indi(i) = Modulus
                    offset = 1
                    Dim sString() As String = Reduction.Replace("  ", " ").Trim().Split(" ")
                    Dim colReduction As List(Of String) = New List(Of String)(sString)
                    j = 1
                    For Each tempa As String In colReduction
                        If j <= 23 Then
                            If (tempa.Length > 1) Then
                                RightBob(i, j) = tempa
                                j += 1
                            End If
                        End If
                    Next
                    i += 1
                End If
            Loop
            FileClose()

            Dim q, P, offset2 As Object
            Dim modulus2, Reduction2 As String
            Dim pindi(19) As Object
            fileNum = FreeFile()
            FileOpen(fileNum, sSettingsPath & "\powernet.dat", Microsoft.VisualBasic.OpenMode.Input)

            P = 1
            Do While Not EOF(fileNum)
                textline = LineInput(fileNum)
                modulus2 = textline.ToString().Substring(0, 5).Trim()
                Reduction2 = textline.ToString().Substring(6).Trim()
                pindi(P) = modulus2
                offset2 = 1

                Dim sString() As String = Reduction2.Replace("  ", " ").Trim().Split(" ")
                Dim colReduction2 As List(Of String) = New List(Of String)(sString)
                q = 1
                For Each tempa As String In colReduction2
                    If (tempa.Length > 1) Then
                        RightPow(P, q) = tempa
                        q += 1
                    End If
                Next
                P += 1
            Loop
            FileClose()

            Dim fileNo As String = "", patient As String = "", diagnosis As String = "", age As String = "", sex As String = ""
            Dim workOrder As String = "", tempDate As String = "", tempEng As String = "", units As String = ""
            Dim obj As New BlockCreation.BlockCreation
            Dim blkId As ObjectId = New ObjectId()
            blkId = obj.LoadBlockInstance()
            If (blkId.IsNull()) Then
                MsgBox("Can't find Patient Details", 48, "Arm Details Dialog")
                g_bIsClose = True
                Me.Close()
                Exit Sub
            End If

            obj.BindAttributes(blkId, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units)
            RighttxtDiagnosis.Text = diagnosis
            Righttxtinchflag.Text = units
            ''RightcboFabric.Text = txtFabric.Text
            readRightDWGInfo()
            RightForm_LinkClose()
            g_sRightArmValues = FN_RightArmValuesString()
        Catch ex As Exception
            g_bIsClose = True
            Me.Close()
        End Try
    End Sub
    Private Sub RightForm_LinkClose()
        Dim Index As Object
        Dim w, X, i, k As Object
        Dim nValue As Double
        Dim nAge As Short
        Dim sDiag As String

        nAge = Val(txtAge.Text)
        sDiag = UCase(Mid(RighttxtDiagnosis.Text, 1, 6))
        If RightcboPressure.Text = "" Then
            'Arterial Insufficiency
            If sDiag = "ARTERI" Then
                RightcboPressure.Text = "15mm"
                'Assist fluid dynamics
            ElseIf sDiag = "ASSIST" Then
                RightcboPressure.Text = "20mm"
                'Blood Clots
            ElseIf sDiag = "BLOOD " Then
                RightcboPressure.Text = "20mm"
                'Burns
            ElseIf sDiag = "BURNS " Or sDiag = "BURNS" Then
                RightcboPressure.Text = "15mm"
                'Cancer
            ElseIf sDiag = "CANCER" Then
                RightcboPressure.Text = "20mm"
                'Cardial Vascular Arrest
            ElseIf sDiag = "CARDIA" Then
                RightcboPressure.Text = "15mm"
                'Carpal Tunnel Syndrome
            ElseIf sDiag = "CARPAL" Then
                RightcboPressure.Text = "20mm"
                'Cellulitis
            ElseIf sDiag = "CELLUL" Then
                RightcboPressure.Text = "20mm"
                'Chronic Venous Insufficiency
            ElseIf sDiag = "CHRONI" Then
                RightcboPressure.Text = "20mm"
                'Heart Condition
            ElseIf sDiag = "HEART " Then
                RightcboPressure.Text = "15mm"
                'Hemangioma
            ElseIf sDiag = "HEMANG" Then
                RightcboPressure.Text = "15mm"
                'Lymphedema, Lymphedema 1+, Lymphedema 2+
            ElseIf sDiag = "LYMPHE" Then
                If InStr(RighttxtDiagnosis.Text, "2") > 0 Then
                    If nAge <= 70 Then
                        RightcboPressure.Text = "25mm"
                    Else
                        RightcboPressure.Text = "20mm"
                    End If
                Else
                    If nAge <= 70 Then
                        RightcboPressure.Text = "20mm"
                    Else
                        RightcboPressure.Text = "15mm"
                    End If
                End If
                'Night Wear
            ElseIf sDiag = "NIGHT " Then
                RightcboPressure.Text = "15mm"
                'Post Fracture
            ElseIf sDiag = "POST F" Then
                RightcboPressure.Text = "20mm"
                'Post Mastectomy
            ElseIf sDiag = "POST M" Then
                RightcboPressure.Text = "20mm"
                'Postphlebetic Syndrome
            ElseIf sDiag = "POSTPH" Then
                RightcboPressure.Text = "20mm"
                'Renal Disease (Kidney)
            ElseIf sDiag = "RENAL " Then
                RightcboPressure.Text = "20mm"
                'Skin Graft
            ElseIf sDiag = "SKIN G" Then
                RightcboPressure.Text = "15mm"
                'Stroke
            ElseIf sDiag = "STROKE" Then
                RightcboPressure.Text = "15mm"
                'Tendonitis
            ElseIf sDiag = "TENDON" Then
                RightcboPressure.Text = "15mm"
                'Thrombophlebitis
            ElseIf sDiag = "THROMB" Then
                RightcboPressure.Text = "20mm"
                'Trauma
            ElseIf sDiag = "TRAUMA" Then
                RightcboPressure.Text = "20mm"
                'Varicose Veins
            ElseIf sDiag = "VARICO" Then
                RightcboPressure.Text = "20mm"
                'Request for 30 m/m, 40 m/m and 50 m/m
            ElseIf sDiag = "REQUES" Then
                If InStr(RighttxtDiagnosis.Text, "30") > 0 Then
                    RightcboPressure.Text = "15mm"
                ElseIf InStr(RighttxtDiagnosis.Text, "40") > 0 Then
                    RightcboPressure.Text = "20mm"
                ElseIf InStr(RighttxtDiagnosis.Text, "50") > 0 Then
                    RightcboPressure.Text = "25mm"
                End If
                'Vein Removal
            ElseIf sDiag = "VEIN R" Then
                RightcboPressure.Text = "20mm"
            End If
        End If
        nValue = FN_GetNumber(RighttxtWristPleat2.Text, 1)
        If nValue > 0 Then RighttxtWristPleat2.Text = CStr(nValue) Else RighttxtWristPleat2.Text = ""
        nValue = FN_GetNumber(RighttxtWristPleat1.Text, 1)
        If nValue > 0 Then RighttxtWristPleat1.Text = CStr(nValue) Else RighttxtWristPleat1.Text = ""
        ' Strip_Bad_Char RighttxtShoulderPleat1
        'NB the order
        nValue = FN_GetNumber(RighttxtShoulderPleat2.Text, 1)
        If nValue > 0 Then RighttxtShoulderPleat2.Text = CStr(nValue) Else RighttxtShoulderPleat2.Text = ""
        nValue = FN_GetNumber(RighttxtShoulderPleat1.Text, 1)
        If nValue > 0 Then RighttxtShoulderPleat1.Text = CStr(nValue) Else RighttxtShoulderPleat1.Text = ""

        If RightchkStump.Checked = True Then
            RightchkStump_CheckedChanged(chkStump, New System.EventArgs())
        End If

        For i = 0 To 17
            k = i + 1
            RightChangeChecker(k) = RightLength(i).Text
        Next i

        If RightchkGauntlets.Checked = True Then
            RightlblPalm.Enabled = True
            RightlblWrist.Enabled = True
            RightlblPalmWrist.Enabled = True
            RightlblThumb.Enabled = True
            RightlblCircum.Enabled = True
            RightlblGauntletExtension.Enabled = True
            RightcboPalm.Enabled = True
            RightcboWrist.Enabled = True
            RighttxtPalmWrist.Enabled = True
            RighttxtCircum.Enabled = True
            RighttxtThumb.Enabled = True
            RighttxtGauntletExtension.Enabled = True
            RightchkEnclosed.Enabled = True
            RightchkDetachable.Enabled = True
            RightchkNoThumb.Enabled = True
        ElseIf RightchkStump.Checked <> True Then
            RightchkNone.Checked = 1
            RightchkNone_CheckedChanged(RightchkNone, New System.EventArgs())
        End If

        If Right_optProximalTape_0.Checked = True Then
            RightoptProximalTape(0).Checked = True
            RightoptProximalTape_CheckedChanged(RightoptProximalTape.Item(0), New System.EventArgs())
        Else
            RightoptProximalTape(1).Checked = True
            RightoptProximalTape_CheckedChanged(RightoptProximalTape.Item(1), New System.EventArgs())
            'Back Strap 
            If Val(RighttxtStrapLength.Text) > 0 Then
                RightlabStrap.Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtStrapLength))
            End If
            'Front Strap 
            If Val(RighttxtFrontStrapLength.Text) > 0 Then
                RightlabFrontStrapLength.Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtFrontStrapLength))
            End If
            'Requested Flap RightLength
            If Val(RighttxtCustFlapLength.Text) > 0 Then
                RightlabCustFlapLength.Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtCustFlapLength))
            End If
            'Style "D" Waist Circum
            If Val(RighttxtWaistCir.Text) > 0 Then
                RightlabWaistCir.Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtWaistCir))
            End If
            RightcboFlaps_SelectedIndexChanged(RightcboFlaps, New System.EventArgs())
        End If

        RightEmptySleeve = True
        For w = 0 To 17
            X = w + 1
            If RightLength(w).Text <> "" Then
                RightEmptySleeve = False
                RightMeasurements(X) = cmtoinches(Righttxtinchflag.Text, RightLength(w))
            End If
        Next w

        Dim inchbit As Double
        For i = 0 To 17
            If Val(RightLength(i).Text) > 0 Then
                inchbit = cmtoinches(Righttxtinchflag.Text, RightLength(i))
                If inchbit = -1 Then
                    RightLength(i).Focus()
                    Exit Sub
                End If
                RightInchText(i).Text = ARMEDDIA1.fnInchesToText(inchbit)
            End If
        Next
        Dim pwinch As Double
        If Val(RighttxtPalmWrist.Text) > 0 Then
            pwinch = cmtoinches(Righttxtinchflag.Text, RighttxtPalmWrist)
            RightPWDist.Text = ARMEDDIA1.fnInchesToText(pwinch)
        End If

        If Val(RighttxtCircum.Text) > 0 Then
            pwinch = cmtoinches(Righttxtinchflag.Text, RighttxtCircum)
            RightTCirc.Text = ARMEDDIA1.fnInchesToText(pwinch)
        End If

        If Val(RighttxtThumb.Text) > 0 Then
            pwinch = cmtoinches(Righttxtinchflag.Text, RighttxtThumb)
            RightTLen.Text = ARMEDDIA1.fnInchesToText(pwinch)
        End If

        If Val(RighttxtGauntletExtension.Text) > 0 Then
            RightlabGauntletExtension.Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtGauntletExtension))
        End If

        PR_GetRightStartEndTapes()
        For Index = 0 To 17
            If RightLength(Index).Text <> "" And RightcboFabric.Text <> "" And RightcboPressure.Text <> "" And Rightmms(Index).Text <> "" Then
                w = Index
                X = Index + 1
                PR_GetRightFabric()
                PR_RightWristRed(w, X, g_sRightModulus)
                PR_RightPalmRed(w, X, g_sRightModulus)
                PR_RightLastTapeRed(w, X, g_sRightModulus)
                ''--------------PR_SetRresults()
                ''-----------------PR_SetGlobals()
            End If
        Next Index
        ''-------------------------------------
        'Set up first and last tapes depending on ID
        'nValue = Val(Mid(txtID.Text, 1, 2))
        'If nValue > 0 Then
        '    cboDistalTape.SelectedIndex = nValue
        'Else
        '    cboDistalTape.SelectedIndex = 0
        'End If
        If RightcboDistalTape.SelectedIndex < 0 Then
            RightcboDistalTape.SelectedIndex = 0
        End If
        RightcboDistalTape_SelectedIndexChanged(RightcboDistalTape, New System.EventArgs())

        'nValue = Val(Mid(txtID.Text, 3, 2))
        'If nValue > 0 Then
        '    cboProximalTape.SelectedIndex = (18 - nValue) + 1
        'ElseIf Mid(txtID.Text, 3, 2) <> "GT" Then  'IE Not for Detachable Gauntlet
        '    cboProximalTape.SelectedIndex = 0
        'End If
        If RightcboProximalTape.SelectedIndex < 0 Then
            RightcboProximalTape.SelectedIndex = 0
        End If
        RightcboProximalTape_SelectedIndexChanged(RightcboProximalTape, New System.EventArgs())
        ''---------------------------------------------
        Show()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub PR_RightLastTapeRed(ByRef w As Object, ByRef X As Object, ByRef Modulus As Object)
        Dim Backgram As Object
        If RightLength(w).Text <> "" Then
            'check last tape reduction
            If (w > 1) And (w = Val(strLastTape) - 1) And (RightcboFlaps.Text = "") Then
                If VB.Left(RightcboFabric.Text, 3) = "Pow" Then
                    Backgram = RightPow(Modulus, 1)
                    If Val(Rightgms(w).Text) <> Backgram Then
                        Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                        Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                    End If
                Else
                    Backgram = RightBob(Modulus, 1)
                    If Val(Rightgms(w).Text) <> Backgram Then
                        Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                        Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                    End If
                End If
                Rightreds(w).Text = CStr(10)
            ElseIf (w > 1) And (w = Val(strLastTape) - 1) And ((RightcboFlaps.Text <> "") Or (VB.Left(RightcboFlaps.Text, 4) = "None") Or (VB.Left(RightcboFlaps.Text, 5) = "Style")) Then
                If VB.Left(RightcboFlaps.Text, 4) <> "Vest" And VB.Left(RightcboFlaps.Text, 4) <> "Body" Then
                    If VB.Left(RightcboFabric.Text, 3) = "Pow" Then
                        Backgram = RightPow(Modulus, 3) '12 reduction for Flaps
                        If Val(Rightgms(w).Text) <> Backgram Then
                            Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                            Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                        End If
                    Else
                        Backgram = RightBob(Modulus, 3) '12 reduction for Flaps
                        If Val(Rightgms(w).Text) <> Backgram Then
                            Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                            Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                        End If
                    End If
                    Rightreds(w).Text = CStr(12)
                Else
                    If VB.Left(RightcboFabric.Text, 3) = "Pow" Then
                        Backgram = RightPow(Modulus, 1) '10 reduction for Vest & Body
                        If Val(Rightgms(w).Text) <> Backgram Then
                            Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                            Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                        End If
                    Else
                        Backgram = RightBob(Modulus, 1) '10 reduction for Vest & Body
                        If Val(Rightgms(w).Text) <> Backgram Then
                            Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
                            Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                        End If
                    End If
                    Rightreds(w).Text = CStr(10)
                End If
            End If
        End If
    End Sub
    Private Sub PR_RightPalmRed(ByRef w As Object, ByRef X As Object, ByRef Modulus As Object)
        Dim Backgram As Object
        If RightLength(w).Text <> "" Then
            'check palm reduction
            If (RightchkGauntlets.Checked = True And X = Val(RightcboPalm.SelectedIndex + 1)) Or (RightchkGauntlets.Checked = True And w = (CDbl(strFirstTape.Replace("st", "").Trim) - 1)) Then
                Rightreds(w).Text = CStr(15)
                If VB.Left(RightcboFabric.Text, 3) = "Pow" Then
                    Backgram = RightPow(Modulus, 6)
                Else
                    Backgram = RightBob(Modulus, 6)
                End If
                Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
                Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
            End If
        End If

    End Sub
    Private Sub PR_RightWristRed(ByRef w As Object, ByRef X As Object, ByRef Modulus As Object)
        Dim Backgram, Reduction As Object
        If RightLength(w).Text <> "" And X = Val(strFirstTape) And RightchkStump.Checked <> True And RightchkGauntlets.Checked <> True Then
            ' check wrist reduction
            If (w < 9) And Rightreds(w).Text = "" Then
                Rightreds(w).Text = CStr(14)
            ElseIf CDbl(Rightreds(w).Text) > 14 And (w < 9) Then
                Rightreds(w).Text = CStr(14)
                'Set elbow reduction to 5
            ElseIf CDbl(Rightreds(w).Text) <> 5 And (w = 10) Then
                Rightreds(w).Text = CStr(5)
                'Above elbow is also a 5
            ElseIf CDbl(Rightreds(w).Text) <> 5 And (w > 10) Then
                Rightreds(w).Text = CStr(5)
            ElseIf (w < 9) And ((CDbl(Rightreds(w).Text) < 10) And (CDbl(Rightreds(w).Text) > 0)) Then
                Rightreds(w).Text = CStr(10)
            End If

            'Set up an index into the reduction to grams mapping table
            'minimum is 1
            Reduction = CDbl(Rightreds(w).Text) - 9
            If Reduction < 1 Then Reduction = 1
            'From either powernet or bobinnet table get the grams and back calculate the
            'Rightmms
            If VB.Left(RightcboFabric.Text, 3) = "Pow" Then
                Backgram = RightPow(Modulus, Reduction)
            Else
                Backgram = RightBob(Modulus, Reduction)
            End If
            Rightmms(w).Text = CStr(ARMDIA1.round(Backgram / RightMeasurements(X)))
            Rightgms(w).Text = CStr(ARMDIA1.round(Backgram))
        End If
    End Sub
    Private Sub PR_GetRightFabric()
        Dim Fabtype, fabnumber, fabnam As Object
        Dim Modul, fabcount As Object
        fabnam = VB.Left(RightcboFabric.Text, 3)
        Fabtype = Mid(RightcboFabric.Text, 5, 3)
        If Val(Fabtype) > 340 Or Val(Fabtype) < 160 Then
            MsgBox("Fabric Modulus Not Recognised", 48, "Arm Details")
            Exit Sub
        End If
        If (fabnam <> "Pow") And (fabnam <> "Bob") Then
            MsgBox("Fabric Type Not Recognised, Should be either Bobbinette or Powernet", 48, "Arm Details")
            Exit Sub
        End If
        ''--------------------------g_sFabnam = RightcboFabric.Text
        ''--------------------------txtFabric.Text = RightcboFabric.Text
        fabnumber = 160
        fabcount = 1
        Do While fabcount < 20
            If Val(Fabtype) = fabnumber Then
                Modul = fabcount
                g_sRightModulus = Modul
                ''--------------------------txtModulus.Text = Modul
            End If
            fabcount = fabcount + 1
            fabnumber = fabnumber + 10
        Loop
    End Sub
    Private Sub PR_GetRightStartEndTapes()
        Dim g_iSecondLastTape As Object = ""
        Dim g_iSecondTape As Object = ""
        'Now sets start and last tape text values w.r.t style start and end
        Dim i, ii As Object
        'first  & second tapes
        'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        i = 0
        strFirstTape = ""
        'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While i < 18
            If RightLength(i).Text <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ARMDIA1.g_iRightFirstTape = i + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object g_iSecondTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                g_iSecondTape = i + 2
                Exit Do
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            i = i + 1
        Loop

        If ARMDIA1.g_iRightStyleFirstTape > ARMDIA1.g_iRightFirstTape And ARMDIA1.g_iRightStyleFirstTape >= 0 And cboDistalTape.SelectedIndex <> 0 Then
            'Use style tape value if it is given ( g_iRightStyleFirstTape = -1 => not given)
            strFirstTape = CStr(ARMDIA1.g_iRightStyleFirstTape)
            strSecondTape = CStr(ARMDIA1.g_iRightStyleFirstTape + 1)
        Else
            strFirstTape = CStr(ARMDIA1.g_iRightFirstTape)
            'UPGRADE_WARNING: Couldn't resolve default property of object g_iSecondTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strSecondTape = g_iSecondTape
        End If

        'secondlast & last tapes
        strLastTape = ""
        'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ii = 17
        'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Do While ii > 0
            If RightLength(ii).Text <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ARMDIA1.g_iRightLastTape = ii + 1
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object g_iSecondLastTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                g_iSecondLastTape = ii
                Exit Do
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ii = ii - 1
        Loop

        If ARMDIA1.g_iRightStyleLastTape < ARMDIA1.g_iRightLastTape And ARMDIA1.g_iRightStyleLastTape >= 0 And cboProximalTape.SelectedIndex <> 0 Then
            'Use style tape value if it is given ( g_iRightStyleLastTape = -1 => not given)
            strLastTape = CStr(ARMDIA1.g_iRightStyleLastTape)
            strSecondLastTape = CStr(ARMDIA1.g_iRightStyleLastTape - 1)
        Else
            strLastTape = CStr(ARMDIA1.g_iRightLastTape)
            'UPGRADE_WARNING: Couldn't resolve default property of object g_iSecondLastTape. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strSecondLastTape = g_iSecondLastTape
        End If

    End Sub

    Private Sub RightchkNone_CheckedChanged(eventSender As System.Object, e As EventArgs) Handles RightchkNone.CheckedChanged
        If eventSender.Checked Then
            'Disable gauntlet
            g_sRightHoleCheck = "False"
            RightfrmGauntlet.Enabled = False
            RightlblPalm.Enabled = False
            RightlblWrist.Enabled = False
            RightlblThumb.Enabled = False
            RightcboPalm.SelectedIndex = -1
            RightcboWrist.SelectedIndex = -1
            RightcboPalm.Enabled = False
            RightcboWrist.Enabled = False
            RightlblGauntletExtension.Enabled = False
            RighttxtGauntletExtension.Enabled = False


            If RightchkDetachable.CheckState = 1 Then
                'Reset the proximal tape to last
                RightchkDetachable.CheckState = System.Windows.Forms.CheckState.Unchecked
                RightchkDetachable_CheckStateChanged(RightchkDetachable, New System.EventArgs())
            End If

            RightchkNoThumb.CheckState = System.Windows.Forms.CheckState.Unchecked
            RightchkEnclosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            PWDist.Text = ""
            TCirc.Text = ""
            TLen.Text = ""
            RighttxtThumb.Text = ""
            RightlabGauntletExtension.Text = ""
            RighttxtGauntletExtension.Text = ""
            RighttxtThumb.Enabled = False
            RightchkEnclosed.Enabled = False
            RightchkDetachable.Enabled = False
            RightchkNoThumb.Enabled = False
            RightlblPalmWrist.Enabled = False
            RighttxtPalmWrist.Text = ""
            RighttxtPalmWrist.Enabled = False
            RighttxtCircum.Text = ""
            RighttxtCircum.Enabled = False
            RightlblCircum.Enabled = False

            ''------------------RighttxtGauntletFlag.Text = "False"

            'Enable distal tape position
            RightcboDistalTape.Enabled = True

            'Disable stump
            RightchkStump.Checked = False
            ''-------------RighttxtStump.Text = "False"

            Dim w, X As Object

            If Val(strFirstTape) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                X = Val(strFirstTape)
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                w = X - 1
                If RightLength(w).Text <> "" And RightcboFabric.Text <> "" And RightcboPressure.Text <> "" And Rightgms(w).Text <> "" And Rightreds(w).Text <> "" Then
                    PR_RightWristRed(w, X, g_sRightModulus)
                    PR_RightReductions(w, X)
                    ''---------------PR_SetRresults()
                    ''-----------------PR_SetGlobals()
                End If
            End If
            If Me.Visible = True Then ARMDIA1.g_Modified = True
        End If
    End Sub

    Private Sub RightchkStump_CheckedChanged(eventSender As System.Object, e As System.EventArgs) Handles RightchkStump.CheckedChanged
        If eventSender.Checked Then
            Dim entries, i As Object
            If RightchkStump.Checked = True Then
                'Disable gauntlet
                g_sRightHoleCheck = "False"
                RightfrmGauntlet.Enabled = False
                RightlblPalm.Enabled = False
                RightlblWrist.Enabled = False
                RightlblThumb.Enabled = False
                RightcboPalm.SelectedIndex = -1
                RightcboPalm.Enabled = False
                RightcboWrist.SelectedIndex = -1
                RightcboWrist.Enabled = False
                If RightchkDetachable.CheckState = 1 Then
                    'Reset the proximal tape to last
                    RightchkDetachable.CheckState = System.Windows.Forms.CheckState.Unchecked
                    RightchkDetachable_CheckStateChanged(RightchkDetachable, New System.EventArgs())
                End If
                RightchkNoThumb.CheckState = System.Windows.Forms.CheckState.Unchecked
                RightchkEnclosed.CheckState = System.Windows.Forms.CheckState.Unchecked
                RighttxtGauntletExtension.Enabled = False
                RightlblGauntletExtension.Enabled = False
                RightlabGauntletExtension.Text = ""
                RighttxtGauntletExtension.Text = ""
                RightPWDist.Text = ""
                RightTCirc.Text = ""
                RightTLen.Text = ""
                RighttxtThumb.Text = ""
                RighttxtThumb.Enabled = False
                RightchkEnclosed.Enabled = False
                RightchkDetachable.Enabled = False
                RightchkNoThumb.Enabled = False
                RightlblPalmWrist.Enabled = False
                RighttxtPalmWrist.Text = ""
                RighttxtPalmWrist.Enabled = False
                RighttxtCircum.Text = ""
                RighttxtCircum.Enabled = False
                RightlblCircum.Enabled = False
                ''----------------------------RighttxtGauntletFlag.Text = "False"

                'Set distal tape position to first
                RightcboDistalTape.Enabled = True
                RightcboDistalTape.SelectedIndex = 0
                RightcboDistalTape_SelectedIndexChanged(RightcboDistalTape, New System.EventArgs())
                'cboDistalTape.Enabled = False

            ElseIf RightchkStump.Checked = 0 Then
                RightchkGauntlets.Enabled = True
            End If

            Dim w, X As Object
            If Val(strFirstTape) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                X = Val(strFirstTape)
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                w = X - 1
                If RightLength(w).Text <> "" And RightcboFabric.Text <> "" And RightcboPressure.Text <> "" And Rightgms(w).Text <> "" And Rightreds(w).Text <> "" Then
                    If RightchkStump.Checked = True Then
                        PR_RightReductions(w, X)
                        PR_RightStump(w, X)
                        ''--------------PR_SetRresults()
                        ''---------------------PR_SetGlobals()
                    End If
                End If
            End If

            If Me.Visible = True Then ARMDIA1.g_Modified = True
        End If
    End Sub

    Private Sub RightchkGauntlets_CheckedChanged(eventsender As System.Object, e As EventArgs) Handles RightchkGauntlets.CheckedChanged
        If eventsender.Checked Then
            If RightchkGauntlets.Checked = True Then
                RightfrmGauntlet.Enabled = True
                RightchkStump.Checked = False
                RightlblPalm.Enabled = True
                RightlblWrist.Enabled = True
                RightlblThumb.Enabled = True
                RightcboPalm.Enabled = True
                RightcboWrist.Enabled = True
                RighttxtThumb.Enabled = True
                RightchkEnclosed.Enabled = True
                RightchkDetachable.Enabled = True
                RightchkNoThumb.Enabled = True
                RightlblPalmWrist.Enabled = True
                RighttxtPalmWrist.Enabled = True
                RighttxtCircum.Enabled = True
                RightlblCircum.Enabled = True
                RightlblGauntletExtension.Enabled = True
                RighttxtGauntletExtension.Enabled = True

                'Set distal tape position to first
                RightcboDistalTape.Enabled = True
                RightcboDistalTape.SelectedIndex = 0
                '        cboDistalTape_Click
                RightcboDistalTape.Enabled = False

            ElseIf RightchkGauntlets.Checked = False Then
                RightfrmGauntlet.Enabled = False
                RightchkNoThumb.Enabled = False
                RightlblPalmWrist.Enabled = False
                RighttxtPalmWrist.Enabled = False
                RightlblPalm.Enabled = False
                RightlblWrist.Enabled = False
                RightlblThumb.Enabled = False
                RightcboPalm.Enabled = False
                RightcboWrist.Enabled = False
                RighttxtThumb.Enabled = False
                RightchkEnclosed.Enabled = False
                RightchkDetachable.Enabled = False
                RighttxtCircum.Enabled = False
                RightlblCircum.Enabled = False
                RightlblGauntletExtension.Enabled = True
                RighttxtGauntletExtension.Enabled = True

            End If

            If Me.Visible = True Then ARMDIA1.g_Modified = True

        End If
    End Sub

    Private Sub RightoptProximalTape_CheckedChanged(eventSender As System.Object, e As EventArgs) Handles RightoptProximalTape.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = RightoptProximalTape.GetIndex(eventSender)
            'only two choices
            '
            If Index = 0 Then
                'Disable flaps
                RightfrmFlaps.Enabled = False
                RightcboFlaps.SelectedIndex = -1
                ''--------------------RightcboFlaps.Text = "none"
                RightcboFlaps_SelectedIndexChanged(RightcboFlaps, New System.EventArgs()) 'Force change of pressure
                RightcboFlaps.Enabled = False

                RightlblStyle.Enabled = False
                RighttxtStrapLength.Text = ""
                RighttxtStrapLength.Enabled = False
                RightlabStrap.Text = ""
                RightlblStrapLength.Enabled = False

                RighttxtFrontStrapLength.Text = ""
                RighttxtFrontStrapLength.Enabled = False
                RightlabFrontStrapLength.Text = ""
                RightlblFrontStrapLength.Enabled = False

                RighttxtCustFlapLength.Text = ""
                RighttxtCustFlapLength.Enabled = False
                RightlabCustFlapLength.Text = ""
                RightlblCustFlapLength.Enabled = False

                RighttxtWaistCir.Text = ""
                RighttxtWaistCir.Enabled = False
                RightlabWaistCir.Text = ""
                RightlblWaistCir.Enabled = False

                'Enable proximal tape position
                RightcboProximalTape.Enabled = True

            Else
                'Reset proximal tape position to last
                RightcboProximalTape.SelectedIndex = 0
                RightcboProximalTape_SelectedIndexChanged(RightcboProximalTape, New System.EventArgs())
                RightcboProximalTape.Enabled = False

                'Enable flaps
                RightfrmFlaps.Enabled = True
                RightcboFlaps.Enabled = True
                RightlblStyle.Enabled = True

                RighttxtStrapLength.Enabled = True
                RightlblStrapLength.Enabled = True
                RightlabStrap.Enabled = True

                RighttxtCustFlapLength.Enabled = True
                RightlblCustFlapLength.Enabled = True
                RightlabCustFlapLength.Enabled = True

                RighttxtFrontStrapLength.Enabled = True
                RightlblFrontStrapLength.Enabled = True
                RightlabFrontStrapLength.Enabled = True

                'Can't have a detachable gauntlet with a flap
                If RightchkDetachable.CheckState = 1 Then RightchkDetachable.CheckState = System.Windows.Forms.CheckState.Unchecked

                '        txtWaistCir.Enabled = True
                '        lblWaistCir.Enabled = True
            End If

            If Me.Visible = True Then ARMDIA1.g_Modified = True

        End If
    End Sub

    Private Sub RightcboFlaps_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RightcboFlaps.SelectedIndexChanged
        Dim w, X As Object
        If RightcboFlaps.Text = "none" Then
            RightcboFlaps.Text = ""
            RightlabStrap.Text = ""
            RighttxtCustFlapLength.Text = ""
            RightlabCustFlapLength.Text = ""
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X = Val(strLastTape)
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        w = X - 1
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If X > 1 Then
            If RightLength(w).Text <> "" And RightcboFabric.Text <> "" And RightcboPressure.Text <> "" And Rightgms(w).Text <> "" And Rightreds(w).Text <> "" Then
                PR_RightLastTapeRed(w, X, g_sRightModulus)
                ''-----------------PR_CheckFlaps(w)
                ''------------------PR_SetRresults()
                ''------------------PR_SetGlobals()
            End If
        End If

        'Enable waist circumference for style "D"
        If InStr(1, RightcboFlaps.Text, "D", 0) > 0 Then
            RightlabWaistCir.Enabled = True 'Inches convesion
            RighttxtWaistCir.Enabled = True 'Edit box
            RightlblWaistCir.Enabled = True
        ElseIf RighttxtWaistCir.Enabled = True Then
            'Disable only if previously enabled
            RighttxtWaistCir.Text = "" 'Delete display of value
            RightlabWaistCir.Text = "" 'Delete inches display of value
            RightlabWaistCir.Enabled = False
            RighttxtWaistCir.Enabled = False
            RightlblWaistCir.Enabled = False
        End If

        'Disable all boxes if Vest Raglan chosen
        If VB.Left(RightcboFlaps.Text, 4) = "Vest" Or VB.Left(RightcboFlaps.Text, 4) = "Body" Then
            RighttxtStrapLength.Text = ""
            RighttxtStrapLength.Enabled = False
            RighttxtFrontStrapLength.Text = ""
            RighttxtFrontStrapLength.Enabled = False
            RightlblFrontStrapLength.Enabled = False
            RightlabFrontStrapLength.Text = ""
            RightlabStrap.Text = ""
            RightlblStrapLength.Enabled = False
            RighttxtCustFlapLength.Text = ""
            RighttxtCustFlapLength.Enabled = False
            RightlabCustFlapLength.Text = ""
            RightlblCustFlapLength.Enabled = False
        Else
            'Enable if not enbaled only
            If RighttxtStrapLength.Enabled = False Then
                RighttxtStrapLength.Enabled = True
                RightlblStrapLength.Enabled = True
            End If
            If RighttxtCustFlapLength.Enabled = False Then
                RighttxtCustFlapLength.Enabled = True
                RightlblCustFlapLength.Enabled = True
            End If
            If RighttxtFrontStrapLength.Enabled = False Then
                RighttxtFrontStrapLength.Enabled = True
                RightlblFrontStrapLength.Enabled = True
            End If
        End If
        If Me.Visible = True Then ARMDIA1.g_Modified = True
    End Sub

    Private Sub RighttxtStrapLength_Leave(sender As Object, e As EventArgs) Handles RighttxtStrapLength.Leave
        If Not FN_CheckValue(RighttxtStrapLength, "Strap") Then
            Exit Sub
        ElseIf Val(RighttxtStrapLength.Text) > 0 Then
            RightlabStrap.Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtStrapLength))
        Else
            RightlabStrap.Text = ""
        End If
        Dim w, X As Object
        X = Val(txtLastTape.Text)
        w = X - 1
        If X > 0 Then
            If RightLength(w).Text <> "" And RightcboFabric.Text <> "" And RightcboPressure.Text <> "" And Rightgms(w).Text <> "" And Rightreds(w).Text <> "" Then
                PR_RightLastTapeRed(w, X, g_sRightModulus)
                ''-------------------------------
                'PR_CheckFlaps(w)

                'PR_SetRresults()

                'PR_SetGlobals()
            End If
        End If
    End Sub

    Private Sub RighttxtFrontStrapLength_Leave(sender As Object, e As EventArgs) Handles RighttxtFrontStrapLength.Leave
        If Not FN_CheckValue(RighttxtFrontStrapLength, "Front Strap") Then
            Exit Sub
        ElseIf Val(RighttxtFrontStrapLength.Text) > 0 Then
            RightlabFrontStrapLength.Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtFrontStrapLength))
        Else
            RightlabFrontStrapLength.Text = ""
        End If
    End Sub

    Private Sub RighttxtCustFlapLength_Leave(sender As Object, e As EventArgs) Handles RighttxtCustFlapLength.Leave
        If Not FN_CheckValue(RighttxtCustFlapLength, "Custom Flap") Then
            Exit Sub
        ElseIf Val(RighttxtCustFlapLength.Text) > 0 Then
            RightlabCustFlapLength.Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtCustFlapLength))
        Else
            RightlabCustFlapLength.Text = ""
        End If
    End Sub

    Private Sub RighttxtWaistCir_Leave(sender As Object, e As EventArgs) Handles RighttxtWaistCir.Leave
        If Not FN_CheckValue(RighttxtWaistCir, "Style-D waist circumference") Then
            Exit Sub
        ElseIf Val(RighttxtWaistCir.Text) > 0 Then
            RightlabWaistCir.Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtWaistCir))
        Else
            RightlabWaistCir.Text = ""
        End If
    End Sub

    Private Sub RighttxtPalmWrist_Leave(sender As Object, e As EventArgs) Handles RighttxtPalmWrist.Leave
        If Not FN_CheckValue(RighttxtPalmWrist, "Palm to Wrist") Then
            Exit Sub
        ElseIf Val(RighttxtPalmWrist.Text) > 0 Then
            RightPWDist.Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtPalmWrist))
        Else
            RightPWDist.Text = ""
        End If
    End Sub

    Private Sub RighttxtCircum_Leave(sender As Object, e As EventArgs) Handles RighttxtCircum.Leave
        If Not FN_CheckValue(RighttxtCircum, "Thumb Circumference") Then
            Exit Sub
        ElseIf Val(RighttxtCircum.Text) > 0 Then
            RightTCirc.Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtCircum))
        Else
            RightTCirc.Text = ""
        End If
    End Sub

    Private Sub RighttxtThumb_Leave(sender As Object, e As EventArgs) Handles RighttxtThumb.Leave
        If Not FN_CheckValue(RighttxtThumb, "Thumb") Then
            Exit Sub
        ElseIf Val(RighttxtThumb.Text) > 0 Then
            RightTLen.Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtThumb))
        Else
            RightTLen.Text = ""
        End If
    End Sub

    Private Sub RighttxtGauntletExtension_Leave(sender As Object, e As EventArgs) Handles RighttxtGauntletExtension.Leave
        If Not FN_CheckValue(RighttxtGauntletExtension, "Gauntlet Extension") Then
            Exit Sub
        ElseIf Val(RighttxtGauntletExtension.Text) > 0 Then
            RightlabGauntletExtension.Text = ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtGauntletExtension))
        Else
            RightlabGauntletExtension.Text = ""
        End If
    End Sub

    Private Sub RightchkDetachable_CheckStateChanged(sender As Object, e As EventArgs) Handles RightchkDetachable.CheckStateChanged
        'Setup for Detachable gauntlet
        If RightchkDetachable.CheckState = 1 Then
            'For detachable gauntlets
            'Note:- RightcboProximalTape gives the tapes in reverse order
            RightcboProximalTape.SelectedIndex = 13 '+1-1/2 tape

            'If Val(txtAge) > 14 Then
            '    RightcboProximalTape.ListIndex = 12 '+3 tape for adults
            'Else
            '    RightcboProximalTape.ListIndex = 13 '+1-1/2 tape for children
            'End If
            RightcboProximalTape_SelectedIndexChanged(RightcboProximalTape, New System.EventArgs())
            RightoptProximalTape(0).Checked = 1
            RightoptProximalTape(1).Checked = 0
        ElseIf RightchkDetachable.CheckState = 0 Then
            'For normal gauntlets us last tape
            RightcboProximalTape.SelectedIndex = 0
            RightcboProximalTape_SelectedIndexChanged(RightcboProximalTape, New System.EventArgs())
        End If
    End Sub

    Private Sub RightchkNoThumb_CheckStateChanged(sender As Object, e As EventArgs) Handles RightchkNoThumb.CheckStateChanged
        If RightchkNoThumb.CheckState = 1 Then
            RightchkEnclosed.Enabled = False
            RightchkEnclosed.CheckState = System.Windows.Forms.CheckState.Unchecked
            RighttxtCircum.Enabled = False
            RighttxtThumb.Enabled = False
            RightlblCircum.Enabled = False
            RightlblThumb.Enabled = False
            RighttxtCircum.Text = ""
            RightTCirc.Text = ""
            RighttxtThumb.Text = ""
            RightTLen.Text = ""

        ElseIf RightchkNoThumb.CheckState = 0 And RightchkGauntlets.Checked = True Then
            RighttxtCircum.Enabled = True
            RighttxtThumb.Enabled = True
            RightlblCircum.Enabled = True
            RightlblThumb.Enabled = True
            RightchkEnclosed.Enabled = True
        End If

        If Me.Visible = True Then ARMDIA1.g_Modified = True
    End Sub

    Private Sub RightchkEnclosed_CheckStateChanged(sender As Object, e As EventArgs) Handles RightchkEnclosed.CheckStateChanged
        If RightchkEnclosed.CheckState = 1 Then
            RightchkNoThumb.Enabled = False
            RightchkNoThumb.CheckState = False
            RighttxtCircum.Enabled = True
            RighttxtThumb.Enabled = True
            RightlblCircum.Enabled = True
            RightlblThumb.Enabled = True
        ElseIf RightchkEnclosed.CheckState = 0 Then
            RightchkNoThumb.Enabled = True
        End If

        If Me.Visible = True Then ARMDIA1.g_Modified = True
    End Sub

    Private Sub RightcboPressure_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RightcboPressure.SelectedIndexChanged
        Dim Index As Short

        If Me.Visible = True Then ARMDIA1.g_Modified = True
    End Sub

    Private Sub RightcboFabric_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RightcboFabric.SelectedIndexChanged
        Dim fabnam, Fabtype As Object
        ' check for fabric modulus
        If RightcboFabric.Text <> "" Then
            fabnam = VB.Left(RightcboFabric.Text, 3)
            Fabtype = Mid(RightcboFabric.Text, 5, 3)
            If Val(Fabtype) > 340 Or Val(Fabtype) < 160 Then
                MsgBox("Fabric Modulus Not Recognised", 48, "Arm Details")
                RightcboFabric.Text = ""
                Exit Sub
            End If
            If (fabnam <> "Pow") And (fabnam <> "Bob") Then
                MsgBox("Fabric Type Not Recognised, Should be either Bobbinette or Powernet", 48, "Arm Details")
                RightcboFabric.Text = ""
                Exit Sub
            End If
            If Me.Visible = True Then ARMDIA1.g_Modified = True
        End If
    End Sub

    Private Sub RightcboDistalTape_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RightcboDistalTape.SelectedIndexChanged
        'Set style first tape
        'Modify tape labels to show this (lblTape index starts at 0)
        'NOTE:-
        '    cboDistalTape.ListIndex = 0 = 1st
        '    cboDistalTape.ListIndex = 1 = -6
        '    cboDistalTape.ListIndex = 2 = -4½
        '    cboDistalTape.ListIndex = 3 = -3
        '                .
        '                .
        '                .
        '    cboDistalTape.ListIndex = 18 = 19½
        '
        Dim iTape As Short

        'Reset display of previous to Black
        If ARMDIA1.g_iRightStyleFirstTape > 0 Then
            RightlblTape(ARMDIA1.g_iRightStyleFirstTape - 1).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 0, 0))
        End If

        'set style first tape value
        iTape = RightcboDistalTape.SelectedIndex - 1
        If iTape < 0 Then
            'Use last tape
            If ARMDIA1.g_iRightFirstTape > 0 Then
                ARMDIA1.g_iRightStyleFirstTape = ARMDIA1.g_iRightFirstTape
            Else
                ARMDIA1.g_iRightStyleFirstTape = -1
            End If
        Else
            'use tape pointed at by index
            ARMDIA1.g_iRightStyleFirstTape = RightcboDistalTape.SelectedIndex
        End If

        'Set colour of tape label to green
        iTape = ARMDIA1.g_iRightStyleFirstTape - 1 'lblTape Index starts at 0
        If iTape >= 0 Then
            RightlblTape(iTape).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 255, 0))
        End If

        If Me.Visible = True Then ARMDIA1.g_Modified = True
    End Sub

    Private Sub RightcboWrist_Leave(sender As Object, e As EventArgs) Handles RightcboWrist.Leave
        If RightcboPalm.SelectedIndex > -1 Then
            If RightcboPalm.SelectedIndex >= RightcboWrist.SelectedIndex Then
                MsgBox("Wrist No. must be greater than Palm No.", 48, "Gauntlet Details")
                RightcboPalm.Focus()
                Exit Sub
            End If
        End If

        If Me.Visible = True Then ARMDIA1.g_Modified = True
    End Sub

    Private Sub RightcboProximalTape_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RightcboProximalTape.SelectedIndexChanged
        'Set style last tape
        'Modify tape labels to show this (lblTape index starts at 0)
        'NOTE:-
        '    cboProximal.ListIndex = 0 = Last
        '    cboProximal.ListIndex = 1 = 19½
        '    cboProximal.ListIndex = 2 = 18
        '    cboProximal.ListIndex = 3 = 16½
        '                .
        '                .
        '                .
        '    cboProximal.ListIndex = 18 = -6
        '
        Dim iTape As Short

        'Reset display of previous to Black
        If ARMDIA1.g_iRightStyleLastTape > 0 Then
            RightlblTape(ARMDIA1.g_iRightStyleLastTape - 1).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(0, 0, 0))
        End If

        'set style last tape value
        iTape = RightcboProximalTape.SelectedIndex - 1
        If iTape < 0 Then
            'Use last tape
            If ARMDIA1.g_iRightLastTape > 0 Then
                ARMDIA1.g_iRightStyleLastTape = ARMDIA1.g_iRightLastTape
            Else
                ARMDIA1.g_iRightStyleLastTape = -1
            End If
        Else
            'use tape pointed at by index
            ARMDIA1.g_iRightStyleLastTape = (18 - RightcboProximalTape.SelectedIndex) + 1
        End If

        'Set colour of tape label to red
        iTape = ARMDIA1.g_iRightStyleLastTape - 1 'lblTape Index starts at 0
        If iTape >= 0 Then
            RightlblTape(iTape).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(255, 0, 0))
        End If


        If Me.Visible = True Then ARMDIA1.g_Modified = True
    End Sub
    Private Sub PR_SetRightDistances(ByRef w As Object, ByRef X As Object)
        'Check Distances between tapes
        Dim tdist, tpace As Object
        tpace = 1.375
        If RightLength(w).Text <> "" Then
            'First tape
            If X = Val(strFirstTape) Then
                Righttapewidths(X) = "0"
                'secondTape
            ElseIf X = Val(strSecondTape) And RightchkGauntlets.Checked = False Then
                If Val(RighttxtWristPleat1.Text) = 0 Or Not RighttxtWristPleat1.Enabled Then
                    Righttapewidths(X) = tpace
                ElseIf Val(RighttxtWristPleat1.Text) > 0 Then
                    tdist = Val(RighttxtWristPleat1.Text)
                    tdist = FN_PleatDistances(Righttxtinchflag, tdist)
                    Righttapewidths(X) = Righttapewidths(X - 1) + tdist
                End If

            ElseIf X = Val(strSecondTape) And RightchkGauntlets.Checked = True Then
                If Val(RightLength(w).Text) = 0 Then
                    Righttapewidths(X) = ""
                ElseIf Val(RightLength(w).Text) > 0 And Val(strWristNo) = X Then
                    ''----------------------Righttapewidths(X) = Val(txtPalmWristDist.Text)
                    Righttapewidths(X) = cmtoinches(Righttxtinchflag.Text, RighttxtPalmWrist)
                ElseIf Val(RightLength(w).Text) > 0 And Val(strWristNo) <> X Then
                    Righttapewidths(X) = tpace
                End If

                'thirdtape
            ElseIf X = Val(CStr(CDbl(strSecondTape) + 1)) And RightchkGauntlets.Checked = False Then
                If Val(RighttxtWristPleat2.Text) = 0 Or Not RighttxtWristPleat2.Enabled Then
                    Righttapewidths(X) = Righttapewidths(X - 1) + tpace
                ElseIf Val(RighttxtWristPleat2.Text) > 0 Then
                    tdist = Val(RighttxtWristPleat2.Text)
                    tdist = FN_PleatDistances(Righttxtinchflag, tdist)
                    Righttapewidths(X) = Val(Righttapewidths(X - 1)) + tdist
                End If

            ElseIf X = Val(CStr(CDbl(strSecondTape) + 1)) And RightchkGauntlets.Checked = True Then
                If (RightLength(w - 1).Text = "") Then
                    ''-----------Righttapewidths(X) = Val(Righttapewidths(X - 2)) + Val(txtPalmWristDist.Text)
                    Righttapewidths(X) = Val(Righttapewidths(X - 2)) + cmtoinches(Righttxtinchflag.Text, RighttxtPalmWrist)
                ElseIf RightLength(w - 1).Text <> "" And Val(strWristNo) = X Then
                    ''---------Righttapewidths(X) = Righttapewidths(X - 1) + Val(txtPalmWristDist.Text)
                    Righttapewidths(X) = Righttapewidths(X - 1) + cmtoinches(Righttxtinchflag.Text, RighttxtPalmWrist)
                ElseIf RightLength(w - 1).Text <> "" And Val(strWristNo) <> X Then
                    Righttapewidths(X) = Righttapewidths(X - 1) + tpace
                ElseIf RightLength(w - 1).Text <> "" Then
                    Righttapewidths(X) = ""
                End If

                'fourthtape
            ElseIf X = Val(CStr(CDbl(strSecondTape) + 2)) And RightchkGauntlets.Checked = False Then
                Righttapewidths(X) = Righttapewidths(X - 1) + tpace

            ElseIf X = Val(CStr(CDbl(strSecondTape) + 2)) And RightchkGauntlets.Checked = True Then
                If RightLength(w - 1).Text = "" And Val(strWristNo) = X Then
                    ''-------------------Righttapewidths(X) = Righttapewidths(X - 2) + Val(txtPalmWristDist.Text)
                    Righttapewidths(X) = Righttapewidths(X - 2) + cmtoinches(Righttxtinchflag.Text, RighttxtPalmWrist)
                ElseIf RightLength(w - 1).Text <> "" And Val(strWristNo) = X Then
                    ''-----------------Righttapewidths(X) = Righttapewidths(X - 1) + Val(txtPalmWristDist.Text)
                    Righttapewidths(X) = Righttapewidths(X - 1) + cmtoinches(Righttxtinchflag.Text, RighttxtPalmWrist)
                ElseIf RightLength(w - 1).Text <> "" And Val(strWristNo) <> X Then
                    Righttapewidths(X) = Righttapewidths(X - 1) + tpace
                End If

                'Tapes above fourthtape
            ElseIf X > (Val(strSecondTape) + 2) And X < Val(strSecondLastTape) Then
                Righttapewidths(X) = Righttapewidths(X - 1) + tpace

                'Lasttape -1
            ElseIf X = Val(strSecondLastTape) And (X > 1) Then
                If Val(RighttxtShoulderPleat2.Text) = 0 Or Not RighttxtShoulderPleat2.Enabled Then
                    Righttapewidths(X) = Val(Righttapewidths(X - 1)) + tpace
                ElseIf Val(RighttxtShoulderPleat2.Text) > 0 Then
                    tdist = Val(RighttxtShoulderPleat2.Text)
                    tdist = FN_PleatDistances(Righttxtinchflag, tdist)
                    Righttapewidths(X) = Val(Righttapewidths(X - 1)) + tdist
                End If

                'Last tape
            ElseIf X = Val(strLastTape) And (X > 1) Then
                If Val(RighttxtShoulderPleat1.Text) = 0 Or Not RighttxtShoulderPleat1.Enabled Then
                    Righttapewidths(X) = Val(Righttapewidths(X - 1)) + tpace
                ElseIf Val(RighttxtShoulderPleat1.Text) > 0 Then
                    tdist = Val(RighttxtShoulderPleat1.Text)
                    tdist = FN_PleatDistances(Righttxtinchflag, tdist)
                    Righttapewidths(X) = Val(Righttapewidths(X - 1)) + tdist
                End If
            End If

            'The following statements duplicate some of the above.  They are here
            'as bug fixes as I can't understand the above code.  The above does work
            '(mostly) so I have left well enough alone.

            'Account for wrist pleats if a gauntlet given
            If Val(strWristNo) + 1 = X And (Val(RighttxtWristPleat1.Text) > 0 And RighttxtWristPleat1.Enabled) And RightchkGauntlets.Checked = True Then
                Righttapewidths(X) = Val(Righttapewidths(X - 1)) + FN_PleatDistances(Righttxtinchflag, Val(RighttxtWristPleat1.Text))
            End If
            If Val(strWristNo) + 2 = X And (Val(RighttxtWristPleat2.Text) > 0 And RighttxtWristPleat2.Enabled) And RightchkGauntlets.Checked = True Then
                Righttapewidths(X) = Val(Righttapewidths(X - 1)) + FN_PleatDistances(Righttxtinchflag, Val(RighttxtWristPleat2.Text))
            End If

            'Account for 1ST shoulder pleat
            If Val(strLastTape) = X And (Val(RighttxtShoulderPleat1.Text) > 0 And RighttxtShoulderPleat1.Enabled) Then
                Righttapewidths(X) = Val(Righttapewidths(X - 1)) + FN_PleatDistances(Righttxtinchflag, Val(RighttxtShoulderPleat1.Text))
            End If
            'Account for 2nd shoulder pleat
            If Val(strSecondLastTape) = X And (Val(RighttxtShoulderPleat2.Text) > 0 And RighttxtShoulderPleat2.Enabled) Then
                Righttapewidths(X) = Val(Righttapewidths(X - 1)) + FN_PleatDistances(Righttxtinchflag, Val(RighttxtShoulderPleat2.Text))
            End If
        End If
    End Sub
    Private Sub PR_RightPressurecalc(ByRef Index As Short)
        Dim w, X As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        w = Index
        'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        X = Index + 1
        PR_RightInitaliseEmpty(w) 'Zero any empty fields
        PR_RightInitaliseArrays(w, X)
        If (RightLength(Index).Text <> "" And Rightmms(Index).Text = "") Or (RightLength(Index).Text <> "" And g_sRightPressureChange = "True") Then
            PR_SetRightMMs(w, g_sRightAmm, g_sRightBmm, g_sRightCmm, g_sRightDmm)
        End If

        PR_RightReductions(w, X)
        PR_RightWristRed(w, X, g_sRightModulus)
        PR_RightPalmRed(w, X, g_sRightModulus)
        PR_RightStump(w, X)
        PR_LastRightCircum(w, X)
        PR_RightLastTapeRed(w, X, g_sRightModulus)
        ''-----------------------PR_CheckFlaps(w)
        PR_RightDetachable(w, X, g_sRightModulus, RightcboFabric.Text)
    End Sub
    Private Sub RightCalculate_Click()
        Dim ii As Short
        '
        If RightcboPressure.Text = "" Then
            If Me.Visible = True Then MsgBox("Pressure has not been selected", 48, "Arm Details")
            Exit Sub
        End If
        If RightcboFabric.Text = "" Then
            If Me.Visible = True Then MsgBox("Fabric has not been selected", 48, "Arm Details")
            Exit Sub
        End If

        PR_CheckRightValues()

        If RightcboProximalTape.SelectedIndex = 0 Then RightcboProximalTape_SelectedIndexChanged(RightcboProximalTape, New System.EventArgs())
        If RightcboDistalTape.SelectedIndex = 0 Then RightcboDistalTape_SelectedIndexChanged(RightcboDistalTape, New System.EventArgs())
        If ARMDIA1.g_iRightStyleLastTape < ARMDIA1.g_iRightStyleFirstTape Then
            If Me.Visible = True Then MsgBox("Chosen distal tape is above Proximal tape!", 48, "Arm Details")
            Exit Sub
        End If

        If RightchkGauntlets.Checked = True Then
            If RightcboPalm.SelectedIndex < 0 Then
                If Me.Visible = True Then MsgBox("Palm tape position not given", 48, "Gauntlet Details")
                Exit Sub
            End If
            If RightcboWrist.SelectedIndex < 0 Then
                If Me.Visible = True Then MsgBox("Wrist tape position not given", 48, "Gauntlet Details")
                Exit Sub
            End If
            PR_GetRightGauntletDetails()
            PR_RightGauntletHoleCheck()
        End If

        g_sRightPressureChange = "True" 'Use this to force a recalculation
        PR_GetRightStartEndTapes()
        PR_GetRightPressure()
        PR_GetRightFabric()
        PR_GetRightGauntletDetails()
        For ii = 0 To 17
            If Val(RightLength(ii).Text) > 0 Then
                PR_RightPressurecalc(ii)
            End If
        Next ii
        ''---------------PR_SetRresults()
        ''-------------- PR_SetGlobals()
        g_sRightPressureChange = "False"

        'Clean up displayed values
        For ii = 0 To CDbl(strFirstTape) - 2
            If Rightmms(ii).Text = "" Or Rightmms(ii).Text = "0" Then
                Rightmms(ii).Text = ""
                Rightgms(ii).Text = ""
                Rightreds(ii).Text = ""
            End If
        Next ii

        For ii = 17 To CInt(strLastTape) Step -1
            If Rightmms(ii).Text = "" Or Rightmms(ii).Text = "0" Then
                Rightmms(ii).Text = ""
                Rightgms(ii).Text = ""
                Rightreds(ii).Text = ""
            End If
        Next ii

        If Me.Visible = True Then ARMDIA1.g_Modified = True
        'UPGRADE_WARNING: Couldn't resolve default property of object EmptySleeve. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        EmptySleeve = False

    End Sub

    Private Sub armdia_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If g_bIsClose = False Then
            Dim Response As Short
            Dim sLeftCurrentValue As String = FN_LeftArmValuesString()
            Dim sRightCurrentValue As String = FN_RightArmValuesString()
            Dim bIsChanged As Boolean = False
            If sLeftCurrentValue <> g_sLeftArmValues Then
                bIsChanged = True
            ElseIf sRightCurrentValue <> g_sRightArmValues Then
                bIsChanged = True
            End If
            ''If ARMDIA1.g_Modified = True Then
            If bIsChanged = True Then
                Response = MsgBox("Changes have been made, Save changes before closing", 35, "Arm Details")
                Select Case Response
                    Case IDYES
                        Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
                        ARMDIA1.fNum = ARMDIA1.FN_Open(sDrawFile, txtPatientName.Text, txtFileNo.Text, txtSleeve.Text, "Save")
                        'Check Flaps
                        txtFlap.Text = cboFlaps.Text
                        If txtFlap.Text = "" Then g_sFlap = "None" Else g_sFlap = txtFlap.Text
                        If txtFlap.Text <> "" And txtStrapLength.Text = "" Then
                            If txtinchflag.Text = "cm" Then
                                txtStrap.Text = "61"
                            Else
                                txtStrap.Text = "24"
                            End If
                        End If
                        PR_GetStartEndTapes()
                        PR_GetGauntletDetails()
                        PR_SetRresults()
                        PR_SetGlobals()
                        ARMDIA1.g_sFabric = cboFabric.Text
                        'Update
                        DataBaseDataUpDate("Save")
                        FileClose(ARMDIA1.fNum)
                        saveInfoToDWG()
                        saveRightInfoToDWG()
                    Case IDCANCEL
                        e.Cancel = True
                End Select
            End If
        End If
        g_bIsClose = False
    End Sub

    Private Sub TypeRightDetails(ByRef Profile As ARMDIA1.curve, ByRef numtapes As Object)
        Dim n, nTextHt As Object
        Dim xyDetails As ARMDIA1.XY
        Dim sDetails, NL As Object
        Dim sSymbol As String
        Dim nSpacing As Double

        nSpacing = 0.2
        ARMDIA1.PR_SetTextData(1, 32, -1, -1, -1) 'Horiz-Left, Vertical-Bottom

        If Profile.X(Profile.n) < 6 Then
            ARMDIA1.PR_MakeXY(xyDetails, Profile.X(Profile.n) - 2, 2.5)
        Else
            ARMDIA1.PR_MakeXY(xyDetails, Profile.X(1) + 4, 2.5)
        End If
        If VB.Left(RightcboFlaps.Text, 4) = "Vest" Or VB.Left(RightcboFlaps.Text, 4) = "Body" Then
            ARMDIA1.PR_MakeXY(xyDetails, Profile.X(Profile.n) + 1.5, 1.75)
        End If

        'Main details in Green on layer notes
        ARMDIA1.PR_SetLayer("Notes")
        Dim sText As String

        'sText = txtSleeve.Text & "\n" & txtPatientName.Text & "\n" & ARMDIA1.g_sWorkOrder & "\n" & Trim(Mid(txtFabric.Text, 4))
        'ARMDIA1.PR_DrawText(sText, xyDetails, 0.125)
        sText = txtSleeve.Text & Chr(10) & txtPatientName.Text & Chr(10) & txtWorkOrder.Text & Chr(10) & Trim(Mid(RightcboFabric.Text, 4))
        ARMDIA1.PR_DrawMText(sText, xyDetails, False)

        'Other patient details in black on layer construct
        ARMDIA1.PR_SetLayer("Construct")

        ARMDIA1.PR_MakeXY(xyDetails, xyDetails.X, xyDetails.y - 0.8)
        'sText = txtFileNo.Text & "\n" & txtDiagnosis.Text & "\n" & txtAge.Text & "\n" & txtSex.Text
        'ARMDIA1.PR_DrawText(sText, xyDetails, 0.125)
        sText = txtFileNo.Text & Chr(10) & txtDiagnosis.Text & Chr(10) & txtAge.Text & Chr(10) & txtSex.Text
        ARMDIA1.PR_DrawMText(sText, xyDetails, False)
    End Sub
    Private Sub RightElastic(ByRef Profile As ARMDIA1.curve)
        Dim elastictop As New ARMDIA1.curve
        Dim xyEpt2, xyEpt1, xyEpt3 As New ARMDIA1.XY

        ARMDIA1.g_nCurrTextAngle = 90
        ARMDIA1.PR_SetTextData(2, 16, -1, -1, -1) 'Horiz center, Vertical center

        elastictop.n = 3
        Dim tempDataX(3) As Double
        Dim tempDataY(3) As Double
        tempDataX(1) = Profile.X(Profile.n)
        tempDataY(1) = Profile.y(Profile.n) / 2
        tempDataY(2) = Profile.y(Profile.n) / 2
        tempDataY(3) = Profile.y(Profile.n) + 0.375

        'Proximal Elastic and Cutback
        If Val(txtAge.Text) > 10 And Val(strLastTape) = 11 Then
            tempDataX(2) = Profile.X(Profile.n) - 0.75
            tempDataX(3) = Profile.X(Profile.n) - 0.75
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
            elastictop.X = tempDataX
            elastictop.y = tempDataY
            ''ARMDIA1.PR_DrawPoly(elastictop)
        ElseIf Val(txtAge.Text) <= 10 And Val(strLastTape) = 11 Then
            'Proximal Elastic @ Elbow
            tempDataX(2) = Profile.X(Profile.n) - 0.375
            tempDataX(3) = Profile.X(Profile.n) - 0.375
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
            elastictop.X = tempDataX
            elastictop.y = tempDataY
            ARMDIA1.PR_DrawPoly(elastictop)
            ARMDIA1.PR_MakeXY(xyEpt1, Profile.X(Profile.n) - 0.75, Profile.y(Profile.n) / 2)
            ARMDIA1.PR_SetLayer("Notes")
            ''----------ARMDIA1.PR_DrawText("1/2'' Elastic", xyEpt1, 0.125, 0)
            ARMDIA1.PR_DrawTextCenter("1/2'' Elastic", xyEpt1, 0.125, (90 * (ARMDIA1.PI / 180)))
        ElseIf RightcboFlaps.Text = "" And Val(txtAge.Text) <= 10 And Val(strLastTape) <> 11 Then
            'Proximal Elastic on plain arms for children under 10
            ARMDIA1.PR_MakeXY(xyEpt1, Profile.X(Profile.n) - 0.25, Profile.y(Profile.n) / 2)
            ARMDIA1.PR_SetLayer("Notes")
            ''----------ARMDIA1.PR_DrawText("1/2'' Elastic", xyEpt1, 0.125, 0)
            ARMDIA1.PR_DrawTextCenter("1/2'' Elastic", xyEpt1, 0.125, (90 * (ARMDIA1.PI / 180)))
        End If

        elastictop.X = tempDataX
        elastictop.y = tempDataY

        'Distal Elastic and Cutback
        elastictop.X(1) = Profile.X(1)
        elastictop.y(1) = Profile.y(1) / 2
        elastictop.y(2) = Profile.y(1) / 2
        elastictop.X(3) = Profile.X(1) + 0.375
        elastictop.y(3) = Profile.y(1) + 0.375
        ARMDIA1.PR_MakeXY(xyEpt1, Profile.X(1) + 0.75, Profile.y(1) / 2)

        If Val(strFirstTape) = 11 And Val(txtAge.Text) <= 10 Then
            elastictop.X(2) = Profile.X(1) + 0.375
            elastictop.X(3) = Profile.X(1) + 0.375
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
            ARMDIA1.PR_DrawPoly(elastictop)
            ARMDIA1.PR_SetLayer("Notes")
            ''-------ARMDIA1.PR_DrawText("1/2'' Elastic", xyEpt1, 0.125, 0)
            ARMDIA1.PR_DrawTextCenter("1/2'' Elastic", xyEpt1, 0.125, (90 * (ARMDIA1.PI / 180)))

        ElseIf Val(strFirstTape) = 11 And Val(txtAge.Text) > 10 Then
            elastictop.X(2) = Profile.X(1) + 0.75
            elastictop.X(3) = Profile.X(1) + 0.75
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
            ARMDIA1.PR_DrawPoly(elastictop)

        ElseIf Val(strFirstTape) > 11 And Val(txtAge.Text) <= 10 Then
            'For children with sleeves above elbow then use 1/2" elastic
            ARMDIA1.PR_SetLayer("Notes")
            ''-----------ARMDIA1.PR_DrawText("1/2'' Elastic", xyEpt1, 0.125, 0)
            ARMDIA1.PR_DrawTextCenter("1/2'' Elastic", xyEpt1, 0.125, (90 * (ARMDIA1.PI / 180)))
        End If
        ARMDIA1.g_nCurrTextAngle = 0
    End Sub
    Private Sub RightDraw_Click()
        Dim nValue As Object
        Dim n2 As Object
        Dim n1 As Object
        Dim ff As Object
        Dim gg As Object

        Dim conures, Profile, Gauntend As New ARMDIA1.curve
        Dim tps, ii, nn, numtapes As Object
        Dim fNum As Short
        Dim xyPleat, xyThumbhole, xyFin, xyMidPoint, xyTHole1, xyPt, xyPt1, xyCen, xyBegin, xyTHole2, xyPt2, xyElbow1 As ARMDIA1.XY
        Dim nWrist, hole, pl, actfabric, Pleatsgn, GauntSize, nPalm, lenty As Object
        Dim m, nghk, nlen, nAng1, nAdj1, nOpp1, nOpp2, nAdj2, nAng2, midylen, n, valcheck As Object
        Dim xyThumb7, xyThumb5, xyThumb3, xyThumb1, xyThumbStart, xyThumb2, xyThumb4, xyThumb6, xyThumb8 As ARMDIA1.XY
        Dim nTextHt, thmAng1, thm2, thm1, thm3, thmang2, numtap As Object
        Dim xyP As ARMDIA1.XY
        Dim w, SeamDist, X As Object
        Dim v As Short
        Dim xyElbMark1, xyElbMark2 As ARMDIA1.XY
        Dim iPleats As Short
        ''--------------Draw.Enabled = False
        SeamDist = 0.1875
        valcheck = NoRightVals()


        If RightcboPressure.Text = "" Then
            MsgBox("Pressure has not been selected", 48, "Arm Details")
            Exit Sub
        End If
        If RightcboFabric.Text = "" Then
            MsgBox("Fabric has not been selected", 48, "Arm Details")
            Exit Sub
        End If

        PR_GetRightStartEndTapes()
        'Set style first and last tape
        If ARMDIA1.g_iRightStyleFirstTape = -1 Then ARMDIA1.g_iRightStyleFirstTape = Val(strFirstTape)
        If ARMDIA1.g_iRightStyleLastTape = -1 Then ARMDIA1.g_iRightStyleLastTape = Val(strLastTape)

        If valcheck = 0 Then
            MsgBox("No tape measurements have been entered", 48, "Arm Details")
            Exit Sub
        End If

        If ARMDIA1.g_iRightStyleLastTape < ARMDIA1.g_iRightStyleFirstTape Then
            MsgBox("Chosen distal tape is above Proximal tape!", 48, "Arm Details")
            Exit Sub
        End If

        numtapes = (Val(strLastTape) - Val(strFirstTape)) + 1
        If numtapes = 1 Then
            MsgBox("At least two tapes have to be entered before drawing", 48, "Arm Details")
            Exit Sub
        End If

        'Check if flap is style D and that a waist circumference is given
        If InStr(1, RightcboFlaps.Text, "D", 0) > 0 And Val(RighttxtWaistCir.Text) = 0 Then
            MsgBox("A Waist circumference must be given for a D-Style flap", 48, "Arm Details")
            Exit Sub
        End If
        Dim tmpFirst, DeleteChecker, offnumtapes, actnumtapes, tmpLast As Object
        offnumtapes = (CDbl(strLastTape) - CDbl(strFirstTape)) + 1
        gg = 0
        tmpFirst = ""
        Do While gg < 18
            If RightLength(gg).Text <> "" Then
                tmpFirst = gg + 1
                Exit Do
            End If
            gg = gg + 1
        Loop
        ' secondlast & last tapes
        ff = 17
        tmpLast = 0
        Do While ff > 0
            If RightLength(ff).Text <> "" Then
                tmpLast = ff + 1
                Exit Do
            End If
            ff = ff - 1
        Loop
        actnumtapes = (tmpLast - tmpFirst) + 1
        PR_GetRightPressure()
        PR_GetRightFabric()
        PR_CheckRightValues()
        PR_GetRightGauntletDetails()
        PR_RightGauntletHoleCheck()
        PR_GetRightStartEndTapes()
        ''-------------PR_SetRresults()
        ''---------------PR_SetGlobals()

        'check Gauntlet details
        If RightchkGauntlets.Checked = True Then
            If RightcboPalm.SelectedIndex = -1 Then
                MsgBox("No value for palm tape", 48, "Gauntlet Details")
                Exit Sub
            End If
            If RightcboWrist.SelectedIndex = -1 Then
                MsgBox("No value for wrist tape", 48, "Gauntlet Details")
                Exit Sub
            End If
        End If
        ' set Righttapewidths
        For w = 0 To 17
            'UPGRADE_WARNING: Couldn't resolve default property of object w. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            X = w + 1
            PR_SetRightDistances(w, X)
            ''-------------PR_CheckFlaps(w)
        Next w

        Profile.n = numtapes
        If numtapes > 3 Then
            ii = 1
            Dim X1(Val(strLastTape) - Val(strFirstTape) + 1) As Double
            Dim Y1(Val(strLastTape) - Val(strFirstTape) + 1) As Double
            For tps = Val(strFirstTape) To Val(strLastTape)
                Dim Measurement As Double = Val(RightMeasurements(tps))
                If Measurement <= 0 Then
                    ''tps = tps + 1
                    Profile.n = Profile.n - 1
                    Continue For
                End If
                actfabric = ((Val(Measurement) * (100 - Val(Rightreds(tps - 1).Text))) / 100) / 2
                X1(ii) = Val(Righttapewidths(tps))
                Y1(ii) = actfabric + SeamDist
                ii = ii + 1
            Next tps
            Profile.X = X1
            Profile.y = Y1
        End If
        If numtapes = 3 Then
            tps = Val(strFirstTape)
            Dim X1(3), Y1(3) As Double
            Profile.X = X1
            Profile.y = Y1
            actfabric = ((Val(RightMeasurements(tps)) * (100 - Val(Rightreds(tps - 1).Text))) / 100) / 2
            Profile.X(1) = Val(Righttapewidths(tps))
            Profile.y(1) = actfabric + SeamDist
            tps = tps + 2
            actfabric = ((Val(RightMeasurements(tps)) * (100 - Val(Rightreds(tps - 1).Text))) / 100) / 2
            Profile.X(3) = Val(Righttapewidths(tps))
            Profile.y(3) = actfabric + SeamDist
            tps = tps - 1
            If RightMeasurements(tps) <> 0 Then 'And RightcboFlaps.Text = ""
                actfabric = ((Val(RightMeasurements(tps)) * (100 - Val(Rightreds(tps - 1).Text))) / 100) / 2
                Profile.X(2) = Val(Righttapewidths(tps))
                Profile.y(2) = actfabric + SeamDist
            End If
        End If
        If numtapes = 3 Then
            tps = Val(strFirstTape)
            If Str(RightMeasurements(tps + 1)) = "" And RightchkGauntlets.Checked = True Then
                Profile.n = 2
                actfabric = ((Val(RightMeasurements(tps)) * (100 - Val(Rightreds(tps - 1).Text))) / 100) / 2
                Profile.X(1) = Val(Righttapewidths(tps))
                Profile.y(1) = actfabric + SeamDist
                tps = tps + 2
                actfabric = ((Val(RightMeasurements(tps)) * (100 - Val(Rightreds(tps - 1).Text))) / 100) / 2
                ''---------------Profile.X(2) = Profile.X(1) + Val(txtPalmWristDist.Text)
                Profile.X(2) = Profile.X(1) + Val(RighttxtPalmWrist.Text)
                Profile.y(2) = actfabric + SeamDist
            End If
        End If
        If numtapes = 2 And RightchkGauntlets.Checked = False Then
            ii = 1
            Dim X1(2), Y1(2) As Double
            Profile.X = X1
            Profile.y = Y1
            For tps = Val(strFirstTape) To Val(strLastTape)
                If RightMeasurements(tps) <= 0 Then
                    tps = tps + 1
                    Profile.n = Profile.n - 1
                End If
                actfabric = ((Val(RightMeasurements(tps)) * (100 - Val(Rightreds(tps - 1).Text))) / 100) / 2
                Profile.X(ii) = Val(Righttapewidths(tps))
                Profile.y(ii) = actfabric + SeamDist
                ii = ii + 1
            Next tps
        End If
        iPleats = 0
        If Val(RighttxtWristPleat1.Text) > 0 And RighttxtWristPleat1.Enabled Then iPleats = iPleats + 1
        If Val(RighttxtWristPleat2.Text) > 0 And RighttxtWristPleat2.Enabled Then iPleats = iPleats + 1
        If Val(RighttxtShoulderPleat1.Text) > 0 And RighttxtShoulderPleat1.Enabled Then iPleats = iPleats + 1
        If Val(RighttxtShoulderPleat2.Text) > 0 And RighttxtShoulderPleat2.Enabled Then iPleats = iPleats + 1
        'As we can't have a pleat between palm and wrist we fudge it here to cut down on code.
        If RightchkGauntlets.Checked = True Then iPleats = iPleats + 1

        Dim sDrawFile As String = fnGetSettingsPath("PathDRAW") & "\draw.d"
        ARMDIA1.fNum = ARMDIA1.FN_Open(sDrawFile, txtPatientName.Text, txtFileNo.Text, txtSleeve.Text, "Draw")
        'FileClose(ARMDIA1.fNum)
        Dim sStyle As String = ""
        PR_GetRightStyle(sStyle)
        ARMDIA1.g_sID = sStyle & txtFileNo.Text & txtSleeve.Text
        Me.Hide()
        PR_GetInsertionPoint()
        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
        ARMDIA1.PR_DrawFitted(Profile, True)
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "Profile") 'g_sID from FN_Open
        ARMDIA1.PR_NamedHandle("hSleeveProfile")

        ARMDIA1.PR_MakeXY(xyPt1, Profile.X(1), 0)
        ARMDIA1.PR_MakeXY(xyPt2, Profile.X(Profile.n), 0)
        ARMDIA1.PR_DrawLine(xyPt1, xyPt2) 'Fold Line
        ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "FoldLine") 'sFileNo, sSide from FN_Open

        ARMDIA1.PR_SetLayer("Notes")
        PR_DrawLineOffset(xyPt1, xyPt2, 0.6875) 'Inner tram line
        PR_DrawLineOffset(xyPt1, xyPt2, 0.1875) 'seam line at 3/16ths

        If RightchkGauntlets.Checked = False Then
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
            ARMDIA1.PR_MakeXY(xyPt1, Profile.X(1), 0)
            If Val(strFirstTape) = 11 Then
                ARMDIA1.PR_MakeXY(xyPt2, Profile.X(1), Profile.y(1) / 2)
            Else
                ARMDIA1.PR_MakeXY(xyPt2, Profile.X(1), Profile.y(1))
            End If
            ARMDIA1.PR_DrawLine(xyPt1, xyPt2) 'Wrist Line
            ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "WristLine") 'sFileNo, sSide from FN_Open
        End If
        If RightcboFlaps.Text = "" Or RightcboFlaps.Text = " " Then
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
            If Val(strLastTape) = 11 Then
                ARMDIA1.PR_MakeXY(xyPt1, Profile.X(Profile.n), Profile.y(Profile.n) / 2)
            Else
                ARMDIA1.PR_MakeXY(xyPt1, Profile.X(Profile.n), Profile.y(Profile.n))
            End If
            ARMDIA1.PR_MakeXY(xyPt2, Profile.X(Profile.n), 0)
            ARMDIA1.PR_DrawLine(xyPt1, xyPt2) 'Shoulder Line
            ARMDIA1.PR_AddEntityID(ARMDIA1.g_sID, "", "ShoulderLine") 'sFileNo, sSide from FN_Open
        End If
        nTextHt = 0.125

        ' Draw Circular Stump
        ARMDIA1.g_sPatient = txtPatientName.Text
        If RightchkStump.Checked = True Then
            ARMDIA1.PR_MakeXY(xyPt1, Profile.X(1) - 5, Profile.y(1) / 2)
            PR_DrawCircularStump(xyPt1, Profile.y(1), txtAge.Text)
            'Add Stump text
            ARMDIA1.PR_SetLayer("Notes")
            ARMDIA1.PR_MakeXY(xyPt1, Profile.X(1) + 0.25, Profile.y(1) / 2)
            ARMDIA1.PR_DrawText("STUMP", xyPt1, nTextHt, 90 * (ARMDIA1.PI / 180))
        End If
        ii = 1
        For tps = Val(strFirstTape) To Val(strLastTape)
            'Elbow Mark
            If tps = 11 And g_sRightHoleCheck <> "True" Then
                ARMDIA1.PR_SetLayer("Notes")
                ARMDIA1.PR_MakeXY(xyElbMark1, Profile.X(ii), Profile.y(ii) - 0.4)
                ARMDIA1.PR_MakeXY(xyElbMark2, Profile.X(ii), Profile.y(ii) + 0.4)
                ARMDIA1.PR_DrawLine(xyElbMark1, xyElbMark2)
            End If
            If tps = 10 And g_sRightHoleCheck = "True" Then
                ARMDIA1.PR_SetLayer("Notes")
                ARMDIA1.PR_MakeXY(xyElbMark1, Profile.X(ii), Profile.y(ii) - 0.4)
                ARMDIA1.PR_MakeXY(xyElbMark2, Profile.X(ii), Profile.y(ii) + 0.4)
                ARMDIA1.PR_DrawLine(xyElbMark1, xyElbMark2)
            End If
            ii = ii + 1
        Next tps
        If Val(RighttxtWristPleat1.Text) > 0 And RighttxtWristPleat1.Enabled Then
            nTextHt = 0.125
            ARMDIA1.PR_SetLayer("Construct")
            If RightchkGauntlets.Checked = True Then
                n1 = ARMDIA1.max(Val(strWristNo) - Val(strFirstTape), 2)
                n2 = n1 + 1
            Else
                n1 = 1
                n2 = 2
            End If
            pl = (Profile.X(n2) - Profile.X(n1)) / 2
            ARMDIA1.PR_MakeXY(xyPleat, Profile.X(n1) + pl, 0.3375)
            ARMDIA1.PR_DrawText("Pleat", xyPleat, nTextHt, 0)
        End If
        If Val(RighttxtWristPleat2.Text) > 0 And RighttxtWristPleat2.Enabled Then
            nTextHt = 0.125
            ARMDIA1.PR_SetLayer("Construct")
            If RightchkGauntlets.Checked = True Then
                n1 = ARMDIA1.max(Val(strWristNo) - Val(strFirstTape), 2) + 1
                n2 = n1 + 1
            Else
                n1 = 2
                n2 = 3
            End If
            pl = (Profile.X(n2) - Profile.X(n1)) / 2
            ARMDIA1.PR_MakeXY(xyPleat, Profile.X(n1) + pl, 0.3375)
            ARMDIA1.PR_DrawText("Pleat", xyPleat, nTextHt, 0)
        End If
        If Val(RighttxtShoulderPleat1.Text) > 0 And RighttxtShoulderPleat1.Enabled Then
            nTextHt = 0.125
            ARMDIA1.PR_SetLayer("Construct")
            n = Profile.n
            m = Profile.n - 1
            pl = (Profile.X(n) - Profile.X(m)) / 2
            ARMDIA1.PR_MakeXY(xyPleat, Profile.X(m) + pl, 0.3375)
            ARMDIA1.PR_DrawText("Pleat", xyPleat, nTextHt, 0)
        End If

        If Val(RighttxtShoulderPleat2.Text) > 0 And RighttxtShoulderPleat2.Enabled Then
            nTextHt = 0.125
            ARMDIA1.PR_SetLayer("Construct")
            n = Profile.n - 1
            m = Profile.n - 2
            pl = (Profile.X(n) - Profile.X(m)) / 2
            ARMDIA1.PR_MakeXY(xyPleat, Profile.X(m) + pl, 0.3375)
            ARMDIA1.PR_DrawText("Pleat", xyPleat, nTextHt, 0)
        End If

        'Draw Gauntlet
        If RightchkGauntlets.Checked = True Then
            DrawRightGauntlet(Profile, numtapes)
        End If

        'Patient Details
        TypeRightDetails(Profile, numtapes)
        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)

        RightElastic(Profile)

        ' Shoulder Flaps
        If RightcboFlaps.Text = "" Or VB.Left(RightcboFlaps.Text, 4) = "None" Then
            ''---------------- g_sFlap = "None"
        ElseIf (VB.Left(RightcboFlaps.Text, 4) = "Vest") Then
            ''---------------g_sFlap = ""
            ''---------------g_sFlap = RightcboFlaps.Text
            ''-------------------------PR_DrawRaglan("Vest", txtVestRaglan.Text, txtAge.Text, txtVestID)

        ElseIf (VB.Left(RightcboFlaps.Text, 4) = "Body") Then
            ''----------------------g_sFlap = ""
            ''----------------------g_sFlap = RightcboFlaps.Text
            ''-----------------------PR_DrawRaglan("Body", txtVestRaglan.Text, txtAge.Text, txtVestID)
        Else
            ''------------------------g_sFlap = RightcboFlaps.Text
            PR_DrawRightShoulderFlaps(Profile, numtapes)
        End If

        ''----------------ARMDIA1.g_sFabric = RightcboFabric.Text


        'Vestarm and Profile Origin Marker data base update
        'For editing purposes the tape lengths are stored with
        'the Profile Origin Marker
        Dim sJustifiedString As String
        For ii = 0 To 17
            nValue = Val(RightLength(ii).Text)
            If nValue <> 0 And ii >= (Val(strFirstTape) - 1) And ii <= (Val(strLastTape) - 1) Then
                nValue = nValue * 10 'Shift decimal place to right
                sJustifiedString = New String(" ", 3)
                sJustifiedString = RSet(Trim(Str(nValue)), Len(sJustifiedString))
            Else
                sJustifiedString = New String(" ", 3)
            End If
            'Tape values
            ''----------------------------g_sEditLengths = g_sEditLengths & sJustifiedString
        Next ii
        'Update
        ''-------------------------------DataBaseDataUpDate("Draw")
        ''Added for #166 in the issue list
        ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)
        ii = 1
        For tps = Val(strFirstTape) To Val(strLastTape)
            Dim Measurement As Double = Val(RightMeasurements(tps))
            If Measurement <= 0 Then
                tps = tps + 1
                Measurement = Val(RightMeasurements(tps))
            End If
            ARMDIA1.PR_MakeXY(xyPt1, Profile.X(ii), 0)
            '-------ARMDIA1.PR_PutTapeLabel(tps, xyPt1, Measurement, Rightmms(tps - 1), Rightgms(tps - 1), Rightreds(tps - 1))
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                ''ARMDIA1.PR_PutTapeLabel(tps, xyPt1, fnDisplayToCM(Measurement), Rightmms(tps - 1), Rightgms(tps - 1), Rightreds(tps - 1))
                ARMDIA1.PR_PutTapeLabel(tps, xyPt1, Val(RightLength(tps - 1).Text), Rightmms(tps - 1), Rightgms(tps - 1), Rightreds(tps - 1))
            Else
                ARMDIA1.PR_PutTapeLabel(tps, xyPt1, Measurement, Rightmms(tps - 1), Rightgms(tps - 1), Rightreds(tps - 1))
            End If
            PR_DrawRuler(tps, xyPt1, RightlblTape(tps - 1).Text)
            ii = ii + 1
        Next tps

        'For Body suit et.al. start program to draw the mesh
        PR_DrawMesh()
        saveRightInfoToDWG()
        g_bIsClose = True
        Me.Close()
        Draw.Enabled = True
        FileClose(ARMDIA1.fNum)
        PR_SetLayer("Construct")
        PR_DrawXMarker()
    End Sub
    Private Function saveRightInfoToDWG()
        Try
            Dim _sClass As New SurroundingClass()
            If (_sClass.GetXrecord("SNRightInfo", "SNRIGHTDIC") IsNot Nothing) Then
                _sClass.RemoveXrecord("SNRightInfo", "SNRIGHTDIC")
            End If
            Dim resbuf As New ResultBuffer

            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_0.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_1.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_2.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_3.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_4.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_5.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_6.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_7.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_8.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_9.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_10.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_11.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_12.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_13.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_14.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_15.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_16.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _Length_17.Text))
            Dim ii As Double
            For ii = 0 To 17
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), RightLength(ii).Text))
            Next

            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RightcboFabric.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RightcboPressure.Text))

            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_0.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_1.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_2.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_3.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_4.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_5.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_6.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_7.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_8.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_9.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_10.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_11.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_12.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_13.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_14.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_15.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_16.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _mms_17.Text))
            For ii = 0 To 17
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), Rightmms(ii).Text))
            Next

            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_0.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_1.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_2.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_3.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_4.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_5.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_6.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_7.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_8.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_9.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_10.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_11.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_12.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_13.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_14.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_15.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_16.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _gms_17.Text))

            For ii = 0 To 17
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), Rightgms(ii).Text))
            Next

            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_0.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_1.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_2.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_3.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_4.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_5.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_6.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_7.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_8.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_9.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_10.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_11.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_12.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_13.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_14.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_15.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_16.Text))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), _reds_17.Text))

            For ii = 0 To 17
                resbuf.Add(New TypedValue(CInt(DxfCode.Text), Rightreds(ii).Text))
            Next
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), RightchkNone.Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), RightchkStump.Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RightcboDistalTape.Text))
            'Dim sGauntlet As String
            'If chkGauntlets.Checked = True Then
            '    sGauntlet = "1"
            'Else
            '    sGauntlet = "0"
            'End If
            'If chkEnclosed.Checked = True Then
            '    sGauntlet = sGauntlet & " 1"
            'Else
            '    sGauntlet = sGauntlet & " 0"
            'End If
            'If chkDetachable.Checked = True Then
            '    sGauntlet = sGauntlet & " 1"
            'Else
            '    sGauntlet = sGauntlet & " 0"
            'End If
            'If chkNoThumb.Checked = True Then
            '    sGauntlet = sGauntlet & " 1"
            'Else
            '    sGauntlet = sGauntlet & " 0"
            'End If
            'sGauntlet = sGauntlet & " " & Str(cboWrist.SelectedIndex + 1) & " " & Str(cboPalm.SelectedIndex + 1) & " " & txtThumb.Text & " " & txtCircum.Text & " " & txtPalmWrist.Text & " " & txtGauntletExtension.Text
            'resbuf.Add(New TypedValue(CInt(DxfCode.Text), sGauntlet))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), RightchkGauntlets.Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RightcboPalm.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RightcboWrist.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RighttxtPalmWrist.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RighttxtCircum.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RighttxtThumb.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RighttxtGauntletExtension.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), RightchkDetachable.Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), RightchkNoThumb.Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Bool), RightchkEnclosed.Checked))

            For ii = 0 To 1
                resbuf.Add(New TypedValue(CInt(DxfCode.Bool), RightoptProximalTape(ii).Checked))
            Next
            'resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optProximalTape(0).Checked))
            'resbuf.Add(New TypedValue(CInt(DxfCode.Bool), optProximalTape(1).Checked))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RightcboProximalTape.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RightcboFlaps.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RighttxtStrapLength.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RighttxtFrontStrapLength.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RighttxtCustFlapLength.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RighttxtWaistCir.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RightcboDistalTape.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RightcboProximalTape.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RighttxtWristPleat1.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RighttxtWristPleat2.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RighttxtShoulderPleat1.Text))
            resbuf.Add(New TypedValue(CInt(DxfCode.Text), RighttxtShoulderPleat2.Text))

            _sClass.SetXrecord(resbuf, "SNRightInfo", "SNRIGHTDIC")
        Catch ex As Exception
        End Try
    End Function
    Private Function readRightDWGInfo()
        Try
            Dim _sClass As New SurroundingClass()
            Dim resbuf As New ResultBuffer
            resbuf = _sClass.GetXrecord("SNRightInfo", "SNRIGHTDIC")
            If (resbuf IsNot Nothing) Then
                Dim arr() As TypedValue = resbuf.AsArray()
                '_Length_0.Text = arr(0).Value
                'Length_Leave(_Length_0, Nothing)
                '_Length_1.Text = arr(1).Value
                'Length_Leave(_Length_1, Nothing)
                '_Length_2.Text = arr(2).Value
                'Length_Leave(_Length_2, Nothing)
                '_Length_3.Text = arr(3).Value
                '_Length_4.Text = arr(4).Value
                '_Length_5.Text = arr(5).Value
                '_Length_6.Text = arr(6).Value
                '_Length_7.Text = arr(7).Value
                '_Length_8.Text = arr(8).Value
                '_Length_9.Text = arr(9).Value
                '_Length_10.Text = arr(10).Value
                '_Length_11.Text = arr(11).Value
                '_Length_12.Text = arr(12).Value
                '_Length_13.Text = arr(13).Value
                '_Length_14.Text = arr(14).Value
                '_Length_15.Text = arr(15).Value
                '_Length_16.Text = arr(16).Value
                '_Length_17.Text = arr(17).Value

                Dim ii As Double
                For ii = 0 To 17
                    RightLength(ii).Text = arr(ii).Value
                Next ii

                RightcboFabric.Text = arr(18).Value
                RightcboPressure.Text = arr(19).Value

                '_mms_0.Text = arr(20).Value
                '_mms_1.Text = arr(21).Value
                '_mms_2.Text = arr(22).Value
                '_mms_3.Text = arr(23).Value
                '_mms_4.Text = arr(24).Value
                '_mms_5.Text = arr(25).Value
                '_mms_6.Text = arr(26).Value
                '_mms_7.Text = arr(27).Value
                '_mms_8.Text = arr(28).Value
                '_mms_9.Text = arr(29).Value
                '_mms_10.Text = arr(30).Value
                '_mms_11.Text = arr(31).Value
                '_mms_12.Text = arr(32).Value
                '_mms_13.Text = arr(33).Value
                '_mms_14.Text = arr(34).Value
                '_mms_15.Text = arr(35).Value
                '_mms_16.Text = arr(36).Value
                '_mms_17.Text = arr(37).Value
                For ii = 20 To 37
                    Rightmms(ii - 20).Text = arr(ii).Value
                Next ii

                '_gms_0.Text = arr(38).Value
                '_gms_1.Text = arr(39).Value
                '_gms_2.Text = arr(40).Value
                '_gms_3.Text = arr(41).Value
                '_gms_4.Text = arr(42).Value
                '_gms_5.Text = arr(43).Value
                '_gms_6.Text = arr(44).Value
                '_gms_7.Text = arr(45).Value
                '_gms_8.Text = arr(46).Value
                '_gms_9.Text = arr(47).Value
                '_gms_10.Text = arr(48).Value
                '_gms_11.Text = arr(49).Value
                '_gms_12.Text = arr(50).Value
                '_gms_13.Text = arr(51).Value
                '_gms_14.Text = arr(52).Value
                '_gms_15.Text = arr(53).Value
                '_gms_16.Text = arr(54).Value
                '_gms_17.Text = arr(55).Value
                For ii = 38 To 55
                    Rightgms(ii - 38).Text = arr(ii).Value
                Next ii

                '_reds_0.Text = arr(56).Value
                '_reds_1.Text = arr(57).Value
                '_reds_2.Text = arr(58).Value
                '_reds_3.Text = arr(59).Value
                '_reds_4.Text = arr(60).Value
                '_reds_5.Text = arr(61).Value
                '_reds_6.Text = arr(62).Value
                '_reds_7.Text = arr(63).Value
                '_reds_8.Text = arr(64).Value
                '_reds_9.Text = arr(65).Value
                '_reds_10.Text = arr(66).Value
                '_reds_11.Text = arr(67).Value
                '_reds_12.Text = arr(68).Value
                '_reds_13.Text = arr(69).Value
                '_reds_14.Text = arr(70).Value
                '_reds_15.Text = arr(71).Value
                '_reds_16.Text = arr(72).Value
                '_reds_17.Text = arr(73).Value
                For ii = 56 To 73
                    Rightreds(ii - 56).Text = arr(ii).Value
                Next ii
                RightchkNone.Checked = arr(74).Value
                RightchkStump.Checked = arr(75).Value
                RightcboDistalTape.Text = arr(76).Value
                RightchkGauntlets.Checked = arr(77).Value
                RightcboPalm.Text = arr(78).Value
                RightcboWrist.Text = arr(79).Value
                RighttxtPalmWrist.Text = arr(80).Value
                RighttxtCircum.Text = arr(81).Value
                RighttxtThumb.Text = arr(82).Value
                RighttxtGauntletExtension.Text = arr(83).Value
                RightchkDetachable.Checked = arr(84).Value
                RightchkNoThumb.Checked = arr(85).Value
                RightchkEnclosed.Checked = arr(86).Value

                For ii = 87 To 88
                    RightoptProximalTape(ii - 87).Checked = arr(ii).Value
                Next ii

                'optProximalTape(0).Checked = arr(76).Value
                'optProximalTape(1).Checked = arr(77).Value

                RightcboProximalTape.Text = arr(89).Value
                RightcboFlaps.Text = arr(90).Value
                RighttxtStrapLength.Text = arr(91).Value
                RighttxtFrontStrapLength.Text = arr(92).Value
                RighttxtCustFlapLength.Text = arr(93).Value
                RighttxtWaistCir.Text = arr(94).Value
                RightcboDistalTape.Text = arr(95).Value
                RightcboProximalTape.Text = arr(96).Value
                RighttxtWristPleat1.Text = arr(97).Value
                RighttxtWristPleat2.Text = arr(98).Value
                RighttxtShoulderPleat1.Text = arr(99).Value
                RighttxtShoulderPleat2.Text = arr(100).Value

                'Length_Leave(_Length_3, Nothing)
                'Length_Leave(_Length_4, Nothing)
                'Length_Leave(_Length_5, Nothing)
                'Length_Leave(_Length_6, Nothing)
                'Length_Leave(_Length_7, Nothing)
                'Length_Leave(_Length_8, Nothing)
                'Length_Leave(_Length_9, Nothing)
                'Length_Leave(_Length_10, Nothing)
                'Length_Leave(_Length_11, Nothing)
                'Length_Leave(_Length_12, Nothing)
                'Length_Leave(_Length_13, Nothing)
                'Length_Leave(_Length_14, Nothing)
                'Length_Leave(_Length_15, Nothing)
                'Length_Leave(_Length_16, Nothing)
                'Length_Leave(_Length_17, Nothing)

                For ii = 0 To 17
                    RightLength_Leave(RightLength(ii), Nothing)
                Next

                ''-----------------cboFabric_SelectedIndexChanged(Nothing, Nothing)
                ''-----------------------Calculate_Click(Nothing, Nothing)
            End If

        Catch ex As Exception

        End Try
    End Function
    Public Sub PR_DrawRightShoulderFlaps(ByRef Profile As ARMDIA1.curve, ByRef numtapes As Object)
        Dim Strap As Object
        Dim RagAng5 As Object
        Dim RagAng4 As Object
        Dim PhiExtra As Object
        Dim ii As Object
        Dim nNotchToTangent As Object
        Dim nTemplateAngle As Object
        Dim nTemplateRadius As Object
        Dim ShoulderFlap, RaglanFlap As ARMDIA1.curve
        Dim LastTapeValue As Object
        Dim xyPt1, xyPt2 As ARMDIA1.XY
        Dim xyTopFlap4, xyTopFlap2, xyTopFlap1, xyTopFlap3, xyTopFlap5 As ARMDIA1.XY
        Dim xyTopFlap9, xyTopFlap7, xyTopFlap6, xyTopFlap8, xyTopFlap10 As ARMDIA1.XY
        Dim xyBotFlap4, xyBotFlap2, xyBotFlap1, xyBotFlap3, xyBotFlap5 As ARMDIA1.XY
        Dim xyPotA, xyFlap4, xyFlap2, xyFlap1, xyFlap3, xyFlap5, xyPotB As ARMDIA1.XY
        Dim Alpha, BottomRadius, Delta, Omega, Beta, Phi, Theta, PI, TopArcIncrement As Object
        Dim TopRadius, LittleBit, opp, h1, x1, y1, adj, hyp, Change, FlapMarkerAngle As Object
        Dim xyFlapText1 As ARMDIA1.XY
        Dim xyFlapText2, xyCentre, xyMid, xyFlapMarker, xyStrapText As ARMDIA1.XY
        Dim FlapLength, FlapMarkerLength As Object
        Dim TopArcAngle, TopArcRadius, TopArcLength, ArcLength, CircleCircum, BottomArcLength, BottomArcRadius, BottomArcAngle As Object
        Dim xfablen, TempMarker, xfabby, xfabdist As Object
        Dim Arrow As String
        Dim xyRaglan6, xyRaglan4, xyRaglan2, xyRaglan1, xyRaglan3, xyRaglan5, xyRaglan7 As ARMDIA1.XY
        Dim TopRagLanAngle, RaglanBottom, RagAng2, Rag4, Rag2, Rag1, Rag3, RagAng1, RagAng3, RaglanTip, InnerRaglanTip As Object
        Dim nNotchOffset As Double
        BottomArcAngle = 47.5
        Arrow = "closed arrow"
        PI = 4 * System.Math.Atan(1)
        'If RightcboFlaps <> "" And RightcboFlaps <> "Vest Raglan" Then
        If RightcboFlaps.Text <> "" And RightcboFlaps.Text <> "Vest" And RightcboFlaps.Text <> "Body" Then
            ' Check for Custom Flap RightLength
            If Val(RighttxtCustFlapLength.Text) > 0 Then
                g_sinchflag = Righttxtinchflag.Text
                g_sFlapLength = cmtoinches(Righttxtinchflag.Text, RighttxtCustFlapLength)
                LastTapeValue = g_sFlapLength
            Else
                'Standard flap length
                LastTapeValue = Val(strLastTape) - 1
                LastTapeValue = cmtoinches(Righttxtinchflag.Text, RightLength(LastTapeValue))
                LastTapeValue = (LastTapeValue / 3.14) * 0.92
            End If


            ' Calculate Bottom part of curve with 5 Points
            If (Profile.y(Profile.n)) >= 5.5 Then
                BottomRadius = LastTapeValue - 0.625
                RaglanBottom = 3.5
                'TopRagLanAngle = 20
                TopRagLanAngle = 17
                RaglanTip = 1.9375
                InnerRaglanTip = 2.173
                nNotchOffset = 0.875 '14/16 ths Raglan Template s287
            ElseIf ((Profile.y(Profile.n)) < 5.5) And ((Profile.y(Profile.n)) > 3) Then
                BottomRadius = 2.8
                RaglanBottom = 2.5
                'TopRagLanAngle = 20
                TopRagLanAngle = 17
                RaglanTip = 1.05
                InnerRaglanTip = 1.22
                nNotchOffset = 0.75 '12/16 ths Raglan Template s289
            ElseIf ((Profile.y(Profile.n)) <= 3) Then
                BottomRadius = 2
                RaglanBottom = 2.0625
                RaglanTip = 0.625
                'TopRagLanAngle = 15
                TopRagLanAngle = 15
                InnerRaglanTip = 0.727
                nNotchOffset = 0.6875 '11/16 ths, Raglan Template s290
            End If
            nTemplateRadius = 3.5
            nTemplateAngle = 40
            nNotchToTangent = 2.125

            'Calc length of bottom Arc
            BottomArcRadius = BottomRadius
            CircleCircum = PI * (2 * (BottomArcRadius))
            BottomArcLength = (BottomArcAngle / 360) * CircleCircum

            ShoulderFlap.n = 10
            Dim X(10), Y(10) As Double
            X(1) = Profile.X(Profile.n) + LastTapeValue
            Y(1) = 0
            'ShoulderFlap.X(1) = Profile.X(Profile.n) + LastTapeValue
            'ShoulderFlap.y(1) = 0
            ShoulderFlap.X = X
            ShoulderFlap.y = Y
            For ii = 1 To 4
                Theta = (ii * 11.875 * PI) / 180
                adj = System.Math.Cos(Theta) * BottomRadius
                opp = System.Math.Sin(Theta) * BottomRadius
                LittleBit = BottomRadius - adj
                ShoulderFlap.X(ii + 1) = Profile.X(Profile.n) + LastTapeValue - LittleBit
                ShoulderFlap.y(ii + 1) = 0 + opp
            Next ii
            ARMDIA1.PR_MakeXY(xyBotFlap5, ShoulderFlap.X(5), ShoulderFlap.y(5))

            'Calculate 6 angles and points for top part of curve
            ARMDIA1.PR_MakeXY(xyTopFlap1, Profile.X(Profile.n), Profile.y(Profile.n))
            ShoulderFlap.X(10) = xyTopFlap1.X
            ShoulderFlap.y(10) = xyTopFlap1.y

            'midpoint
            h1 = ARMDIA1.FN_CalcLength(xyTopFlap1, xyBotFlap5)
            h1 = h1 / 2
            Omega = ARMDIA1.FN_CalcAngle(xyBotFlap5, xyTopFlap1)
            Omega = 180 - Omega
            Omega = (Omega * PI) / 180
            opp = System.Math.Abs(h1 * System.Math.Sin(Omega))
            adj = System.Math.Abs(h1 * System.Math.Cos(Omega))
            ARMDIA1.PR_MakeXY(xyMid, xyBotFlap5.X - adj, xyBotFlap5.y + opp)

            'Centre of Top Arc
            Delta = ARMDIA1.FN_CalcAngle(xyBotFlap5, xyTopFlap1)
            Delta = Delta - 90
            Delta = (Delta * PI) / 180
            hyp = 4 * h1
            opp = System.Math.Abs(hyp * System.Math.Sin(Delta))
            adj = System.Math.Abs(hyp * System.Math.Cos(Delta))
            ARMDIA1.PR_MakeXY(xyCentre, xyMid.X + adj, xyMid.y + opp)
            TopRadius = ARMDIA1.FN_CalcLength(xyCentre, xyTopFlap1)
            TopArcRadius = TopRadius
            Phi = ARMDIA1.FN_CalcAngle(xyCentre, xyTopFlap1)
            Phi = Phi - 180
            Alpha = ARMDIA1.FN_CalcAngle(xyCentre, xyBotFlap5)
            Alpha = Alpha - 180
            TopArcAngle = Alpha - Phi
            TopArcIncrement = (Alpha - Phi) / 5

            'Calc Flap RightLength for position of marker
            CircleCircum = PI * (2 * (TopArcRadius))
            TopArcLength = (TopArcAngle / 360) * CircleCircum
            FlapLength = TopArcLength + BottomArcLength
            FlapMarkerLength = FlapLength / 2.7
            'Minimum length from market to seam is 2.5"
            If FlapLength - FlapMarkerLength < 2.5 Then FlapMarkerLength = FlapLength - 2.5
            If FlapMarkerLength < TopArcLength Then
                FlapMarkerAngle = (FlapMarkerLength * 360) / (PI * (2 * TopArcRadius))
                FlapMarkerAngle = FlapMarkerAngle + Phi
                FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
                opp = System.Math.Abs(TopRadius * System.Math.Sin(FlapMarkerAngle))
                adj = System.Math.Abs(TopRadius * System.Math.Cos(FlapMarkerAngle))
                ARMDIA1.PR_MakeXY(xyFlapMarker, xyCentre.X - adj, xyCentre.y - opp)
            ElseIf FlapMarkerLength > TopArcLength Then
                TempMarker = FlapMarkerLength - TopArcLength
                TempMarker = BottomArcLength - TempMarker
                FlapMarkerAngle = (TempMarker * 360) / (PI * (2 * BottomRadius))
                FlapMarkerAngle = (FlapMarkerAngle * PI) / 180
                opp = System.Math.Abs(BottomArcRadius * System.Math.Sin(FlapMarkerAngle))
                adj = System.Math.Abs(BottomArcRadius * System.Math.Cos(FlapMarkerAngle))
                LittleBit = BottomRadius - adj
                ARMDIA1.PR_MakeXY(xyFlapMarker, Profile.X(Profile.n) + LastTapeValue - LittleBit, 0 + opp)
            ElseIf FlapMarkerLength = TopArcLength Then
                ARMDIA1.PR_MakeXY(xyFlapMarker, xyBotFlap5.X, xyBotFlap5.y)
            End If
            ARMDIA1.PR_SetLayer("Notes")
            'PrintLine(ARMDIA1.fNum, "hEnt = AddEntity(" & ARMDIA1.QQ & "marker" & ARMDIA1.QCQ & "closed arrow" & ARMDIA1.QC & "xyStart.x +" & Str(xyFlapMarker.X) & ARMDIA1.CC & "xyStart.y +" & Str(xyFlapMarker.y) & ",0.25,0.1,225);")
            PR_DrawClosedArrow(xyFlapMarker, 225)
            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)

            For ii = 1 To 4
                PhiExtra = Phi + (ii * TopArcIncrement)
                PhiExtra = (PhiExtra * PI) / 180
                opp = System.Math.Abs(TopRadius * System.Math.Sin(PhiExtra))
                adj = System.Math.Abs(TopRadius * System.Math.Cos(PhiExtra))
                ShoulderFlap.X(10 - ii) = xyCentre.X - adj
                ShoulderFlap.y(10 - ii) = xyCentre.y - opp
            Next ii


            If VB.Left(RightcboFlaps.Text, 6) = "Raglan" Then
                'Copy standard shoulder from profile down
                'only go one vetex past the FlapMarker
                RaglanFlap.n = 0
                Dim XX(10), YY(10) As Double
                XX(1) = 0
                YY(1) = 0
                RaglanFlap.X = XX
                RaglanFlap.y = YY
                For ii = 10 To 1 Step -1
                    RaglanFlap.n = RaglanFlap.n + 1
                    RaglanFlap.X(RaglanFlap.n) = ShoulderFlap.X(ii)
                    RaglanFlap.y(RaglanFlap.n) = ShoulderFlap.y(ii)
                    If ShoulderFlap.X(ii) > xyFlapMarker.X Then Exit For
                Next ii
                ARMDIA1.PR_DrawFitted(RaglanFlap)
            Else
                ShoulderFlap.n = 10
                ARMDIA1.PR_DrawFitted(ShoulderFlap)
            End If

            'Draw Raglan
            If VB.Left(RightcboFlaps.Text, 6) = "Raglan" Then
                Rag1 = System.Math.Sqrt((BottomRadius * BottomRadius) - (nNotchOffset * nNotchOffset))
                Rag1 = BottomRadius - Rag1
                ARMDIA1.PR_MakeXY(xyRaglan1, Profile.X(Profile.n) + LastTapeValue - Rag1, nNotchOffset)
                ARMDIA1.PR_MakeXY(xyRaglan2, xyRaglan1.X - nNotchToTangent, 0)
                RagAng1 = ARMDIA1.FN_CalcAngle(xyRaglan2, xyRaglan1)
                Rag2 = ARMDIA1.FN_CalcLength(xyRaglan2, xyRaglan1)
                Rag2 = Rag2 / 2
                RagAng1 = (RagAng1 * PI) / 180
                opp = System.Math.Abs(Rag2 * System.Math.Sin(RagAng1))
                adj = System.Math.Abs(Rag2 * System.Math.Cos(RagAng1))
                ARMDIA1.PR_MakeXY(xyRaglan3, xyRaglan2.X + adj, xyRaglan2.y + opp)

                RagAng1 = ARMDIA1.FN_CalcAngle(xyRaglan3, xyRaglan1)
                RagAng1 = 90 - RagAng1
                RagAng1 = (RagAng1 * PI) / 180
                Rag3 = ARMDIA1.FN_CalcLength(xyRaglan3, xyRaglan1)
                Rag4 = System.Math.Sqrt((nTemplateRadius * nTemplateRadius) - (Rag3 * Rag3))
                opp = System.Math.Abs(Rag4 * System.Math.Sin(RagAng1))
                adj = System.Math.Abs(Rag4 * System.Math.Cos(RagAng1))
                ARMDIA1.PR_MakeXY(xyRaglan4, xyRaglan3.X - adj, xyRaglan3.y + opp)
                ARMDIA1.PR_DrawArc(xyRaglan4, xyRaglan2, xyRaglan1)

                RagAng2 = (nTemplateAngle * PI) / 180
                opp = System.Math.Abs(RaglanBottom * System.Math.Sin(RagAng2))
                adj = System.Math.Abs(RaglanBottom * System.Math.Cos(RagAng2))
                ARMDIA1.PR_MakeXY(xyRaglan5, xyRaglan1.X + adj, xyRaglan1.y + opp)
                ARMDIA1.PR_DrawLine(xyRaglan1, xyRaglan5)
                RagAng3 = ARMDIA1.FN_CalcAngle(xyRaglan5, xyFlapMarker)

                If RagAng3 > 180 Then
                    RagAng3 = RagAng3 - 180
                    RagAng4 = 180 - 115 - TopRagLanAngle - RagAng3
                ElseIf RagAng3 <= 180 Then
                    RagAng3 = RagAng3 - 90
                    RagAng4 = 90 - RagAng3
                    RagAng4 = (180 - TopRagLanAngle - 115) + RagAng4
                End If
                RagAng5 = RagAng4 - 13
                RagAng4 = (RagAng4 * PI) / 180
                opp = System.Math.Abs(RaglanTip * System.Math.Sin(RagAng4))
                adj = System.Math.Abs(RaglanTip * System.Math.Cos(RagAng4))
                ARMDIA1.PR_MakeXY(xyRaglan6, xyRaglan5.X - adj, xyRaglan5.y + opp)

                'Draw Front line of tip
                RagAng5 = (RagAng5 * PI) / 180
                opp = System.Math.Abs(InnerRaglanTip * System.Math.Sin(RagAng5))
                adj = System.Math.Abs(InnerRaglanTip * System.Math.Cos(RagAng5))
                ARMDIA1.PR_MakeXY(xyRaglan7, xyRaglan5.X - adj, xyRaglan5.y + opp)

                ARMDIA1.PR_DrawLine(xyRaglan5, xyRaglan6)
                ARMDIA1.PR_DrawLine(xyRaglan6, xyFlapMarker)
                ARMDIA1.PR_SetLayer("Notes")
                ARMDIA1.PR_DrawLine(xyRaglan5, xyRaglan7)

                '    PR_AddMarker xyRaglan1, "xyRaglan1"
                '    PR_AddMarker xyRaglan2, "xyRaglan2"
                '    PR_AddMarker xyRaglan3, "xyRaglan3"
                '    PR_AddMarker xyRaglan4, "xyRaglan4"
                '    PR_AddMarker xyRaglan5, "xyRaglan5"
                '    PR_AddMarker xyRaglan6, "xyRaglan6"
                '    PR_AddMarker xyRaglan7, "xyRaglan7"
                '    PR_AddMarker xyRaglan6, "xyRaglan5"
                '    PR_AddMarker xyFlapMarker, "xyFlapMarker"

            End If

            ARMDIA1.PR_SetLayer("Notes")
            ARMDIA1.PR_MakeXY(xyStrapText, xyFlapMarker.X - 1.5, xyFlapMarker.y)
            Strap = "STRAP = " & Trim(ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtStrapLength))) & Chr(34)
            If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                Strap = "STRAP = " & Trim(fnDisplayStrapToCM(cmtoinches(Righttxtinchflag.Text, RighttxtStrapLength)))
            End If
            If Val(RighttxtFrontStrapLength.Text) > 0 Then
                'Display FRONT strap length if given
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    Strap = "BACK  STRAP = " & Trim(fnDisplayStrapToCM(cmtoinches(Righttxtinchflag.Text, RighttxtStrapLength))) & "CM"
                    Strap = Strap & Chr(10) & "FRONT STRAP = " & Trim(fnDisplayStrapToCM(cmtoinches(Righttxtinchflag.Text, RighttxtFrontStrapLength))) & "CM"
                Else
                    Strap = "BACK  STRAP = " & Trim(ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtStrapLength))) & Chr(34)
                    Strap = Strap & Chr(10) & "FRONT STRAP = " & Trim(ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtFrontStrapLength))) & Chr(34)
                End If
            End If
                If Val(RighttxtWaistCir.Text) > 0 Then
                'Display Waist Circumference for D-Style flaps only
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    Strap = Strap & Chr(10) & "WAIST = " & Trim(fnDisplayStrapToCM(cmtoinches(Righttxtinchflag.Text, RighttxtWaistCir)))
                Else
                    Strap = Strap & Chr(10) & "WAIST = " & Trim(ARMEDDIA1.fnInchesToText(cmtoinches(Righttxtinchflag.Text, RighttxtWaistCir))) & Chr(34)
                End If
            End If
                ''ARMDIA1.PR_DrawText(Strap, xyStrapText, 0.125, 0)
                ARMDIA1.PR_DrawMText(Strap, xyStrapText, True)

            xfabby = RightcboFlaps.Text
            xfablen = Len(xfabby)
            xfabdist = (xfablen * 0.075) + 0.1
            ARMDIA1.PR_MakeXY(xyFlapText1, xyFlapMarker.X - xfabdist, xyFlapMarker.y - 0.75)
            ARMDIA1.PR_DrawMText(RightcboFlaps.Text, xyFlapText1, True)

            ARMDIA1.PR_MakeXY(xyPt1, Profile.X(Profile.n), 0)
            ARMDIA1.PR_MakeXY(xyPt2, Profile.X(Profile.n) + LastTapeValue, 0)

            If VB.Left(RightcboFlaps.Text, 6) = "Raglan" Then
                ARMDIA1.PR_MakeXY(xyPotA, xyPt2.X - 0.2875, xyPt2.y)
                ARMDIA1.PR_MakeXY(xyPotB, xyPt2.X - 1.225, xyPt2.y)
                PR_DrawLineOffset(xyPt1, xyPotA, 0.6875) 'Inner tram line
                PR_DrawLineOffset(xyPt1, xyPotB, 0.1875) 'seam line at 3/16ths
            Else
                ARMDIA1.PR_MakeXY(xyPotA, xyPt2.X - 0.09375, xyPt2.y)
                ARMDIA1.PR_MakeXY(xyPotB, xyPt2.X - 0.03125, xyPt2.y)
                PR_DrawLineOffset(xyPt1, xyPotA, 0.6875) 'Inner tram line
                PR_DrawLineOffset(xyPt1, xyPotB, 0.1875) 'seam line at 3/16ths
            End If

            ARMDIA1.PR_SetLayer("Template" & txtSleeve.Text)

            If VB.Left(RightcboFlaps.Text, 6) = "Raglan" Then
                ARMDIA1.PR_MakeXY(xyPotA, xyPt2.X - 1.75, xyPt2.y)
                ARMDIA1.PR_DrawLine(xyPt1, xyRaglan2)
            Else
                ARMDIA1.PR_DrawLine(xyPt1, xyPt2) 'Fold Line
            End If
            ARMDIA1.PR_AddEntityID(ARMDIA1.g_sFileNo, ARMDIA1.g_sSide, "FoldLineFlap") 'sFileNo, sSide from FN_Open

        End If
    End Sub
    Private Function NoRightVals() As Object
        Dim contents, i, total As Object
        i = 0
        contents = 0
        total = 0
        Do While i < 18
            contents = Val(RightLength(i).Text)
            total = total + contents
            i = i + 1
        Loop
        NoRightVals = total
    End Function

    Private Sub RightlblPleat_DoubleClick(eventSender As Object, e As EventArgs) Handles RightlblPleat.DoubleClick
        Dim Index As Short = RightlblPleat.GetIndex(eventSender)
        'Procedure to disable the pleat fields
        'to enable the User to select only the relevent
        'data to be drawn
        'Saves having to delele and re-enter to get variations
        'on pleats
        If System.Drawing.ColorTranslator.ToOle(RightlblPleat(Index).ForeColor) = &HC0C0C0 Then
            'Enable pleat if previously disabled
            RightlblPleat(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
            Select Case Index
                Case 0
                    RighttxtWristPleat1.Enabled = True
                Case 1
                    RighttxtWristPleat2.Enabled = True
                Case 3
                    RighttxtShoulderPleat1.Enabled = True
                Case 2
                    RighttxtShoulderPleat2.Enabled = True
            End Select
        Else
            'Disable pleat if previously enabled
            RightlblPleat(Index).ForeColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            Select Case Index
                Case 0
                    RighttxtWristPleat1.Enabled = False
                Case 1
                    RighttxtWristPleat2.Enabled = False
                Case 3
                    RighttxtShoulderPleat1.Enabled = False
                Case 2
                    RighttxtShoulderPleat2.Enabled = False
            End Select
        End If
    End Sub

    Private Sub RighttxtWristPleat1_Leave(sender As Object, e As EventArgs) Handles RighttxtWristPleat1.Leave
        If Not FN_CheckValue(RighttxtWristPleat1, "First Wrist Pleat") Then
            Exit Sub
        End If
    End Sub

    Private Sub RighttxtWristPleat2_Leave(sender As Object, e As EventArgs) Handles RighttxtWristPleat2.Leave
        If Not FN_CheckValue(RighttxtWristPleat2, "Second Wrist Pleat") Then
            Exit Sub
        End If
    End Sub

    Private Sub RighttxtShoulderPleat1_Leave(sender As Object, e As EventArgs) Handles RighttxtShoulderPleat1.Leave
        If Not FN_CheckValue(RighttxtShoulderPleat1, "First Shoulder Pleat") Then
            Exit Sub
        End If
    End Sub

    Private Sub RighttxtShoulderPleat2_Leave(sender As Object, e As EventArgs) Handles RighttxtShoulderPleat2.Leave
        If Not FN_CheckValue(RighttxtShoulderPleat2, "Second Shoulder Pleat") Then
            Exit Sub
        End If
    End Sub
    Private Sub PR_DrawRuler(ByRef nTape As Object, ByRef xyStart As ARMDIA1.XY, ByRef sTape As String)
        Dim xyPt As ARMDIA1.XY
        Dim nSymbolOffSet As Double
        nSymbolOffSet = 0.6877
        Dim sSymbol As String
        sSymbol = nTape & "tape"
        Dim xyTapeFst, xyTapeSec, xyTapeEnd, xyTapeText As ARMDIA1.XY
        Dim strTape As String = sTape
        If strTape.Contains("-") Then
            Dim iPos As Integer = InStr(strTape, "-")
            strTape = VB.Right(strTape, strTape.Length - 1)
        End If

        ARMDIA1.PR_MakeXY(xyPt, 0, 0)
        ARMDIA1.PR_MakeXY(xyTapeFst, xyPt.X, xyPt.y - 0.05)
        ARMDIA1.PR_MakeXY(xyTapeSec, xyPt.X, xyPt.y - 0.5)
        ARMDIA1.PR_MakeXY(xyTapeEnd, xyPt.X, xyTapeSec.y + 0.05)
        ARMDIA1.PR_MakeXY(xyTapeText, xyPt.X, xyPt.y - 0.25)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim strLayout As String = BlockTableRecord.ModelSpace
        If fnGetSettingsPath("DRAW_IN_LAYOUT").ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase) Then
            strLayout = BlockTableRecord.PaperSpace
        End If
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim blkRecId As ObjectId = ObjectId.Null
            If Not acBlkTbl.Has(sSymbol) Then
                Dim blkTblRecTape As BlockTableRecord = New BlockTableRecord()
                blkTblRecTape.Name = sSymbol

                Dim acLine As Line = New Line(New Point3d(xyPt.X, xyPt.y, 0), New Point3d(xyTapeFst.X, xyTapeFst.y, 0))
                blkTblRecTape.AppendEntity(acLine)

                acLine = New Line(New Point3d(xyTapeSec.X, xyTapeSec.y, 0), New Point3d(xyTapeEnd.X, xyTapeEnd.y, 0))
                blkTblRecTape.AppendEntity(acLine)

                '' Create a single-line text object
                Dim acText As DBText = New DBText()
                acText.Position = New Point3d(xyTapeText.X, xyTapeText.y, 0)
                acText.Height = 0.125
                acText.TextString = strTape
                acText.Rotation = 0
                acText.Justify = AttachmentPoint.MiddleCenter
                acText.AlignmentPoint = New Point3d(xyTapeText.X, xyTapeText.y, 0)
                blkTblRecTape.AppendEntity(acText)

                acBlkTbl.UpgradeOpen()
                acBlkTbl.Add(blkTblRecTape)
                acTrans.AddNewlyCreatedDBObject(blkTblRecTape, True)
                blkRecId = blkTblRecTape.Id
            Else
                blkRecId = acBlkTbl(sSymbol)
            End If
            ' Insert the block into the current space
            If blkRecId <> ObjectId.Null Then
                Dim blkTblRec As BlockTableRecord = acTrans.GetObject(blkRecId, OpenMode.ForWrite)
                For Each objID As ObjectId In blkTblRec
                    Dim dbObj As Entity = acTrans.GetObject(objID, OpenMode.ForWrite)
                    dbObj.Layer = "Construct"
                Next
                'Create new block reference 
                Dim blkRef As BlockReference = New BlockReference(New Point3d(xyStart.X + xyInsertion.X, xyStart.y + nSymbolOffSet + xyInsertion.y, 0), blkRecId)
                If fnGetSettingsPath("UNITS").ToUpper().Equals("CM", StringComparison.InvariantCultureIgnoreCase) Then
                    blkRef.TransformBy(Matrix3d.Scaling(2.54, New Point3d(xyInsertion.X, xyInsertion.y, 0)))
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
    Function fnDisplayStrapToCM(ByRef nInches As Double) As String
        Dim nCMVal, nDec As Double
        Dim iInt As Short
        nCMVal = nInches * 2.54
        nCMVal = System.Math.Abs(nCMVal)
        iInt = Int(nCMVal)
        nDec = nCMVal - iInt
        If nDec >= 0.5 Then
            fnDisplayStrapToCM = Str(ARMDIA1.round(nCMVal))
            Exit Function
        End If
        If iInt <> 0 Then
            nDec = (nCMVal - iInt) * 10
            nDec = ARMDIA1.round(nDec) / 10
            nCMVal = iInt + nDec
            fnDisplayStrapToCM = Str(nCMVal)
        Else
            fnDisplayStrapToCM = Str(nCMVal)
        End If
    End Function
    Function fnDisplayToCM(ByRef nInches As Double) As String
        Dim nCMVal, nDec As Double
        Dim iInt As Short
        iInt = Int(nInches)
        nDec = nInches - iInt
        Dim iEighths As Double = 0
        Dim nPrecision As Double = 0.125
        If nDec <> 0 Then 'Avoid overflow
            iEighths = Int(nDec / nPrecision)
        Else
            iEighths = 0
        End If
        If iEighths <> 0 Then
            Select Case iEighths
                Case 2, 6
                    Dim nVal As Double = iEighths / 2
                    iEighths = nVal / 4
                    ''Commented for #169 in the issue list
                    'Case 4
                    '    iEighths = iEighths + 0.5
                Case Else
                    iEighths = iEighths / 8
            End Select
        End If

        nCMVal = (iInt + iEighths) * 2.54
        nCMVal = System.Math.Abs(nCMVal)
        iInt = Int(nCMVal)
        If iInt <> 0 Then
            nDec = (nCMVal - iInt) * 10
            nDec = ARMDIA1.round(nDec) / 10
            nCMVal = iInt + nDec
            fnDisplayToCM = Str(nCMVal)
        Else
            fnDisplayToCM = Str(1)
        End If
    End Function
    Private Sub PR_GetRightStyle(ByRef sStyle As String)
        Dim sDistalStyle, sProximalStyle As String
        sDistalStyle = "XX"
        If ARMDIA1.g_iRightStyleFirstTape = ARMDIA1.g_iRightFirstTape Then
            'Plain
            sDistalStyle = "PL"
        Else
            'Start at tape
            If ARMDIA1.g_iStyleFirstTape < 10 Then
                sDistalStyle = "0" & LTrim(Str(ARMDIA1.g_iStyleFirstTape))
            Else
                sDistalStyle = LTrim(Str(ARMDIA1.g_iStyleFirstTape))
            End If
        End If
        If RightchkGauntlets.Checked = True Then
            'Gauntlet
            sDistalStyle = "GT"
        End If
        If RightchkStump.Checked = True Then
            'Stump
            sDistalStyle = "ST"
        End If

        'Proximal
        sProximalStyle = "XX"
        If ARMDIA1.g_iRightStyleLastTape = ARMDIA1.g_iRightLastTape Then
            'Plain
            sProximalStyle = "PL"
        Else
            'Start at tape
            If ARMDIA1.g_iRightStyleLastTape < 10 Then
                sProximalStyle = "0" & LTrim(Str(ARMDIA1.g_iRightStyleLastTape))
            Else
                sProximalStyle = LTrim(Str(ARMDIA1.g_iRightStyleLastTape))
            End If
        End If
        If RightchkDetachable.Checked = True And ARMDIA1.g_iRightStyleLastTape <= 6 Then
            'Detachable Gauntlet only, if drawn from or less than tape +1-1/2
            sProximalStyle = "GT"
        End If
        If Right_optProximalTape_1.Checked = True Then
            'Flap
            sProximalStyle = "FP"
        End If
        If VB.Left(RightcboFlaps.Text, 4) = "Vest" Then
            'Vest raglan
            sProximalStyle = "VR"
        End If
        If VB.Left(RightcboFlaps.Text, 4) = "Body" Then
            'Bodysuit raglan
            sProximalStyle = "BR"
        End If

        'Join distal and proximal styles
        sStyle = sDistalStyle & sProximalStyle
    End Sub

    Private Sub RightLength_Enter(sender As Object, e As EventArgs) Handles RightLength.Enter
        Dim Index As Short = RightLength.GetIndex(sender)
        select_text_in_Box1(RightLength(Index))
        ''RightLength(Index).Focus()
        RightLength(Index).SelectAll()
    End Sub
End Class


Public Class SurroundingClass
    Public dictN As String = "SNInfo"
    Public dictK As String = "SNDIC"
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
                Dim XRecord As Xrecord = DirectCast(tr.GetObject(dict.GetAt(key), OpenMode.ForWrite), Xrecord)
                XRecord.Erase()
                dict.Remove(key)
                tr.Commit()
            End Using
        End Using
    End Sub


End Class