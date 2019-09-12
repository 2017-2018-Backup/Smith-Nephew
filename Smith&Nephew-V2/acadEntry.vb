Imports Autodesk.AutoCAD.Runtime

Public Class acadEntry
    <CommandMethod("NewPatient")>
    Public Sub NewPatient()
        Dim patient As New cltxfdia
        patient.ShowDialog()
    End Sub

    <CommandMethod("Arm")>
    Public Sub Arm()
        Dim arm As New armdia
        arm.ShowDialog()
    End Sub
    <CommandMethod("ManGlove")>
    Public Sub Glove()
        ''Dim Manglove As New manglove
        Dim Manglove As New MANGLOVEMAIN_frm
        ManGlvMain.ManGlvMainDlg = Manglove
        Manglove.ShowDialog()
    End Sub

    <CommandMethod("HeadNeck")>
    Public Sub HeadNeck()
        Dim HN As New HeadNeck
        HN.ShowDialog()
    End Sub
    <CommandMethod("Vest")>
    Public Sub vest()
        Dim VMain As New VESTMAIN_frm
        VestMain.VestMainDlg = VMain
        VMain.ShowDialog()
    End Sub
    <CommandMethod("Waist")>
    Public Sub Waist()
        Dim WHMain As New WHMAIN_frm
        WaistMain.WaistMainDlg = WHMain
        WHMain.ShowDialog()
    End Sub
    <CommandMethod("Leg")>
    Public Sub Leg()
        Dim LgMain As New LEGMAIN_frm
        LegMain.LegMainDlg = LgMain
        LgMain.ShowDialog()
    End Sub
    <CommandMethod("BodySuit")>
    Public Sub BodySuit()
        Dim BDMain As New BODYMAIN_frm
        BodyMain.BodyMainDlg = BDMain
        BDMain.ShowDialog()
    End Sub
End Class
