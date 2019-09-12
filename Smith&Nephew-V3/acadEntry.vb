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
    <CommandMethod("SN_Edit")>
    Public Sub EditArm()
        Dim editArm As New armeddia
        editArm.ShowDialog()
    End Sub
    <CommandMethod("MIRPROF")>
    Public Sub MirrorProfile()
        Dim mirProf As New MIRBODYPROF
        mirProf.PR_GetBodyProfile()
    End Sub

    '<CommandMethod("HeelCtr")>
    'Public Sub HeelContracture()
    '    Dim mirProf As New MIRBODYPROF
    '    mirProf.PR_DrawHeelContracture()
    'End Sub

    <CommandMethod("KneeCtr")>
    Public Sub KneeContracture()
        Dim mirProf As New MIRBODYPROF
        mirProf.PR_DrawKneeContracture()
    End Sub

    <CommandMethod("ZipBod")>
    Public Sub ZipperBody()
        Dim mirProf As New MIRBODYPROF
        mirProf.PR_DrawWaistBodyZipper()
    End Sub

    <CommandMethod("EditZip")>
    Public Sub EditZipper()
        Dim mirProf As New MIRBODYPROF
        mirProf.PR_EditZipper()
    End Sub
    <CommandMethod("ZipAnk")>
    Public Sub AnkleZipper()
        Dim mirProf As New MIRBODYPROF
        mirProf.PR_DrawAnkleZipper()
    End Sub
    <CommandMethod("ZipLat")>
    Public Sub LateralZipper()
        Dim mirProf As New MIRBODYPROF
        mirProf.PR_DrawWaistLateralBodyZipper()
    End Sub
    <CommandMethod("ZipDistal")>
    Public Sub DistalZipper()
        Dim mirProf As New MIRBODYPROF
        mirProf.PR_DrawDistalZipper()
    End Sub
    <CommandMethod("WHZipAnk")>
    Public Sub WaistAnkleZipper()
        Dim mirProf As New MIRBODYPROF
        mirProf.PR_DrawWaistAnkleZipper()
    End Sub
    <CommandMethod("Chap")>
    Public Sub WaistChap()
        Dim mirProf As New MIRBODYPROF
        mirProf.PR_DrawWaistChap()
    End Sub
    <CommandMethod("ZipChpAnk")>
    Public Sub WaistChapAnkZip()
        Dim mirProf As New MIRBODYPROF
        mirProf.PR_DrawWaistChapAnkleZipper()
    End Sub
    <CommandMethod("ZipPalmer")>
    Public Sub GloveZipPalmer()
        Dim mirProf As New MIRBODYPROF
        mirProf.PR_DrawGlovePalmZipper()
    End Sub
    <CommandMethod("ZipDorsal")>
    Public Sub GloveZipDorsal()
        Dim mirProf As New MIRBODYPROF
        mirProf.PR_DrawGloveDorsalZipper()
    End Sub
    <CommandMethod("DrawBody")>
    Public Sub DrawBody()
        Dim Body As New bodysuit
        Body.PR_DrawBody()
    End Sub
    <CommandMethod("Cutout")>
    Public Sub DrawCutout()
        Dim WHBody As New whboddia
        WHBody.PR_DrawWaistCutout()
    End Sub
    <CommandMethod("FirstLeg")>
    Public Sub DrawFirstLeg()
        Dim WHLeg As New whlegdia
        WHLeg.PR_DrawWaistFirstLeg()
    End Sub
    <CommandMethod("SecondLeg")>
    Public Sub DrawSecondLeg()
        Dim WHRightLeg As New WHRightLeg
        WHRightLeg.PR_DrawWaistSecondLeg()
    End Sub
    <CommandMethod("ZipChpPanty")>
    Public Sub DrawChapPantyZipper()
        Dim mirrProf As New MIRBODYPROF
        mirrProf.PR_DrawWaistChapPantyZipper()
    End Sub
End Class
