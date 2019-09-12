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

    <CommandMethod("HeadNeck")>
    Public Sub HeadNeck()
        Dim HN As New HeadNeck
        HN.ShowDialog()
    End Sub

    '<CommandMethod("BodySuit")>
    'Public Sub HeadNeck()
    '    Dim HN As New HeadNeck
    '    HN.ShowDialog()
    'End Sub

    '<CommandMethod("HN")>
    'Public Sub HeadNeck()
    '    Dim HN As New HeadNeck
    '    HN.ShowDialog()
    'End Sub

    '<CommandMethod("HN")>
    'Public Sub HeadNeck()
    '    Dim HN As New HeadNeck
    '    HN.ShowDialog()
    'End Sub

End Class
