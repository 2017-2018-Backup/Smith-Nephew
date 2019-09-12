Attribute VB_Name = "HEADNECK1"
    'XY data type to represent points
    Public MainForm As HeadNeck
        
    Type xy
        X As Double
        Y As Double
    End Type
    
    Type curve
        X(1 To 100)   As Double
        Y(1 To 100)   As Double
    End Type

   'Globals set by FN_Open
    Global CC   'Comma
    Global QQ   'Quote
    Global NL   'Newline
    Global QCQ  'Quote Comma Quote
    Global QC   'Quote Comma
    Global CQ   'Comma Quote

    Global g_sWorkOrder As String
    Global g_sPathJOBST As String

    Global Const PI = 3.141593

