Attribute VB_Name = "BODYDRAW1"
    Option Explicit

    Global MainForm As bodydia

    Global g_nFoldHt As Double
    Global g_nFoldHtRevised As Double
 
    Global g_sPreviousLegStyle As String
    Global sTask As String

'Globals set by FN_Open
    Global CC   'Comma
    Global QQ   'Quote
    Global NL   'Newline
    Global fNum 'Macro file number
    Global QCQ  'Quote Comma Quote
    Global QC   'Quote Comma
    Global CQ   'Comma Quote

'Store current layer and text setings to reduce DRAFIX code
'this value is checked in PR_SetLayer
    Global g_sCurrentLayer As String
    Global g_nCurrTextHt
    Global g_nCurrTextAspect
    Global g_nCurrTextHorizJust
    Global g_nCurrTextVertJust
    Global g_nCurrTextFont
    Global g_nCurrTextAngle

    Global g_LeftLegProfile     As curve
    Global g_RightLegProfile    As curve
    
    Global g_sChangeChecker As String
    Global g_sPathJOBST     As String


    
'PI as a Global Constant
    Global Const PI = 3.141592654


'Used to check if any modifications have been
'done on the data
    Global g_Modified As Integer

    'Global dimension variables

