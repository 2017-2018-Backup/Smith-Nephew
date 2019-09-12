VERSION 5.00
Begin VB.Form meshdraw 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "MeshDraw"
   ClientHeight    =   1605
   ClientLeft      =   4620
   ClientTop       =   2475
   ClientWidth     =   3165
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
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1605
   ScaleWidth      =   3165
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   720
   End
End
Attribute VB_Name = "meshdraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Module:    MESHVEST.MAK
'Purpose:   Mesh axilla drawing for vest and vest sleeves
'
'
'Version:   3.00
'Date:      12.Jan.98
'Author:    Gary George
'           © C-Gem Ltd.
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'Dec 98     GG      Ported to VB5
'Notes:-
'
'
'

    Option Explicit

Private Sub Form_Load()

    Dim sFileName       As String
    Dim sTask           As String
    Dim fFileNum        As Integer
    Dim iError          As Integer
    Dim iHandleBody     As Integer


    Dim xyMeshStart     As XY
    Dim xyTmp           As XY
    Dim xyProfileStart  As XY
    Dim xySeamStart     As XY
    Dim xyRaglanStart   As XY
    Dim xyRaglanEnd     As XY

    Dim nMeshLength             As Double
    Dim nDistanceAlongRaglan    As Double
    
    Dim Meshprofile     As BiArc
    Dim MeshSeam        As BiArc

    Dim nLength         As Double
    Dim aAngle          As Double

    Dim sMeshLeft       As String
    Dim sMeshRight      As String
        
   'Setup error handling
    On Error GoTo ErrorStarting

   'Hide form
    Hide

   'Check if a previous instance is running
   'If it is exit
    If App.PrevInstance Then End

   'Start a timer
   'This ensures that the dialogue dies in event of a failure
   'on the drafix macro side
    Timer1.Interval = 6000  'Approx 6 Seconds
    Timer1.Enabled = True

   'Find the path to the jobst system
    g_sPathJOBST = fnPathJOBST()

    Set MainForm = meshdraw
    
   'Open Data file
   'Note that this is a fixed file name
    sFileName = "C:\JOBST\MESHDRAW.DAT"
    If FileLen(sFileName) = 0 Then
        MsgBox "Can't open the file " & sFileName & ".  Unable to draw the mesh.", 48
        End
    End If

   fNum = FreeFile
   PR_Open "C:\JOBST\DRAW_1.D", fNum

   fFileNum = FreeFile
   Open sFileName For Input Access Read As fFileNum
    Input #fFileNum, g_sCallingApplication
    If g_sCallingApplication = "vestmesh" Then
         Input #fFileNum, iHandleBody
    End If
   
   Do While Not EOF(fFileNum)
    Input #fFileNum, xyMeshStart.X, xyMeshStart.Y
    Input #fFileNum, xyAxilla.X, xyAxilla.Y
    Input #fFileNum, xyRaglanStart.X, xyRaglanStart.Y
    Input #fFileNum, xyRaglanEnd.X, xyRaglanEnd.Y
    Input #fFileNum, nMeshLength
    Input #fFileNum, nDistanceAlongRaglan
    Input #fFileNum, g_sSide
    Line Input #fFileNum, g_sID
    
   'Offset xyMeshStart 1/8" from line (xyMeshStart,xyAxilla)
   'to allow for seam
    aAngle = FN_CalcAngle(xyMeshStart, xyAxilla) - 90
    PR_CalcPolar xyMeshStart, aAngle, 0.125, xyProfileStart
'    PR_CalcPolar xyMeshStart, aAngle, .125, xyProfileStart
'nMeshLength = nMeshLength + .125
'xyProfileStart = xyMeshStart

     PR_SetLayer "Construct"
     PR_DrawMarker xyProfileStart

    'Calculate mesh
    'note that if we fail the first time on a sleeve then we try again
    'using iError = FN_CalcAxillaMeshVest(1, ...
    '
     iError = FN_CalcAxillaMeshVest(0, xyRaglanStart, xyRaglanEnd, xyProfileStart, False, nMeshLength, nDistanceAlongRaglan, MeshSeam, Meshprofile)
     If Not iError And g_sCallingApplication = "sleevemesh" Then iError = FN_CalcAxillaMeshVest(1, xyRaglanStart, xyRaglanEnd, xyProfileStart, False, nMeshLength, nDistanceAlongRaglan, MeshSeam, Meshprofile)
     If iError = True Then
       'Draw the mesh
            PR_SetLayer "Template" & g_sSide
            PR_DrawArc Meshprofile.xyR1, Meshprofile.xyStart, Meshprofile.xyTangent
            PR_DrawArc Meshprofile.xyR2, Meshprofile.xyTangent, Meshprofile.xyEnd
            iError = FN_CirLinInt(xyMeshStart, xyAxilla, Meshprofile.xyR1, Meshprofile.nR1, xyTmp)
            PR_DrawLine Meshprofile.xyStart, xyTmp

           'Draw closing line
            PR_DrawLine xyRaglanStart, xyAxilla
            
            PR_SetLayer "Notes"
            PR_DrawArc MeshSeam.xyR1, MeshSeam.xyStart, MeshSeam.xyTangent
            PR_DrawArc MeshSeam.xyR2, MeshSeam.xyTangent, MeshSeam.xyEnd
            PR_CalcPolar xyMeshStart, FN_CalcAngle(xyAxilla, xyMeshStart), 3, xyTmp
            iError = FN_CirLinInt(xyTmp, xyAxilla, MeshSeam.xyR1, MeshSeam.nR1, xyTmp)
            PR_DrawLine MeshSeam.xyStart, xyTmp
            PR_SetLayer "Construct"
            PR_DrawMarker MeshSeam.xyEnd
     Else
        MsgBox "Failure to Calculate mesh. Draw manually and contact supervisor", 48
        PR_DrawMarker MeshSeam.xyEnd 'Bad practice here (using a side effect)
                                     'but I know that this point
                                     'is at the nDistanceAlongRaglan
     End If
    
    If g_sSide = "Left" Then
        sMeshLeft = Trim$(Str$(nDistanceAlongRaglan)) & "," & Trim$(Str$(nMeshLength))
    Else
        sMeshRight = Trim$(Str$(nDistanceAlongRaglan)) & "," & Trim$(Str$(nMeshLength))
    End If

   Loop
    
   Close #fFileNum
   PR_UpdateDB iHandleBody, sMeshLeft, sMeshRight

   Close #fNum
    
    sTask = fnGetDrafixWindowTitleText()
    If sTask <> "" Then
            AppActivate sTask
            SendKeys "@c:\jobst\draw_1.d{enter}", True
        Else
            MsgBox "Can't find a Drafix Drawing to update! - Mesh Drawing Terminating."
    End If

   'Terminate
    End

ErrorStarting:
    
    MsgBox "Error starting - Mesh Drawing Terminating."
    End

End Sub

Private Sub PR_UpdateDB(nHandleBody%, sMeshLeft$, sMeshRight$)
   
    Dim sSymbol As String, sMeshData  As String
    
    sSymbol = "vestbody"

   'Use existing symbol
    PR_PutLine "hBody = UID (" & QQ & "find" & QC & nHandleBody & ");"
    PR_PutLine "if (!hBody) Exit(%cancel," & QQ & "Can't find >" & sSymbol & "< symbol to update!" & QQ & ");"

   'Mesh data for use by sleeve drawing programme
   'We use commas as the delimiters to show empty data
    If sMeshLeft = "" And sMeshRight = "" Then
       sMeshData = ",,,"

    ElseIf sMeshLeft = "" Then
       sMeshData = ",," & sMeshRight

    ElseIf sMeshRight = "" Then
       sMeshData = sMeshLeft & ",,"

    Else
       sMeshData = sMeshLeft & "," & sMeshRight
    
    End If
    
    PR_PutLine "SetDBData( hBody" & CQ & "Data" & QCQ & sMeshData & QQ & ");"


End Sub

