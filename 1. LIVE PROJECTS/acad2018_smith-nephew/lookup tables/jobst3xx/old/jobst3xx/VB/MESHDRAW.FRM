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
'Module:    MESHDRAW.MAK
'Purpose:   Mesh axilla drawing
'
'Projects:  1. BODYDRAW.MAK
'           2. MESHDRAW.MAK
'
'Version:   1.00
'Date:      5.Nov.1997
'Author:    Gary George
'           © C-Gem Ltd.
'
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'Dec 98     GG      Ported to VB5
'
'Notes:-
'
' This module is designed to be retrofitted into the
' existing Vest code and as such it is stand alone
' with data being transfered by file
'
'
'

    Option Explicit

Private Sub Form_Load()

    Dim sFileName       As String
    Dim sTask           As String
    Dim fFileNum        As Integer
    Dim iError          As Integer

    Dim xyMeshStart     As XY
    Dim xyTmp           As XY
    Dim xyProfileStart  As XY
    Dim xySeamStart     As XY
    Dim xyRaglanStart   As XY
    Dim xyRaglanEnd     As XY
    Dim nMeshLength             As Double
    Dim nDistanceAlongRaglan    As Double
    
    Dim MeshProfile     As BiArc
    Dim MeshSeam        As BiArc

    Dim nLength         As Double
    Dim aAngle          As Double
        
   'Setup error handling
    On Error GoTo ErrorStarting

   'Hide form
    Hide

   'Check if a previous instance is running
   'If it is exit
    If App.PrevInstance Then End
    
    Set MainForm = meshdraw

   'Start a timer
   'This ensures that the dialogue dies in event of a failure
   'on the drafix macro side
    Timer1.Interval = 6000  'Approx 6 Seconds
    Timer1.Enabled = True

   'Find the path to the jobst system
    g_sPathJOBST = fnPathJOBST()

   'Open Data file
   'Note that this is a fixed file name
    sFileName = "C:\JOBST\MESHDRAW.DAT"
    fFileNum = FreeFile
    If FileLen(sFileName) = 0 Then
        MsgBox "Can't open the file " & sFileName & ".  Unable to draw the mesh.", 48
        End
    End If

   'Extract from BR_MESH.D
   'hCurve = Open ("file", "C:\\JOBST\\MESHDRAW.DAT", "write") ;
   ' PrintFile(hCurve, xyAxillaConstruct_2, "\n") ;
   ' PrintFile(hCurve, xyAxilla, "\n") ;
   ' PrintFile(hCurve, xyAxillaBodySuit, "\n") ;
   ' PrintFile(hCurve, xyBackNeck, "\n") ;
   ' PrintFile(hCurve, nMeshLength, "\n") ;
   ' PrintFile(hCurve, nDistanceAlongRaglan, "\n") ;
   ' PrintFile(hCurve, sSleeve, "\n") ;
   ' PrintFile(hCurve, sID, "\n") ;
   'Close ("file", hCurve) ;

    
    Open sFileName For Input Access Read As fFileNum
        Input #fFileNum, xyMeshStart.X, xyMeshStart.y
        Input #fFileNum, xyAxilla.X, xyAxilla.y
        Input #fFileNum, xyRaglanStart.X, xyRaglanStart.y
        Input #fFileNum, xyRaglanEnd.X, xyRaglanEnd.y
        Input #fFileNum, nMeshLength
        Input #fFileNum, nDistanceAlongRaglan
        Input #fFileNum, g_sSide
        Input #fFileNum, g_sID
    Close #fFileNum
    
   'Offset xyMeshStart 1/8" from line (xyMeshStart,xyAxilla)
   'to allow for seam
    aAngle = FN_CalCAngle(xyMeshStart, xyAxilla) - 90
    PR_CalcPolar xyMeshStart, aAngle, 0.125, xyProfileStart

    fNum = FN_Open("C:\JOBST\DRAW_1.D")
     PR_SetLayer "Construct"
     PR_DrawMarker xyProfileStart

    'Body suit
     iError = FN_CalcAxillaMesh(xyRaglanStart, xyRaglanEnd, xyProfileStart, False, nMeshLength, nDistanceAlongRaglan, MeshSeam, MeshProfile)
     If iError = True Then
       'Draw the mesh
            PR_SetLayer "Template" & g_sSide
            PR_DrawArc MeshProfile.xyR1, MeshProfile.xyStart, MeshProfile.xyTangent
            PR_DrawArc MeshProfile.xyR2, MeshProfile.xyTangent, MeshProfile.xyEnd
            iError = FN_CirLinInt(xyMeshStart, xyAxilla, MeshProfile.xyR1, MeshProfile.nR1, xyTmp)
            PR_DrawLine MeshProfile.xyStart, xyTmp

           'Draw closing line
            PR_DrawLine xyRaglanStart, xyAxilla
            
            PR_SetLayer "Notes"
            PR_DrawArc MeshSeam.xyR1, MeshSeam.xyStart, MeshSeam.xyTangent
            PR_DrawArc MeshSeam.xyR2, MeshSeam.xyTangent, MeshSeam.xyEnd
            PR_CalcPolar xyMeshStart, FN_CalCAngle(xyAxilla, xyMeshStart), 3, xyTmp
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

