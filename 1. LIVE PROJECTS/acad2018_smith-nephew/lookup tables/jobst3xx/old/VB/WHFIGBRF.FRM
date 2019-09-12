VERSION 5.00
Begin VB.Form whfigbrf 
   Caption         =   "Waist Height - Brief"
   ClientHeight    =   2325
   ClientLeft      =   1095
   ClientTop       =   1500
   ClientWidth     =   3780
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2325
   ScaleWidth      =   3780
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1980
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   420
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Brief Fabric"
      Height          =   1395
      Left            =   180
      TabIndex        =   1
      Top             =   150
      Width           =   3435
      Begin VB.ComboBox cboFabric 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   330
         Width           =   2475
      End
      Begin VB.OptionButton optFabric 
         Caption         =   "Powernet"
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optFabric 
         Caption         =   "Jobstex"
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   2
         Top             =   870
         Width           =   1155
      End
      Begin VB.Label Label27 
         Caption         =   "Fabric:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   390
         Width           =   1035
      End
   End
End
Attribute VB_Name = "whfigbrf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project:   WHFIGURE.MAK
'Form:      WHFIGBRF.FRM
'Purpose:   Restricted dialog for Waist Height Brief fabric
'           selection
'
'Version:   1.01
'Date:      16.Jan.94
'Author:    Gary George
'-------------------------------------------------------
'REVISIONS:
'Date       By      Action
'-------------------------------------------------------
'
'Notes:-
'
'   As this is a later addition to WHFIGURE it does not
'   use the same method as the main form to transfer data
'   to the drawing through a DDE link.
'
'   It uses the method of creating a macro with all the data
'   in place which is then executed by drafix.

    Option Explicit

Private Sub Cancel_Click()
    End
End Sub

Private Sub Form_Load()
    Dim ii As Integer
   'Position to center of screen
    Left = (Screen.Width - whfigbrf.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - whfigbrf.Height) / 2  ' Center form vertically.
    
    Screen.MousePointer = 0    ' Change pointer to default.
    
   'Set up for fabric if fabric is given
    If g_sLastFabric <> "" Then
        If UCase$(Left$(g_sLastFabric, 3)) = "POW" Then
            optFabric(0).Value = 1
            optFabric_Click (0)
        Else
            optFabric(1).Value = 1
            optFabric_Click (1)
        End If
        For ii = 0 To cboFabric.ListCount - 1
            If g_sLastFabric = cboFabric.List(ii) Then
                cboFabric.ListIndex = ii
                Exit For
            End If
        Next ii
    Else
       'Use POWERNET for burns
        If g_sDiagnosis = "Burns" Then
            optFabric(0).Value = 1
            optFabric_Click (0)
        Else
            optFabric(1).Value = 1
            optFabric_Click (1)
        End If
    End If
End Sub

Private Sub OK_Click()
    Dim sFabric As String

   'Check if afabric has been given
    sFabric = cboFabric.List(cboFabric.ListIndex)
    If sFabric = "" Then
        Beep
        Exit Sub
    End If
   
   'If a fabric has beeb given then write macro file to update drawing
   'Open the DRAFIX macro file
   'Initialise Global variables

   'Open file
    fNum = FreeFile

   'Initialise String variables
    CC = Chr$(44)       'The comma ( , )
    NL = Chr$(10)       'The new line character
    QQ = Chr$(34)       'Double quotes ( " )
    QCQ = QQ & CC & QQ  'Quote Comma Quote ( "," )
    QC = QQ & CC        'Quote Comma ( ", )
    CQ = CC & QQ        'Comma Quote ( ," )

   'Write header information etc. to the DRAFIX macro file
   '
    Open "C:\JOBST\DRAW.D" For Output As fNum

        Print #fNum, "//DRAFIX Figuring Macro created - "; Date$; "  "; Time$
        Print #fNum, "//Updates waist box with BRIEF Fabric"
        Print #fNum, "//by Visual Basic"
        Print #fNum, "HANDLE  hBody;"
        Print #fNum, "hBody = UID ("; QQ; "find"; QC; g_iUidBody; ");"
        Print #fNum, "if (!hBody)Exit(%cancel,"; QQ; "Can't find WAIST BOX to Update"; QQ; ");"

        Print #fNum, "SetDBData( hBody, "; QQ; "Fabric"; QCQ; sFabric; QQ; ");"
        
    Close #fNum

    AppActivate fnGetDrafixWindowTitleText()
    SendKeys "@C:\JOBST\DRAW.D{enter}", True
    End
End Sub

Private Sub optFabric_Click(index As Integer)

    cboFabric.Clear
    Select Case index
    Case 0
        g_POWERNET = True
        PR_GetComboListFromFile cboFabric, g_sPathJOBST + "\WHFABRIC.DAT"
    Case 1
        g_JOBSTEX = True
        cboFabric.AddItem "53 - JOBSTEX"
        cboFabric.AddItem "55 - JOBSTEX"
        cboFabric.AddItem "57 - JOBSTEX"
        cboFabric.AddItem "63 - JOBSTEX"
        cboFabric.AddItem "65 - JOBSTEX"
        cboFabric.AddItem "67 - JOBSTEX"
        cboFabric.AddItem "73 - JOBSTEX"
        cboFabric.AddItem "75 - JOBSTEX"
        cboFabric.AddItem "77 - JOBSTEX"
        cboFabric.AddItem "83 - JOBSTEX"
        cboFabric.AddItem "85 - JOBSTEX"
        cboFabric.AddItem "87 - JOBSTEX"
    End Select

End Sub

