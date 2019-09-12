VERSION 5.00
Begin VB.Form EditUSER 
   Caption         =   "Edit USER Stamps"
   ClientHeight    =   5235
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   5160
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
   ScaleHeight     =   5235
   ScaleWidth      =   5160
   Begin VB.CommandButton Change 
      Caption         =   "Change"
      Height          =   315
      Left            =   3360
      TabIndex        =   6
      Top             =   3780
      Width           =   975
   End
   Begin VB.CommandButton Del 
      Caption         =   "Delete"
      Height          =   315
      Left            =   1980
      TabIndex        =   5
      Top             =   3780
      Width           =   975
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   315
      Left            =   600
      TabIndex        =   4
      Top             =   3780
      Width           =   975
   End
   Begin VB.ListBox StampList 
      Height          =   2595
      Left            =   180
      TabIndex        =   3
      Top             =   720
      Width           =   4755
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   4500
      Width           =   1095
   End
   Begin VB.TextBox txtStamp 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   4755
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   4500
      Width           =   1095
   End
End
Attribute VB_Name = "EditUSER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   'Index to stamp selected from StampList
    Dim nStampIndex As Integer

   'Handle to file number
    Dim hFileNumber


Private Sub Add_Click()
    StampList.AddItem txtStamp.Text
    'Deselect stamplist and disable buttons
    nStampIndex = -1
    StampList.ListIndex = nStampIndex
    Add.Enabled = False
    Del.Enabled = False
    Change.Enabled = False
End Sub

Private Sub Cancel_Click()
    Unload EditUSER
End Sub

Private Sub Change_Click()
    StampList.RemoveItem nStampIndex
    StampList.AddItem txtStamp.Text, nStampIndex
    'Deselect stamplist and disable buttons
    nStampIndex = -1
    StampList.ListIndex = nStampIndex
    Add.Enabled = False
    Change.Enabled = False
    Del.Enabled = False
End Sub

Private Sub Del_Click()
    StampList.RemoveItem nStampIndex
    txtStamp.Text = ""
    'Deselect stamplist and disable buttons
    nStampIndex = -1
    StampList.ListIndex = nStampIndex
    Add.Enabled = False
    Del.Enabled = False
    Change.Enabled = False
End Sub

Private Sub Form_Load()
    Dim sLine, sPathJOBST As String

    Left = (Screen.Width - EditUSER.Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - EditUSER.Height) / 2  ' Center form vertically.
    
   'Read USER stamps from file
    hFileNumber = FreeFile
    If Dir$("C:\JOBST\STAMPS.USR") <> "" Then
        Open "C:\JOBST\STAMPS.USR" For Input As hFileNumber
        Do While Not EOF(hFileNumber)
            Line Input #hFileNumber, sLine
            StampList.AddItem sLine
            Loop
        Close hFileNumber
    End If

   'Disable Add, Delete and Change buttons
    Add.Enabled = False
    Del.Enabled = False
    Change.Enabled = False
    nStampIndex = -1
   
End Sub

Private Sub OK_Click()
    Dim I

   'Write revised USER stamps to file
    Open "C:\JOBST\STAMPS.USR" For Output As hFileNumber
    For I = 0 To (StampList.ListCount - 1)
          Print #hFileNumber, StampList.List(I)
    Next I
    Close hFileNumber

   'Return to main form
    Unload EditUSER

End Sub

Private Sub StampList_Click()
    nStampIndex = StampList.ListIndex
    txtStamp.Text = StampList.List(nStampIndex)
    Del.Enabled = True
    Add.Enabled = False
    Change.Enabled = False
End Sub

Private Sub txtStamp_Change()
    Add.Enabled = True
    If nStampIndex >= 0 Then
        Change.Enabled = True
    End If
End Sub

