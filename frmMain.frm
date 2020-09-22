VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "15 game!"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "use arrows to move pieces."
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label LblQuit 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      TabIndex        =   2
      Top             =   4200
      Width           =   1515
   End
   Begin VB.Label lblScramble 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scramble"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      TabIndex        =   1
      Top             =   4200
      Width           =   1515
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   435
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   435
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'If you need something, contact me on
'Email  : a_ster0id@yahoo.com
'or ICQ : 15039767
'no stupid questions please

Private Type Pos
x As Integer
y As Integer
End Type

Private Const XSIZE = 3 'try to change it to 4 or 5
Private Const YSIZE = 3 'try to change it to 4 or 5

Dim BoardIndexes(XSIZE, YSIZE) As Integer

Dim Missing As Pos
Dim StartPos As Pos

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 37
    If Missing.x < XSIZE Then
        Call MoveToMissing(lblMain(BoardIndexes(Missing.x + 1, Missing.y)), Missing.x + 1, Missing.y)
    End If
Case 38
    If Missing.y < YSIZE Then
        Call MoveToMissing(lblMain(BoardIndexes(Missing.x, Missing.y + 1)), Missing.x, Missing.y + 1)
    End If
Case 39
    If Missing.x > 0 Then
        Call MoveToMissing(lblMain(BoardIndexes(Missing.x - 1, Missing.y)), Missing.x - 1, Missing.y)
    End If
Case 40
    If Missing.y > 0 Then
        Call MoveToMissing(lblMain(BoardIndexes(Missing.x, Missing.y - 1)), Missing.x, Missing.y - 1)
    End If
End Select

End Sub

Public Sub MoveToMissing(vLbl As Label, x As Integer, y As Integer)
Dim TmpMissingX As Integer
Dim TmpMissingY As Integer

    TmpMissingX = Missing.x
    TmpMissingY = Missing.y
    
    Missing.x = x
    Missing.y = y
   
    vLbl.Left = StartPos.x + ((vLbl.Width + 100) * TmpMissingX)
    vLbl.Top = StartPos.y + ((vLbl.Height + 100) * TmpMissingY)
    
    BoardIndexes(x, y) = -1
    BoardIndexes(TmpMissingX, TmpMissingY) = vLbl.Index
End Sub

Public Sub SetPicturePos(vLbl As Label, x As Integer, y As Integer)
    vLbl.Left = StartPos.x + (vLbl.Width * x)
    vLbl.Top = StartPos.y + (vLbl.Height * y)
End Sub

Private Sub Form_Load()
StartPos.x = lblMain(0).Left
StartPos.y = lblMain(0).Top

BoardIndexes(0, 0) = 0
Dim TmpCounter As Integer
TmpCounter = 0
For i = 0 To YSIZE
    For j = 0 To XSIZE
        If (i + j > 0) And (i + j < XSIZE + YSIZE) Then
            TmpCounter = TmpCounter + 1
            Load lblMain(TmpCounter)
            lblMain(TmpCounter).Caption = TmpCounter + 1
            lblMain(TmpCounter).Visible = True
            lblMain(TmpCounter).Left = StartPos.x + ((lblMain(TmpCounter).Width + 100) * j)
            lblMain(TmpCounter).Top = StartPos.y + ((lblMain(TmpCounter).Height + 100) * i)
            BoardIndexes(j, i) = TmpCounter
        End If
    Next
Next
Missing.x = XSIZE
Missing.y = YSIZE
End Sub

Private Sub lblMain_Click(Index As Integer)
'MsgBox Index
End Sub

Private Sub LblQuit_Click()
End
End Sub

Private Sub lblScramble_Click()
Randomize
Dim Tmp As Double

For i = 0 To 300 '300 random moves
    Tmp = Rnd
    If Tmp < (1 / 4) Then
        Call Form_KeyDown(37, 0)
    ElseIf Tmp < (2 / 4) Then
        Call Form_KeyDown(38, 0)
    ElseIf Tmp < (3 / 4) Then
        Call Form_KeyDown(39, 0)
    Else
        Call Form_KeyDown(40, 0)
    End If
Next
End Sub
