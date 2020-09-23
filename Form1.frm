VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   Caption         =   "Jumper"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   495
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Left            =   6720
      Top             =   3720
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   5760
      Top             =   4200
   End
   Begin VB.Timer Timer2 
      Left            =   6240
      Top             =   4200
   End
   Begin VB.Timer Timer1 
      Left            =   6720
      Top             =   4200
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Shape PlayerBox 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   480
      Left            =   3465
      Top             =   4455
      Width           =   495
   End
   Begin VB.Shape Box 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   1
      Left            =   6840
      Top             =   495
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Boxes As Integer
Public PlayerX As Integer
Public PlayerY As Integer

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
If PlayerY > -1 Then
PlayerBox.Top = PlayerBox.Top - 33
End If
End If
If KeyCode = 37 Then
If Not PlayerX < 1 Then
PlayerBox.Left = PlayerBox.Left - 33
End If
End If
If KeyCode = 39 Then
If Not PlayerX > 13 Then
PlayerBox.Left = PlayerBox.Left + 33
End If
End If
If KeyCode = 40 Then
If Not PlayerY > 9 Then
PlayerBox.Top = PlayerBox.Top + 33
End If
End If
End Sub

Private Sub Timer1_Timer()
MakeBox
End Sub

Private Sub Form_Load()
  Randomize
  If Form2.Level = 1 Then
  Timer4.Interval = 1000
  Timer1.Interval = 1000
  Timer2.Interval = 50
  End If
  If Form2.Level = 2 Then
  Timer4.Interval = 700
  Timer1.Interval = 700
  Timer2.Interval = 40
  End If
  If Form2.Level = 3 Then
  Timer4.Interval = 500
  Timer1.Interval = 500
  Timer2.Interval = 30
  End If
  If Form2.Level = 4 Then
  Timer4.Interval = 250
  Timer1.Interval = 250
  Timer2.Interval = 15
  End If
Boxes = 1
End Sub
Function MakeBox()
High = 4
Low = 1
rand = Int((High - Low + 1) * Rnd) + Low
If rand = 1 Then
Boxes = Boxes + 1
Load Box(Boxes)
Box(Boxes).Top = 33
Box(Boxes).Left = Form1.ScaleWidth
Box(Boxes).Visible = True
End If
If rand = 2 Then
Boxes = Boxes + 1
Load Box(Boxes)
Box(Boxes).Top = 99
Box(Boxes).Left = Form1.ScaleWidth
Box(Boxes).Visible = True
End If
If rand = 3 Then
Boxes = Boxes + 1
Load Box(Boxes)
Box(Boxes).Top = 165
Box(Boxes).Left = Form1.ScaleWidth
Box(Boxes).Visible = True
End If
If rand = 4 Then
Boxes = Boxes + 1
Load Box(Boxes)
Box(Boxes).Top = 231
Box(Boxes).Left = Form1.ScaleWidth
Box(Boxes).Visible = True
End If
End Function
Function MakeBox2()
High = 5
Low = 1
rand = Int((High - Low + 1) * Rnd) + Low
If rand = 1 Then
Boxes = Boxes + 1
Load Box(Boxes)
Box(Boxes).Top = 0
Box(Boxes).Left = Form1.ScaleWidth
Box(Boxes).Visible = True
End If
If rand = 2 Then
Boxes = Boxes + 1
Load Box(Boxes)
Box(Boxes).Top = 66
Box(Boxes).Left = Form1.ScaleWidth
Box(Boxes).Visible = True
End If
If rand = 3 Then
Boxes = Boxes + 1
Load Box(Boxes)
Box(Boxes).Top = 132
Box(Boxes).Left = Form1.ScaleWidth
Box(Boxes).Visible = True
End If
If rand = 4 Then
Boxes = Boxes + 1
Load Box(Boxes)
Box(Boxes).Top = 198
Box(Boxes).Left = Form1.ScaleWidth
Box(Boxes).Visible = True
End If
If rand = 5 Then
Boxes = Boxes + 1
Load Box(Boxes)
Box(Boxes).Top = 264
Box(Boxes).Left = Form1.ScaleWidth
Box(Boxes).Visible = True
End If
End Function

Private Sub Timer2_Timer()
For i = Box.LBound To Box.UBound
Box(i).Left = Box(i).Left - 8
Next i
End Sub

Private Sub Timer3_Timer()
If PlayerY = -1 Then
If Form2.Level = 1 Then
MsgBox "You completed the level, Mind you - Even a spastic could do this level on Easy"
PlayerBox.Left = 231
PlayerBox.Top = 297
End If
If Form2.Level = 2 Then
MsgBox "You did the level on medium... Now try hard :)"
PlayerBox.Left = 231
PlayerBox.Top = 297
End If
If Form2.Level = 3 Then
MsgBox "Was the hard level not hard? Hmm... Try the super hard level"
PlayerBox.Left = 231
PlayerBox.Top = 297
End If
If Form2.Level = 4 Then
MsgBox "Damn, I guess I should of made a SuperSuperSuperSuper hard level for people like you :|"
PlayerBox.Left = 231
PlayerBox.Top = 297
End If
Form2.OpenForm.Hide
Form2.Show
End If
PlayerX = PlayerBox.Left / 33
PlayerY = PlayerBox.Top / 33
Label2.Caption = "X: " & PlayerX & "   Y: " & PlayerY
For i = Box.LBound To Box.UBound
If Box(i).Left <= PlayerBox.Left + PlayerBox.Width Then
    If Box(i).Left >= PlayerBox.Left Then
        If Box(i).Top <= PlayerBox.Top + PlayerBox.Height Then
            If Box(i).Top >= PlayerBox.Top Then
                PlayerBox.Top = 297
                PlayerBox.Left = 231
            End If
        End If
    End If
End If
Next i
End Sub

Private Sub Timer4_Timer()
MakeBox2
End Sub
