VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Jumper - Menu"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   2775
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   4455
   End
   Begin VB.CommandButton cmdSuperHard 
      Caption         =   "Super-Hard"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4455
   End
   Begin VB.CommandButton cmdHard 
      Caption         =   "Hard"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton cmdMedium 
      Caption         =   "Medium"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton cmdEasy 
      Caption         =   "Easy"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Jumper v1.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Level As Integer
Public OpenForm As New Form1
Private Sub cmdEasy_Click()
Level = 1
OpenForm.Show
Form2.Hide
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdHard_Click()
Level = 3
OpenForm.Show
Form2.Hide
End Sub

Private Sub cmdMedium_Click()
Level = 2
OpenForm.Show
Form2.Hide
End Sub

Private Sub cmdSuperHard_Click()
Level = 4
OpenForm.Show
Form2.Hide
End Sub
