VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H000080FF&
   Caption         =   "MAIN FORM"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Lucida Calligraphy"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form7"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT  "
      Height          =   855
      Left            =   8760
      TabIndex        =   1
      Top             =   8520
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   " COMPUTERIZED ATTENDENCE WORK"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   6360
      TabIndex        =   0
      Top             =   6360
      Width           =   7575
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   7440
      Picture         =   "Form7.frx":0000
      Top             =   4560
      Width           =   5640
   End
   Begin VB.Image Image1 
      Height          =   3300
      Left            =   5520
      Picture         =   "Form7.frx":0AB7
      Stretch         =   -1  'True
      Top             =   720
      Width           =   9570
   End
   Begin VB.Menu menu 
      Caption         =   "RUN"
      Begin VB.Menu menuform1 
         Caption         =   "Absent Form"
      End
      Begin VB.Menu menuform2 
         Caption         =   "Total Present Days-Month"
      End
      Begin VB.Menu menuform3 
         Caption         =   "OD Form"
      End
      Begin VB.Menu menuform4 
         Caption         =   "RC Form"
      End
      Begin VB.Menu menuform5 
         Caption         =   "Student Present-Month"
      End
      Begin VB.Menu menuform6 
         Caption         =   "Stuent Present-Sem"
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Command2_Click()
x = Text1.Text
If x = "123456" Then
MsgBox ("the entered password is correct")
Else
MsgBox ("re-enter the password")
End If

End Sub



Private Sub menuform1_Click()
Form1.Show
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
End Sub

Private Sub menuform2_Click()
Form2.Show
Form1.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
End Sub

Private Sub menuform3_Click()
Form3.Show
Form1.Hide
Form2.Hide
Form4.Hide
Form5.Hide
Form6.Hide
End Sub

Private Sub menuform4_Click()
Form4.Show
Form1.Hide
Form2.Hide
Form3.Hide
Form5.Hide
Form6.Hide
End Sub

Private Sub menuform5_Click()
Form5.Show
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form6.Hide
End Sub

Private Sub menuform6_Click()
Form6.Show
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
End Sub
