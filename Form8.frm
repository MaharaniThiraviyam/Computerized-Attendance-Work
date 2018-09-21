VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6000
      TabIndex        =   2
      Top             =   0
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLIKE ME     "
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   4440
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   3135
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do
MsgBox ("sorry!!! re-enter the password")
Text1.SetFocus
Loop Until Text2.Text = "123456"
Form7.Show
End Sub

