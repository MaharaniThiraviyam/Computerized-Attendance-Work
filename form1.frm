VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008080&
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
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
      Height          =   735
      Left            =   12120
      TabIndex        =   5
      Top             =   3000
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   1680
      TabIndex        =   4
      Text            =   "Date"
      Top             =   5280
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   5880
      TabIndex        =   3
      Text            =   "Month"
      Top             =   5280
      Width           =   2415
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   9840
      TabIndex        =   2
      Text            =   "Year"
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLICK ME  "
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
      Left            =   7920
      TabIndex        =   1
      Top             =   8160
      Width           =   2895
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   14160
      TabIndex        =   0
      Text            =   "Select"
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Data Data1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "      ROLL NO"
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
      Left            =   3720
      TabIndex        =   7
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "      ABSENT FORM"
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
      Left            =   7320
      TabIndex        =   6
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   6840
      Picture         =   "form1.frx":0000
      Top             =   240
      Width           =   5640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim R As String
Dim find As String
Dim mon As String
Private Sub Command1_Click()
mon = Combo2.Text
If mon = "dec" Or mon = "jan" Or mon = "feb" Or mon = "mar" Or mon = "apr" Then
semester = "sem-even"
find = "e-month"
Else
semester = "sem-odd"
find = "o-month"
End If
Data1.DatabaseName = "f:\attendance\" & find & ".xlsx"
Data1.RecordSource = "'" & Combo2.Text & "-" & Combo3.Text & "'" & "$"
Data1.Refresh
R = Text1.Text
d = Combo1.Text
d = d + 1
Data1.Recordset.MoveFirst
Do While Data1.Recordset.EOF <> True
If ((IsNull(Data1.Recordset(1))) Or (Data1.Recordset(1) = "STUDENT NAME")) Then
Data1.Recordset.MoveNext
End If
If Data1.Recordset(0) = R Then
If Combo4.Text = "ABSENT" Then
Data1.Recordset.Edit
Data1.Recordset(d) = "A"
Data1.Recordset.Update
Exit Do
End If
If Combo4.Text = "HALF DAY" Then
Data1.Recordset.Edit
Data1.Recordset(d) = "H"
Data1.Recordset.Update
Exit Do
End If
Else
Data1.Recordset.MoveNext
End If
Loop
MsgBox ("the process done successfully")
End Sub

Private Sub Form_Load()
For v = 1 To 31
Combo1.AddItem v
Next
Combo2.AddItem "jan"
Combo2.AddItem "feb"
Combo2.AddItem "mar"
Combo2.AddItem "apr"
Combo2.AddItem "jun"
Combo2.AddItem "july"
Combo2.AddItem "aug"
Combo2.AddItem "sep"
Combo2.AddItem "oct"
Combo2.AddItem "dec"
Combo3.AddItem "I year"
Combo3.AddItem "II year"
Combo3.AddItem "III year"
Combo4.AddItem "ABSENT"
Combo4.AddItem "HALF DAY"
End Sub




