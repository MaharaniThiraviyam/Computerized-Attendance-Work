VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FF8080&
   Caption         =   "Form5"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
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
      Left            =   12360
      TabIndex        =   3
      Top             =   3480
      Width           =   4095
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
      Height          =   975
      Left            =   7920
      TabIndex        =   2
      Top             =   7920
      Width           =   3495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   2895
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
      Left            =   4680
      TabIndex        =   1
      Text            =   "Month"
      Top             =   5640
      Width           =   2895
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
      Left            =   11280
      TabIndex        =   0
      Text            =   "Year"
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "  ROLL NUMBER"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   3600
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   " NUMBER OF PRESENT         DAYS OF A STUDENT"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   6840
      TabIndex        =   4
      Top             =   1920
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   6720
      Picture         =   "form5.frx":0000
      Top             =   360
      Width           =   5640
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S As String
Dim find As String
Dim mon As String
Dim c
Private Sub Command1_Click()
S = Text1.Text
mon = Combo1.Text
If mon = "dec" Or mon = "jan" Or mon = "feb" Or mon = "mar" Or mon = "apr" Then
semester = "sem-even"
find = "e-month"
Else
semester = "sem-odd"
find = "o-month"
End If
Data1.DatabaseName = "f:\attendance\" & find & ".xlsx"
Data1.RecordSource = "'" & Combo1.Text & "-" & Combo2.Text & "'" & "$"
Data1.Refresh
Data1.Recordset.MoveFirst
Do While Data1.Recordset.EOF <> True
c = c + 1
If ((IsNull(Data1.Recordset(1))) Or (Data1.Recordset(1) = "STUDENT NAME")) Then
Data1.Recordset.MoveNext
ElseIf S = Data1.Recordset(0) Then
days = Data1.Recordset(33)
MsgBox ("Student" & " " & Data1.Recordset(1) & " " & "has attended the college for" & " " & days & " " & "in the month of" & " " & Combo1.Text)
Exit Do
Else
Data1.Recordset.MoveNext
End If
Loop
End Sub
Private Sub Form_Load()
Combo1.AddItem "jan"
Combo1.AddItem "feb"
Combo1.AddItem "mar"
Combo1.AddItem "apr"
Combo1.AddItem "jun"
Combo1.AddItem "july"
Combo1.AddItem "aug"
Combo1.AddItem "sep"
Combo1.AddItem "oct"
Combo1.AddItem "dec"
Combo2.AddItem "I year"
Combo2.AddItem "II year"
Combo2.AddItem "III year"
End Sub





