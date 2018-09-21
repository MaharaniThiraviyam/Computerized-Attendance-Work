VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H008080FF&
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
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
      Left            =   8040
      TabIndex        =   3
      Top             =   7680
      Width           =   3855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   2895
   End
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
      Height          =   615
      Left            =   11880
      TabIndex        =   2
      Top             =   3360
      Width           =   3615
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
      Text            =   "Select"
      Top             =   5760
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
      Left            =   10800
      TabIndex        =   0
      Text            =   "Year"
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "NUMBER OF PRESENT DAYS"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   6480
      TabIndex        =   5
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Label Label2 
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
      Height          =   735
      Left            =   3240
      TabIndex        =   4
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   6840
      Picture         =   "form6.frx":0000
      Top             =   0
      Width           =   5640
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S As String
Private Sub Command1_Click()
S = Text1.Text
Data1.DatabaseName = "f:\attendance\" & Combo1.Text & ".xlsx"
Data1.RecordSource = "'" & Combo2.Text & "'" & "$"
Data1.Refresh
Data1.Recordset.MoveFirst
Do While Data1.Recordset.EOF <> True
If ((IsNull(Data1.Recordset(1))) Or (Data1.Recordset(1) = "STUDENT NAME")) Then
Data1.Recordset.MoveNext
ElseIf S = Data1.Recordset(0) Then
days = Data1.Recordset(7)
MsgBox "Student" & " " & Data1.Recordset(1) & " " & "has attended the college for" & " " & days & " " & "in the " & " " & Combo1.Text
Exit Do
Else
Data1.Recordset.MoveNext
End If
Loop
End Sub
Private Sub Form_Load()
Combo1.AddItem "sem-even"
Combo1.AddItem "sem-odd"
Combo2.AddItem "I year"
Combo2.AddItem "II year"
Combo2.AddItem "III year"
End Sub


