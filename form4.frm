VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form4"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
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
      Left            =   3600
      TabIndex        =   2
      Text            =   "select"
      Top             =   4200
      Width           =   3015
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
      Left            =   12240
      TabIndex        =   1
      Text            =   "year"
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULATE  "
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7920
      TabIndex        =   0
      Top             =   7080
      Width           =   3135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "       RC CALCULATION"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   6600
      TabIndex        =   3
      Top             =   1800
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   6600
      Picture         =   "form4.frx":0000
      Top             =   120
      Width           =   5640
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sum As Double
Dim sum1 As Double
Dim h As Double
Dim fine As Double
Private Sub Command1_Click()
sum = 0
sum1 = 0
Data1.DatabaseName = "f:\attendance\" & Combo1.Text & ".xlsx"
Data1.RecordSource = "'" & Combo2.Text & "'" & "$"
Data1.Refresh
Data1.Recordset.MoveFirst
Do While Data1.Recordset.EOF <> True
If ((IsNull(Data1.Recordset(1))) Or (Data1.Recordset(1) = "STUDENT NAME")) Then
Data1.Recordset.MoveNext
Else
For i = 2 To 6
sum = sum + Data1.Recordset(i)
Next
Data1.Recordset.Edit
Data1.Recordset(7) = sum
Data1.Recordset.Update
If Data1.Recordset(1) = "TOTAL WORKING DAYS" Then
h = sum
GoTo y
ElseIf sum <= 67 Then
Data1.Recordset.Edit
Data1.Recordset(8) = "RC"
Data1.Recordset.Update
Else
Data1.Recordset.Edit
Data1.Recordset(8) = "-"
Data1.Recordset.Update
End If
y: If Data1.Recordset(1) = "TOTAL WORKING DAYS" Then
GoTo h
Else
sum1 = h - sum
fine = sum1 * 0.5
Data1.Recordset.Edit
Data1.Recordset(9) = fine
Data1.Recordset.Update
End If
h: sum = 0
sum1 = 0
fine = 0
Data1.Recordset.MoveNext
End If
Loop
MsgBox ("RC calculation is done properly")
End Sub
Private Sub Form_Load()
Combo1.AddItem "sem-even"
Combo1.AddItem "sem-odd"
Combo2.AddItem "I year"
Combo2.AddItem "II year"
Combo2.AddItem "III year"
End Sub



