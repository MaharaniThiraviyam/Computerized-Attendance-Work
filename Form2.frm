VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
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
      Left            =   12720
      TabIndex        =   3
      Text            =   "Year"
      Top             =   5880
      Width           =   2655
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
      Left            =   3840
      TabIndex        =   2
      Text            =   "Month"
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Data Data1 
      Caption         =   "Month"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data2 
      Caption         =   "SEMESTER"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULATE"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   1
      Top             =   7680
      Width           =   2535
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
      Left            =   13560
      TabIndex        =   0
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   7680
      Picture         =   "Form2.frx":0000
      Top             =   480
      Width           =   5640
   End
   Begin VB.Label Label1 
      Caption         =   "  WORKING DAYS"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   5
      Top             =   3720
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "CALCULATE THE PRESENT DAYS PER MONTH"
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
      Height          =   975
      Left            =   7560
      TabIndex        =   4
      Top             =   1800
      Width           =   5895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Double
Dim m As Integer
Dim S As Double
Dim mh As String
Dim z As Integer
Dim a As Double
Dim mon As String
Dim semester As String
Dim find As String
Private Sub Command1_Click()
z = 0
c = 0
S = 0
a = 0
mon = Combo1.Text
If mon = "dec" Or mon = "jan" Or mon = "feb" Or mon = "mar" Or mon = "apr" Then
semester = "sem-even"
find = "e-month"
Else
semester = "sem-odd"
find = "o-month"
End If
Data1.DatabaseName = "f:\attendance\" & find & ".xlsx"
Data1.RecordSource = "'" & mon & "-" & Combo2.Text & "'" & "$"
Data1.Refresh
Data2.DatabaseName = "f:\attendance\" & semester & ".xlsx"
Data2.RecordSource = "'" & Combo2.Text & "'" & "$"
Data2.Refresh
mh = Combo1.Text
Select Case mh
Case "jan": m = 3
Case "feb": m = 4
Case "mar": m = 5
Case "apr": m = 6
Case "jun": m = 2
Case "july": m = 3
Case "aug": m = 4
Case "sep": m = 5
Case "oct": m = 6
Case "dec": m = 2
End Select
Data1.Recordset.MoveFirst
Data2.Recordset.MoveFirst
Do While Data1.Recordset.EOF <> True And Data2.Recordset.EOF <> True
If ((IsNull(Data1.Recordset(1))) Or (Data1.Recordset(1) = "STUDENT NAME")) Then
Data1.Recordset.MoveNext
ElseIf ((IsNull(Data2.Recordset(1))) Or (Data2.Recordset(1) = "STUDENT NAME") Or (Data2.Recordset(1) = "TOTAL WORKING DAYS")) Then
Data2.Recordset.MoveNext
Else
For j = 2 To 32
If Data1.Recordset(j) = "A" Then
c = c + 1
End If
If Data1.Recordset(j) = "H" Then
c = c + 0.5
End If
Next
z = Text1.Text
a = z - c
Data1.Recordset.Edit
Data1.Recordset(33) = a
Data1.Recordset.Update

Data2.Recordset.Edit
Data2.Recordset(m) = a
Data2.Recordset.Update
Data1.Recordset.MoveNext
Data2.Recordset.MoveNext

a = 0
c = 0
S = 0
End If
Loop
Data2.Recordset.MoveFirst
Do While Data2.Recordset.EOF <> True
If Data2.Recordset(1) = "TOTAL WORKING DAYS" Then
Data2.Recordset.Edit
Data2.Recordset(m) = Text1.Text
Data2.Recordset.Update
End If
Data2.Recordset.MoveNext
Loop
MsgBox ("The calculation is completed sucessfully")
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





