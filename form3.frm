VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FF80FF&
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
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
      Left            =   11760
      TabIndex        =   4
      Top             =   3600
      Width           =   3735
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
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   11640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
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
      Height          =   735
      Left            =   7080
      TabIndex        =   3
      Top             =   8400
      Width           =   4095
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
      Left            =   14400
      TabIndex        =   2
      Text            =   "Year"
      Top             =   5640
      Width           =   2535
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
      TabIndex        =   1
      Text            =   "Date"
      Top             =   5640
      Width           =   2535
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
      Left            =   7680
      TabIndex        =   0
      Text            =   "Month"
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "   Roll Number"
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
      Left            =   2520
      TabIndex        =   6
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "    OD CALCULATION FORM "
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   5880
      TabIndex        =   5
      Top             =   1800
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   6000
      Picture         =   "form3.frx":0000
      Top             =   240
      Width           =   5640
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mon As String
Dim g As String
Dim find As String
Private Sub Command1_Click()
g = Text1.Text
mon = Combo2.Text
If mon = "dec" Or m = "jan" Or mon = "feb" Or mon = "mar" Or mon = "apr" Then
semester = "sem-even"
find = "e-month"
Else
semester = "sem-odd"
find = "o-month"
End If
Data1.DatabaseName = "f:\attendance\" & find & ".xlsx"
Select Case mon
Case "dec": Data1.RecordSource = "'" & mon & "-" & Combo3.Text & "'" & "$"
Data1.Refresh: S = 2
Case "jan": Data1.RecordSource = "'" & mon & "-" & Combo3.Text & "'" & "$"
Data1.Refresh: S = 3
Case "feb": Data1.RecordSource = "'" & mon & "-" & Combo3.Text & "'" & "$"
Data1.Refresh: S = 4
Case "mar": Data1.RecordSource = "'" & mon & "-" & Combo3.Text & "'" & "$"
Data1.Refresh: S = 5
Case "apr": Data1.RecordSource = "'" & mon & "-" & Combo3.Text & "'" & "$"
Data1.Refresh: S = 6
Case "jun": Data1.RecordSource = "'" & mon & "-" & Combo3.Text & "'" & "$"
Data1.Refresh: S = 2
Case "july": Data1.RecordSource = "'" & mon & "-" & Combo3.Text & "'" & "$"
Data1.Refresh: S = 3
Case "aug": Data1.RecordSource = "'" & mon & "-" & Combo3.Text & "'" & "$"
Data1.Refresh: S = 4
Case "sep": Data1.RecordSource = "'" & mon & "-" & Combo3.Text & "'" & "$"
Data1.Refresh: S = 5
Case "oct": Data1.RecordSource = "'" & mon & "-" & Combo3.Text & "'" & "$"
Data1.Refresh: S = 6
End Select
Data1.Recordset.MoveFirst
Data2.DatabaseName = "f:\attendance\" & semester & ".xlsx"
Data2.RecordSource = "'" & Combo3.Text & "'" & "$"
Data2.Refresh
d = Combo1.Text
Do While Data1.Recordset.EOF <> True
If ((IsNull(Data1.Recordset(1))) Or (Data1.Recordset(1) = "STUDENT NAME")) Then
Data1.Recordset.MoveNext
End If
If Data1.Recordset(0) = g Then
d = d + 1
If Data1.Recordset(d) = "A" Then
Data1.Recordset.Edit
Data1.Recordset(d) = "O"
Data1.Recordset.Update
Data1.Recordset.Edit
Data1.Recordset(33) = Data1.Recordset(33) + 1
Data1.Recordset.Update
Exit Do
Else
Data1.Recordset.Edit
Data1.Recordset(d) = "O"
Data1.Recordset.Update
Data1.Recordset.Edit
Data1.Recordset(33) = Data1.Recordset(33) + 0.5
Data1.Recordset.Update
Exit Do
End If
Else
Data1.Recordset.MoveNext
End If
Loop
Data2.Recordset.MoveFirst
Do While Data2.Recordset.EOF <> True
Do While IsNull(Data2.Recordset(1))
Data2.Recordset.MoveNext
Loop
If Data2.Recordset(0) = g Then
Data2.Recordset.Edit
Data2.Recordset(S) = Data1.Recordset(33)
Data2.Recordset.Update
Exit Do
Else
Data2.Recordset.MoveNext
End If
Loop
MsgBox ("OD provided successfully")
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
End Sub





