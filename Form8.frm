VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form8"
   ScaleHeight     =   6345
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show &Graph"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<-- &Back"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form8.frx":0000
      Left            =   3360
      List            =   "Form8.frx":0013
      TabIndex        =   2
      Text            =   "Layer"
      Top             =   4680
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form8.frx":0040
      Left            =   3360
      List            =   "Form8.frx":0042
      TabIndex        =   1
      Text            =   "Place"
      Top             =   3000
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form8.frx":0044
      Left            =   3360
      List            =   "Form8.frx":0051
      TabIndex        =   0
      Text            =   "Level of Comparison"
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Level of Comparison "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Place of Comparison "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   1800
      TabIndex        =   4
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Frame Frame3 
      Caption         =   " Layer of Comparison "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   1800
      TabIndex        =   5
      Top             =   4320
      Width           =   4575
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5055
      Left            =   360
      OleObjectBlob   =   "Form8.frx":006B
      TabIndex        =   9
      Top             =   360
      Width           =   7815
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x(1, 5) As Variant, intmsg As Integer
Private Sub Form_Load()
Command1.Top = screen.Height * 0.85
Command1.Left = screen.Width / 2 - 3375
Command3.Left = screen.Width / 2 + 2000
Command2.Left = screen.Width / 2 - 700
Command2.Top = screen.Height * 0.85
Command3.Top = screen.Height * 0.85
Form8.WindowState = 2
MSChart1.Width = screen.Width
MSChart1.Height = screen.Height - 1500
MSChart1.Top = 0
MSChart1.Left = 0
MSChart1.Visible = False
End Sub
Private Sub Combo1_Click()
Dim col As Integer, combo2text As String, combo2texttemp As String
Dim ctr As Long
Dim appexcel As Workbook, wsheet As Worksheet
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
Combo2.Clear
Combo2.Text = "Place"
If Combo1.ListIndex = 0 Then
    col = 4
ElseIf Combo1.ListIndex = 1 Then
    col = 3
ElseIf UCase(Combo1.Text) = "ALL" Then
    ' do nothing
    col = 3
Else
    intmsg = MsgBox("Something is wrong here..Fix it...")
End If
For ctr = 2 To Form1.total - 2
    combo2texttemp = wsheet.Cells(ctr, col).Value
    If (InStr(combo2text, combo2texttemp) = 0) Then
        combo2text = combo2text & combo2texttemp
        If (col = 4 And InStr(UCase(combo2texttemp), "UP") <> 0) Then
            Combo2.AddItem "Uttar Pradesh"
        ElseIf (col = 4 And ((InStr(UCase(combo2texttemp), "UT") <> 0) Or (InStr(UCase(combo2texttemp), "UA") <> 0))) Then
            Combo2.AddItem "Uttaranchal"
        ElseIf (col = 4 And InStr(UCase(combo2texttemp), "BR") <> 0) Then
            Combo2.AddItem "Bihar"
        ElseIf (col = 3) Then
            Combo2.AddItem combo2texttemp & " (" & wsheet.Cells(ctr, 4) & ")"
        Else
            intmsg = MsgBox("There is something wrong here")
    End If
End If
Next ctr
On Error Resume Next
appexcel.Close (vbNo)
Set appexcel = Nothing
Set wsheet = Nothing
End Sub
Private Sub Command1_Click()
If (Command2.Visible = False) Then
    Frame1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = True
    Combo1.Visible = True
    Combo2.Visible = True
    Combo3.Visible = True
    MSChart1.Visible = False
    Command2.Visible = True
Else
Me.Hide
Form1.Show
Form1.Timer1.Interval = 1250
End If
End Sub
Private Sub Command3_Click()
End
End Sub
Private Sub Command2_Click()
If (Combo1.ListIndex < 0) Then
    intmsg = MsgBox("First select U'r choice from LEVEL")
    Exit Sub
ElseIf (Combo2.ListIndex < 0 And UCase(Combo1.Text) <> "ALL") Then
    intmsg = MsgBox("First select U'r choice from COMBO 2")
    Exit Sub
ElseIf (Combo3.ListIndex < 0) Then
    intmsg = MsgBox("First select U'r choice from COMBO 3")
    Exit Sub
End If
Command3.Visible = True
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Combo1.Visible = False
Combo2.Visible = False
Combo3.Visible = False
MSChart1.Visible = True
Call matlabcost
'MSChart1.ChartData = x
MSChart1.Refresh
Command2.Visible = False
End Sub
Private Sub matlabcost()
Dim appexcel As Workbook, wsheet As Worksheet
Dim mat As Long, lab As Long, total As Long, popln As Integer
Dim bit As Long, screen As Long, aggt As Long
Dim ctr As Long
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
Dim col As Integer
Dim texttemp As String, textstr As String
If (Combo1.ListIndex = 0) Then
    col = 4
ElseIf (Combo1.ListIndex = 1) Then
    col = 3
ElseIf (UCase(Combo1.Text) = "ALL") Then
    col = 3
Else
    intmsg = MsgBox("There is something wrong here")
End If
    If (col = 3) Then
        texttemp = Left(Combo2.Text, Len(Combo2.Text) - 5)
    ElseIf col = 4 Then
        If UCase(Combo2.Text) = "UTTARANCHAL" Then
            texttemp = "UTUA"
        ElseIf UCase(Combo2.Text) = "UTTAR PRADESH" Then
            texttemp = "UP"
        ElseIf UCase(Combo2.Text) = "BIHAR" Then
            texttemp = "BR"
        Else
            intmsg = MsgBox("There is something wrong here")
        End If
    End If
    Form2.Show
    Form8.Hide
For ctr = 2 To Form1.total - 2
    Form2.Label2 = "  " & CByte(ctr * 100 / Form1.total) & "% Completed...  "
    Form2.Label2.Refresh
    Form2.ProgressBar1.Value = ctr / Form1.total * 100
    If ((InStr(texttemp, wsheet.Cells(ctr, col)) <> 0) Or (UCase(Combo1.Text) = "ALL")) Then
     popln = popln + 1
     If (Combo3.ListIndex = 0) Then
        mat = mat + wsheet.Cells(ctr, 62).Value + wsheet.Cells(ctr, 77).Value + wsheet.Cells(ctr, 96).Value + wsheet.Cells(ctr, 98).Value + wsheet.Cells(ctr, 104).Value + wsheet.Cells(ctr, 106).Value + wsheet.Cells(ctr, 112).Value + wsheet.Cells(ctr, 114).Value + wsheet.Cells(ctr, 120).Value + wsheet.Cells(ctr, 122).Value + wsheet.Cells(ctr, 128).Value + wsheet.Cells(ctr, 130).Value
        lab = lab + wsheet.Cells(ctr, 20).Value + wsheet.Cells(ctr, 63).Value + wsheet.Cells(ctr, 78).Value + wsheet.Cells(ctr, 93).Value + wsheet.Cells(ctr, 97).Value + wsheet.Cells(ctr, 105).Value + wsheet.Cells(ctr, 113).Value + wsheet.Cells(ctr, 121).Value + wsheet.Cells(ctr, 129).Value
        bit = bit + wsheet.Cells(ctr, 96).Value + wsheet.Cells(ctr, 104).Value + wsheet.Cells(ctr, 112).Value + wsheet.Cells(ctr, 120).Value + wsheet.Cells(ctr, 128).Value
        aggt = aggt + wsheet.Cells(ctr, 64).Value + wsheet.Cells(ctr, 79).Value + wsheet.Cells(ctr, 98).Value + wsheet.Cells(ctr, 106).Value + wsheet.Cells(ctr, 114).Value + wsheet.Cells(ctr, 122).Value + wsheet.Cells(ctr, 130).Value
        screen = screen + wsheet.Cells(ctr, 65).Value + wsheet.Cells(ctr, 80).Value + wsheet.Cells(ctr, 99).Value + wsheet.Cells(ctr, 107).Value + wsheet.Cells(ctr, 115).Value + wsheet.Cells(ctr, 123).Value + wsheet.Cells(ctr, 131).Value
        total = total + wsheet.Cells(ctr, 138).Value
    ElseIf (Combo3.ListIndex = 1) Then
        mat = mat + wsheet.Cells(ctr, 96).Value + wsheet.Cells(ctr, 98).Value + wsheet.Cells(ctr, 104).Value + wsheet.Cells(ctr, 106).Value + wsheet.Cells(ctr, 112).Value + wsheet.Cells(ctr, 114).Value + wsheet.Cells(ctr, 120).Value + wsheet.Cells(ctr, 122).Value + wsheet.Cells(ctr, 128).Value + wsheet.Cells(ctr, 130).Value
        lab = lab + wsheet.Cells(ctr, 93).Value + wsheet.Cells(ctr, 97).Value + wsheet.Cells(ctr, 105).Value + wsheet.Cells(ctr, 113).Value + wsheet.Cells(ctr, 121).Value + wsheet.Cells(ctr, 129).Value
        bit = bit + wsheet.Cells(ctr, 96).Value + wsheet.Cells(ctr, 104).Value + wsheet.Cells(ctr, 112).Value + wsheet.Cells(ctr, 120).Value + wsheet.Cells(ctr, 128).Value
        aggt = aggt + wsheet.Cells(ctr, 98).Value + wsheet.Cells(ctr, 106).Value + wsheet.Cells(ctr, 114).Value + wsheet.Cells(ctr, 122).Value + wsheet.Cells(ctr, 130).Value
        screen = screen + wsheet.Cells(ctr, 99).Value + wsheet.Cells(ctr, 107).Value + wsheet.Cells(ctr, 115).Value + wsheet.Cells(ctr, 123).Value + wsheet.Cells(ctr, 131).Value
        total = total + wsheet.Cells(ctr, 93).Value + wsheet.Cells(ctr, 100).Value + wsheet.Cells(ctr, 108).Value + wsheet.Cells(ctr, 116).Value + wsheet.Cells(ctr, 124).Value + wsheet.Cells(ctr, 132).Value
    ElseIf (Combo3.ListIndex = 2) Then
        mat = mat + wsheet.Cells(ctr, 77).Value
        lab = lab + wsheet.Cells(ctr, 78).Value
        bit = bit + 0
        aggt = aggt + wsheet.Cells(ctr, 79).Value
        screen = screen + wsheet.Cells(ctr, 80).Value
        total = total + wsheet.Cells(ctr, 81).Value
    ElseIf (Combo3.ListIndex = 3) Then
        mat = mat + wsheet.Cells(ctr, 62).Value
        lab = lab + wsheet.Cells(ctr, 63).Value
        aggt = aggt + wsheet.Cells(ctr, 64).Value
        screen = screen + wsheet.Cells(ctr, 65).Value
        bit = bit + 0
        total = total + wsheet.Cells(ctr, 66).Value
    ElseIf (Combo3.ListIndex = 4) Then
        lab = lab + wsheet.Cells(ctr, 20)
        aggt = aggt + 0
        screen = screen + 0
        total = total + wsheet.Cells(ctr, 20)
        bit = bit + 0
    Else
        intmsg = MsgBox("There is something wrong here")
    End If
End If
Next ctr
Form2.Hide
Form8.Show
x(0, 0) = "Material"
x(0, 1) = "Labour"
x(0, 2) = "Aggt"
x(0, 3) = "Screening"
x(0, 4) = "Bitumen"
x(0, 5) = "Total"
x(1, 0) = mat / popln
x(1, 1) = lab / popln
x(1, 2) = aggt / popln
x(1, 3) = screen / popln
x(1, 4) = bit / popln
x(1, 5) = total / popln
MSChart1.ChartData = x
On Error Resume Next
appexcel.Close (vbNo)
Set appexcel = Nothing
Set wsheet = Nothing
End Sub
