VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form9 
   Caption         =   "Item"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form9"
   ScaleHeight     =   5550
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form9.frx":0000
      Left            =   3000
      List            =   "Form9.frx":000D
      TabIndex        =   5
      Text            =   "Level of Comparison"
      Top             =   960
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Text            =   "Place"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form9.frx":0027
      Left            =   3000
      List            =   "Form9.frx":0031
      TabIndex        =   3
      Text            =   "Item"
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<-- &Back"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show &Graph"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
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
      Left            =   1440
      TabIndex        =   7
      Top             =   2280
      Width           =   4575
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
      Left            =   1440
      TabIndex        =   6
      Top             =   600
      Width           =   4575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item of Comparison "
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
      Left            =   1440
      TabIndex        =   8
      Top             =   3960
      Width           =   4575
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5055
      Left            =   0
      OleObjectBlob   =   "Form9.frx":0055
      TabIndex        =   9
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x() As Variant, intmsg As Integer
Dim var() As Long
Private Sub Form_Load()
Command1.Top = screen.Height * 0.85
Command1.Left = screen.Width / 2 - 3375
Command3.Left = screen.Width / 2 + 2000
Command2.Left = screen.Width / 2 - 700
Command2.Top = screen.Height * 0.85
Command3.Top = screen.Height * 0.85
Form9.WindowState = 2
MSChart1.Width = screen.Width
MSChart1.Height = screen.Height - 1500
MSChart1.Top = 0
MSChart1.Left = 0
MSChart1.Visible = False
End Sub
Private Sub Combo1_Click()
Dim col As Integer, combo2text As String, combo2texttemp As String
Dim ctr As Integer
Dim appexcel As Workbook, wsheet As Worksheet
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
Combo2.Clear
Combo2.Text = "Place "
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
Call layercd
MSChart1.Refresh
Command2.Visible = False
End Sub
Private Sub layercd()
Dim appexcel As Workbook, wsheet As Worksheet
Dim mat As Long, lab As Long, total As Long, popln As Integer, cdpopln(11) As Integer
Dim bit As Long, screen As Long, aggt As Long
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
    Form9.Hide
If (Combo3.ListIndex = 0) Then
     ReDim x(1, 5)
     ReDim var(5)
ElseIf (Combo3.ListIndex = 1) Then
    ReDim x(1, 11)
    ReDim var(11)
Else
    intmsg = MsgBox("There is something wrong here")
End If
For ctr = 2 To Form1.total - 2
    Form2.Label2 = "  " & CByte(ctr * 100 / Form1.total) & "% Completed...  "
    Form2.Label2.Refresh
    Form2.ProgressBar1.Value = ctr / Form1.total * 100
    If ((InStr(texttemp, wsheet.Cells(ctr, col)) <> 0) Or (UCase(Combo1.Text) = "ALL")) Then
     If (Combo3.ListIndex = 0) Then
     popln = popln + 1
     var(0) = var(0) + wsheet.Cells(ctr, 20).Value
     var(1) = var(1) + wsheet.Cells(ctr, 66).Value
     var(2) = var(2) + wsheet.Cells(ctr, 81).Value
     var(3) = var(3) + wsheet.Cells(ctr, 93).Value + wsheet.Cells(ctr, 100).Value + wsheet.Cells(ctr, 108).Value + wsheet.Cells(ctr, 116).Value + wsheet.Cells(ctr, 124).Value + wsheet.Cells(ctr, 132).Value
     var(4) = var(4) + (wsheet.Cells(ctr, 25).Value + wsheet.Cells(ctr, 27).Value + wsheet.Cells(ctr, 29).Value + wsheet.Cells(ctr, 31).Value + wsheet.Cells(ctr, 33).Value + wsheet.Cells(ctr, 35).Value + wsheet.Cells(ctr, 37).Value + wsheet.Cells(ctr, 39).Value + wsheet.Cells(ctr, 42).Value + wsheet.Cells(ctr, 45).Value + wsheet.Cells(ctr, 48).Value) / wsheet.Cells(ctr, 13).Value
     var(5) = var(5) + wsheet.Cells(ctr, 138).Value
     ElseIf (Combo3.ListIndex = 1) Then
        Dim cdcol, costcol, nocol As Integer
        For cdcol = 0 To 11
            If (cdcol <= 7) Then
            costcol = cdcol * 2 + 25
            nocol = costcol - 1
            ElseIf (cdcol <= 10) Then
            costcol = 42 + (cdcol - 8) * 3
            nocol = costcol - 2
            ElseIf (cdcol = 11) Then
            costcol = 52
            nocol = 51
            Else
            intmsg = MsgBox("Thdere is somethign wrong here")
            End If
            x(1, cdcol) = x(1, cdcol) + wsheet.Cells(ctr, costcol)
            cdpopln(cdcol) = cdpopln(cdcol) + wsheet.Cells(ctr, nocol)
        Next cdcol
      Else
        intmsg = MsgBox("There is something wrong here")
     End If
    End If
Next ctr
If (Combo3.ListIndex = 0) Then
     x(0, 0) = "Earthwork"
     x(0, 1) = "Sub-Base"
     x(0, 2) = "Base"
     x(0, 3) = "Surface"
     x(0, 4) = "C/D Structure"
     x(0, 5) = "Total"
    x(1, 0) = var(0) / popln
    x(1, 1) = var(1) / popln
    x(1, 2) = var(2) / popln
    x(1, 3) = var(3) / popln
    x(1, 4) = var(4) / popln
    x(1, 5) = var(5) / popln
    For ctr = 0 To 5
    Next
ElseIf (Combo3.ListIndex = 1) Then
        x(0, 0) = "1000 mm. dia Hume Pipe Culvert"
        x(0, 1) = "650 mm. dia Hume Pipe Culvert"
        x(0, 2) = "600 mm. dia Hume Pipe Culvert"
        x(0, 3) = "450 mm. dia Hume Pipe Culvert"
        x(0, 4) = "350 mm. dia Hume Pipe Culvert"
        x(0, 5) = "300 mm. dia Hume Pipe Culvert"
        x(0, 6) = "2 row 1000 mm dia hume pipe culvert"
        x(0, 7) = "2 row 1200 mm dia hume pipe culvert"
        x(0, 8) = "RCC Box Culvert"
        x(0, 9) = "Cause Ways"
        x(0, 10) = "Minor Bridge"
        x(0, 11) = "Scupper"
        For ctr = 0 To 11
        If cdpopln(ctr) <> 0 Then
        x(1, ctr) = x(1, ctr) / cdpopln(ctr)
        End If
    Next
End If
MSChart1.ChartData = x
Form2.Hide
Form9.Show
On Error Resume Next
appexcel.Close (vbNo)
Set appexcel = Nothing
Set wsheet = Nothing
End Sub
