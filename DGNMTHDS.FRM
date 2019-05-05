VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form11 
   Caption         =   "Design Methods"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form11"
   ScaleHeight     =   8310
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Select Place"
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   5535
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Text            =   "Combo2"
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Level"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   5535
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "dgnmthds.frx":0000
         Left            =   1680
         List            =   "dgnmthds.frx":000D
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show &Graph"
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<-- &Back"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6255
      Left            =   240
      OleObjectBlob   =   "dgnmthds.frx":0027
      TabIndex        =   3
      Top             =   360
      Width           =   9975
   End
End
Attribute VB_Name = "FORM11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Dim col As Integer, combo2text As String, combo2texttemp As String
Dim ctr As Long, intmsg As Integer
Dim appexcel As Workbook, wsheet As Worksheet
Set appexcel = GetObject("c:\anant\anant2.xls")
Set wsheet = appexcel.Sheets(1)
Combo2.Clear
Combo2.Text = "Select Place"
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
        ElseIf (col = 4 And InStr(UCase(combo2texttemp), "UT") <> 0) Then
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
If keystat = True Then
    Frame1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = True
    Combo1.Visible = True
    Combo2.Visible = True
    MSChart1.Visible = False
    Command3.Visible = True
    keystat = False
Else
Me.Hide
Form1.Show
Form1.Timer1.Interval = 1250
End If
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Command3_Click()
If (Combo1.ListIndex < 0) Then
    intmsg = MsgBox("First select U'r choice from LEVEL")
Exit Sub
End If
Command3.Visible = False
keystat = True
Frame1.Visible = False
Frame2.Visible = False
'Frame3.Visible = False
Combo1.Visible = False
Combo2.Visible = False
MSChart1.Visible = True
'Call calconve
'MSChart1.ChartData = x1
'MSChart1.Refresh
End Sub
Private Sub Form_Load()
Command1.Top = screen.Height * 0.85
Command1.Left = screen.Width / 2 - 3375
Command2.Left = screen.Width / 2 + 2000
Command3.Left = screen.Width / 2 - 700
Command2.Top = screen.Height * 0.85
Command3.Top = screen.Height * 0.85
FORM11.WindowState = 2
MSChart1.Width = screen.Width
MSChart1.Height = screen.Height - 1500
MSChart1.Top = 0
MSChart1.Left = 0
MSChart1.Visible = False
End Sub
Public Sub mthds(key As Integer)
Private Sub matlabcost()
Dim appexcel As Workbook, appexcel1 As Workbook, wsheet As Worksheet, wsheet1 As Worksheet
Dim mat As Long, lab As Long, total As Long, popln As Integer
Dim bit As Long, screen As Long, aggt As Long
Dim ctr As Long
Set appexcel = GetObject("c:\anant\anant2.xls")
Set appexcel1 = GetObject("c:\anant\result1.xls")
Set wsheet = appexcel.Sheets(1)
Set wsheet1 = appexcel1.Sheets(1)
wsheet1.Range("a1:d310").ClearContents
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
    End If
Next ctr
Form2.Hide
On Error Resume Next
'appexcel1.SaveAs FileName:="c:\anant\matlab.xls"
appexcel.Close (vbNo)
appexcel1.Close (vbNo)
Set appexcel = Nothing
Set appexcel1 = Nothing
Set wsheet = Nothing
Set wsheet1 = Nothing
End Sub

End Sub
