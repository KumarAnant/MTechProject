VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conveyance Cost"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form7.frx":0000
      Left            =   3840
      List            =   "Form7.frx":000A
      TabIndex        =   6
      Text            =   "Select Parameter"
      Top             =   5040
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show &Graph"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<-- &Back"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form7.frx":003F
      Left            =   3840
      List            =   "Form7.frx":0041
      TabIndex        =   1
      Text            =   "Select Place"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form7.frx":0043
      Left            =   3840
      List            =   "Form7.frx":004D
      TabIndex        =   0
      Text            =   "Select Level"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select  Level"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Place"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   1920
      TabIndex        =   8
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Select Parametre"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   1920
      TabIndex        =   9
      Top             =   4680
      Width           =   4455
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5175
      Left            =   -960
      OleObjectBlob   =   "Form7.frx":0062
      TabIndex        =   5
      Top             =   360
      Width           =   10935
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim keystat As Boolean
Dim intmsg As Integer, x(300, 1) As Long
Dim x1(300, 1) As Long
Private Sub Combo1_Click()
Dim col As Integer, combo2text As String, combo2texttemp As String
Dim ctr As Long, intmsg As Integer
Dim appexcel As Workbook, wsheet As Worksheet
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
Combo2.Clear
Combo2.Text = "Select Place"
If Combo1.ListIndex = 0 Then
    col = 4
ElseIf Combo1.ListIndex = 1 Then
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
On Error Resume Next
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
    Combo3.Visible = True
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
Frame3.Visible = False
Combo1.Visible = False
Combo2.Visible = False
Combo3.Visible = False
MSChart1.Visible = True
Call calconve
MSChart1.ChartData = x1
MSChart1.Refresh
End Sub
Private Sub Form_Load()
Command1.Top = screen.Height * 0.85
Command1.Left = screen.Width / 2 - 3375
Command2.Left = screen.Width / 2 + 2000
Command3.Left = screen.Width / 2 - 700
Command2.Top = screen.Height * 0.85
Command3.Top = screen.Height * 0.85
Form7.WindowState = 2
MSChart1.Width = screen.Width
MSChart1.Height = screen.Height - 1500
MSChart1.Top = 0
MSChart1.Left = 0
MSChart1.Visible = False
End Sub
Private Sub calconve()
Dim ctr As Long, col As Integer
Dim appexcel As Workbook, wsheet As Worksheet
Dim texttemp As String, popln(100) As Integer, popln1 As Integer, duplicate As Integer
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
duplicate = 0
If (Combo1.ListIndex = 0) Then
    col = 4
ElseIf (Combo1.ListIndex = 1) Then
    col = 3
Else
    intmsg = MsgBox("There id something wrong here")
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
Form7.Hide
Form2.Show
For ctr = 2 To Form1.total - 2
    Form2.Label2 = "  " & CByte(ctr * 100 / Form1.total) & "% Completed...  "
    Form2.Label2.Refresh
    Form2.ProgressBar1.Value = ctr / Form1.total * 100
    Dim km As Integer
    duplicate = 1
    Do While (wsheet.Cells(ctr, 193).Value = wsheet.Cells(ctr + 1, 193).Value And wsheet.Cells(ctr, 194).Value = wsheet.Cells(ctr + 1, 194).Value And wsheet.Cells(ctr, 195).Value = wsheet.Cells(ctr + 1, 195).Value And wsheet.Cells(ctr, 196).Value = wsheet.Cells(ctr + 1, 196).Value And wsheet.Cells(ctr, 197).Value = wsheet.Cells(ctr + 1, 197).Value And wsheet.Cells(ctr, 198).Value = wsheet.Cells(ctr + 1, 198).Value And wsheet.Cells(ctr, 199).Value = wsheet.Cells(ctr + 1, 199).Value And wsheet.Cells(ctr, 200).Value = wsheet.Cells(ctr + 1, 200).Value)
        duplicate = duplicate + 1
        ctr = ctr + 1
    Loop
     If (InStr(texttemp, wsheet.Cells(ctr, col)) <> 0) Then
        popln1 = popln1 + duplicate
        For km = 1 To 300
            Dim vsbase As Double, vbase As Double, cell188 As Double, cell190 As Double, cell192 As Double, cell193 As Double, cell194 As Double, cell195 As Double, cell196 As Double, cell197 As Double, cell198 As Double, cell199 As Double, cell200 As Double, cell193temp As Double, cell194temp As Double, cell195temp As Double, cell196temp As Double, cell197temp As Double, cell198temp As Double, cell199temp As Double, cell200temp As Double
            vsbase = wsheet.Cells(ctr, 60).Value * wsheet.Cells(ctr, 178).Value
            vbase = (wsheet.Cells(ctr, 179).Value + wsheet.Cells(ctr, 180).Value + wsheet.Cells(ctr, 181).Value + wsheet.Cells(ctr, 182).Value + wsheet.Cells(ctr, 183).Value + wsheet.Cells(ctr, 184).Value) / 2 * wsheet.Cells(ctr, 75).Value
            cell188 = wsheet.Cells(ctr, 188).Value
            cell190 = wsheet.Cells(ctr, 190).Value
            cell192 = wsheet.Cells(ctr, 192).Value
            cell193 = wsheet.Cells(ctr, 193).Value
            cell194 = wsheet.Cells(ctr, 194).Value
            cell195 = wsheet.Cells(ctr, 195).Value
            cell196 = wsheet.Cells(ctr, 196).Value
            cell197 = wsheet.Cells(ctr, 197).Value
            cell198 = wsheet.Cells(ctr, 198).Value
            cell199 = wsheet.Cells(ctr, 199).Value
            cell200 = wsheet.Cells(ctr, 200).Value
            If (km > 0) Then
                x(km, 1) = x(km, 1) + (vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell193 * duplicate
            Else
                intmsg = MsgBox("There is something wrong here")
            End If
            If (km > 1 And km <= 3) Then
                x(km, 1) = x(km, 1) + (km - 1) * duplicate * (vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell194
            ElseIf (km > 3) Then
                x(km, 1) = x(km, 1) + 2 * duplicate * (vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell194
            End If
            If (km >= 4 And km <= 5) Then
                x(km, 1) = x(km, 1) + (km - 3) * duplicate * (vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell195
            ElseIf (km > 5) Then
                x(km, 1) = x(km, 1) + 2 * duplicate * (vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell195
            End If
            If (km >= 6 And km <= 10) Then
                x(km, 1) = x(km, 1) + (km - 5) * duplicate * (vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell196
            ElseIf (km > 10) Then
                x(km, 1) = x(km, 1) + 5 * duplicate * (vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell196
            End If
            If (km >= 11 And km <= 20) Then
                x(km, 1) = x(km, 1) + (km - 10) * duplicate * (vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell197
            ElseIf (km > 20) Then
                x(km, 1) = x(km, 1) + 10 * duplicate * (vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell197
            End If
            If (km > 20 And km <= 30) Then
                x(km, 1) = x(km, 1) + (km - 20) * duplicate * (vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell198
            ElseIf (km > 30) Then
                x(km, 1) = x(km, 1) + 10 * duplicate * (vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell198
            End If
            If (km > 30 And km <= 50) Then
                x(km, 1) = x(km, 1) + (km - 30) * duplicate * (vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell199
            ElseIf (km > 50) Then
                x(km, 1) = x(km, 1) + 20 * duplicate * (vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell199
            End If
            If (km > 50) Then
                x(km, 1) = x(km, 1) + CLng(km - 50) * CLng(duplicate) * CLng((vbase + vsbase + (cell188 + cell190 + cell192) * 3750) * cell200)
            End If
        Next km
  End If
        For km = 0 To 300
            x1(km, 0) = km
            x1(km, 1) = x(km, 1) + x1(km, 1)
            x(km, 1) = 0
        Next km
    Next ctr
    For ctr = 0 To 300
        If (popln1 <> 0) Then
        x1(ctr, 1) = x1(ctr, 1) / popln1 / 10000
        End If
    Next ctr
    MSChart1.ChartData = x1
    Form2.Hide
    Form7.Show
    On Error Resume Next
    appexcel.Close (vbNo)
    Set appexcel = Nothing
    Set wsheet = Nothing
End Sub

