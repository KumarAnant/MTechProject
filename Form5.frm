VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form5 
   Caption         =   "District"
   ClientHeight    =   6450
   ClientLeft      =   2580
   ClientTop       =   2220
   ClientWidth     =   8580
   LinkTopic       =   "Form5"
   ScaleHeight     =   6450
   ScaleWidth      =   8580
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form5.frx":0000
      Left            =   5400
      List            =   "Form5.frx":0013
      TabIndex        =   4
      Text            =   "Type of Work"
      Top             =   5760
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form5.frx":0062
      Left            =   5400
      List            =   "Form5.frx":0072
      TabIndex        =   3
      Text            =   "Select State"
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<-- &Back"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   5160
      Width           =   975
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6615
      Left            =   0
      OleObjectBlob   =   "Form5.frx":009D
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim appexcel As excel.Workbook
Dim wsheet As Worksheet
Private x() As Variant
Private Sub Combo1_Click()
Call showchart
End Sub
Private Sub Combo2_Click()
Call showchart
End Sub
Private Sub Command1_Click()
Form1.Show
Form1.WindowState = 2
Form1.Timer1.Interval = 1250
Form5.Hide
End Sub
Private Sub Command3_Click()
End
End Sub
Private Sub Form_Load()
Form5.WindowState = 2
Command1.Top = screen.Height * 0.85
Command3.Top = Command1.Top
Command1.Left = screen.Width / 2 - 800 - Command1.Width
Command3.Left = screen.Width / 2 + 800
MSChart1.Width = screen.Width - 300
MSChart1.Height = screen.Height - 2000
MSChart1.Top = 400
MSChart1.Left = 190
Combo1.Top = screen.Height * 0.8
Combo2.Top = Combo1.Top + 475
Combo1.Left = screen.Width - 3050
Combo2.Left = Combo1.Left
MSChart1.Width = screen.Width
MSChart1.Left = 0
Call showchart
End Sub
Private Sub showchart()
Dim state As String, distlist As String, disttemp As String, filetemp As String
state = ""
distlist = ""
disttemp = ""
Dim ctr As Long
ctr = 0
Dim intmsg As Integer, nconstr As Integer, inttemp As Long
intmsg = nconstr = inttemp = 0
Dim otptno As Integer, serialno As Integer
otptno = 0
serialno = 0
Dim dist As New Collection, distrcrd As New Collection, serial As New Collection
Set dist = Nothing
Set distrcrd = Nothing
Set serial = Nothing
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
Dim costtemp As Long
costtemp = 0
distlist = "EMPTY"
If (Combo1.ListIndex = 0) Then
state = "UP,Up,up,Up,Ut,ut,UT,uT,Bh,BH,bH,bh,Ua,UA,uA,uA"
ElseIf (Combo1.ListIndex = 1) Then
    state = "UP,Up,up,Up"
ElseIf (Combo1.ListIndex = 2) Then
    state = "UT,ut,Ut,uT,Ua,UA,uA,ua"
ElseIf (Combo1.ListIndex = 3) Then
    state = "BR,Br,bR,br"
Else
    intmsg = MsgBox("Wrong Combo INDEX.. Fix It...By default 'ALL' has been selected")
    state = "UP,Up,up,Up,Ut,ut,UT,uT,Br,BR,bR,br,Ua,UA,uA,uA"
otptno = 2
End If
nconstr = 0
serialno = 1
Form2.Show
Form5.Hide
For ctr = 2 To Form1.total Step 1
costtemp = 0
If (wsheet.Cells(ctr, 3).Value <> "") Then
    Form2.ProgressBar1.Value = ctr * 100 / Form1.total
    Form2.Label2.Caption = CByte(ctr * 100 / Form1.total) & " % Completed"
    Form2.Label2.Refresh
 If (InStr(state, wsheet.Cells(ctr, 4)) <> 0 And nconstr = 0) Then
    disttemp = CStr(wsheet.Cells(ctr, 3).Value)
    If (Combo2.ListIndex = 0) Then
    costtemp = wsheet.Cells(ctr, 138)
    filetemp = "_RD"
    ElseIf (Combo2.ListIndex = 1) Then
    costtemp = (wsheet.Cells(ctr, 25) + wsheet.Cells(ctr, 27) + wsheet.Cells(ctr, 29) + wsheet.Cells(ctr, 31) + wsheet.Cells(ctr, 33) + wsheet.Cells(ctr, 35) + wsheet.Cells(ctr, 37) + wsheet.Cells(ctr, 39) + wsheet.Cells(ctr, 42) + wsheet.Cells(ctr, 45) + wsheet.Cells(ctr, 48) + wsheet.Cells(ctr, 50) + wsheet.Cells(ctr, 52) + wsheet.Cells(ctr, 54)) / wsheet.Cells(ctr, 13)
    filetemp = "_CD"
    ElseIf (Combo2.ListIndex = 2) Then
    costtemp = wsheet.Cells(ctr, 16)
    filetemp = "_JClear"
    ElseIf (Combo2.ListIndex = 3) Then
    costtemp = wsheet.Cells(ctr, 57)
    filetemp = "_RWL"
    ElseIf (Combo2.ListIndex = 4) Then
    costtemp = wsheet.Cells(ctr, 138) + wsheet.Cells(ctr, 57) + (wsheet.Cells(ctr, 25) + wsheet.Cells(ctr, 27) + wsheet.Cells(ctr, 29) + wsheet.Cells(ctr, 31) + wsheet.Cells(ctr, 33) + wsheet.Cells(ctr, 35) + wsheet.Cells(ctr, 37) + wsheet.Cells(ctr, 39) + wsheet.Cells(ctr, 42) + wsheet.Cells(ctr, 45) + wsheet.Cells(ctr, 48) + wsheet.Cells(ctr, 50) + wsheet.Cells(ctr, 52) + wsheet.Cells(ctr, 54)) / wsheet.Cells(ctr, 13)
    filetemp = "Ttlcst"
    Else
    costtemp = wsheet.Cells(ctr, 138)
    filetemp = "_RD"
    End If
    If (InStr(distlist, disttemp) = 0) Then
        If (distlist = "EMPTY") Then
         distlist = disttemp
        Else
         distlist = distlist & disttemp
        End If
        dist.Add CStr(costtemp), disttemp
        distrcrd.Add 1, disttemp
        serial.Add disttemp, CStr(serialno)
        serialno = serialno + 1
    Else
        costtemp = costtemp + dist(disttemp)
        inttemp = distrcrd(disttemp)
        dist.Remove (disttemp)
        distrcrd.Remove (disttemp)
        dist.Add costtemp, disttemp
        distrcrd.Add (CInt(inttemp) + 1), disttemp
    End If
End If
costtemp = 0
End If
Next ctr
Form2.Hide
 ReDim x(1, distrcrd.Count)
 Dim row As Integer
 row = 1
For inttemp = 1 To serialno - 1
  disttemp = serial(CStr(inttemp))
  x(0, row) = disttemp
  x(1, row) = dist(disttemp) / distrcrd(disttemp)
  row = row + 1
 Next inttemp
 MSChart1.AllowSeriesSelection = True
 MSChart1.ChartData = x
 Form5.Show
 MSChart1.Refresh
On Error Resume Next
appexcel.AlertBeforeOverwriting = False
Select Case (Combo1.ListIndex)
    Case 0
    Case 1
    Case 2
    Case 3
    Case Else
        'intmsg = MsgBox("File for the out put couldn't be decided.. Fix it now... By default 'c:\anant\result1.xls! distall' has been selected")
    End Select
appexcel.Close (vbNo)
Set appexcel = Nothing
Set wsheet = Nothing
End Sub

Private Sub MSChart1_SeriesActivated(Series As Integer, MouseFlags As Integer, Cancel As Integer)
Dim intmsg As Integer
Dim str As String
Call Form10.frm10ld(3, CStr(x(0, Series - 1)))
Form10.Show
End Sub

