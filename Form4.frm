VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form4 
   Caption         =   "Region"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkMode        =   1  'Source
   ScaleHeight     =   5625
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3495
      Left            =   360
      OleObjectBlob   =   "Form4.frx":0000
      TabIndex        =   3
      Top             =   720
      Width           =   6255
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      ItemData        =   "Form4.frx":2511
      Left            =   5760
      List            =   "Form4.frx":2524
      TabIndex        =   2
      Text            =   "Type of Work"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<-- &Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   4920
      Width           =   855
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strrgn As String, strdst As String
Private x(1, 2) As Variant
Private Sub Combo1_Click()
Call Command
End Sub
Private Sub Command1_Click()
End
End Sub
Private Sub Command()
Dim appexcel As excel.Workbook
Dim wsheet As Worksheet
Dim region As New Collection
Dim intmsg As Integer, ctr As Long, inttemp As Integer
Dim uppopln As Long, utpopln As Integer, bhpopln As Integer
Dim strtemp As String
Dim cost As Long, costtemp As Long
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
MSChart1.AllowSeriesSelection = True
Form2.Show
Form4.Hide
inttemp = 0
uppopln = utpopln = bhpopln = 0
region.Add 0, "UP"
region.Add 0, "UT"
region.Add 0, "BR"
If (Combo1.ListIndex <> 0 And Combo1.ListIndex <> 1 And Combo1.ListIndex And 2) Then
intmsg = MsgBox("By Default Combo.1 index has been selected as 'ROAD CONSTRUCTION' ")
End If
For ctr = 2 To (Form1.total - 2) Step 1
    cost = 0
    costtemp = 0
    strtemp = ""
    If (Combo1.ListIndex = 0) Then
    cost = wsheet.Cells(ctr, 138)
    ElseIf (Combo1.ListIndex = 1) Then
    cost = (wsheet.Cells(ctr, 25) + wsheet.Cells(ctr, 27) + wsheet.Cells(ctr, 29) + wsheet.Cells(ctr, 31) + wsheet.Cells(ctr, 33) + wsheet.Cells(ctr, 35) + wsheet.Cells(ctr, 37) + wsheet.Cells(ctr, 39) + wsheet.Cells(ctr, 42) + wsheet.Cells(ctr, 45) + wsheet.Cells(ctr, 48) + wsheet.Cells(ctr, 50) + wsheet.Cells(ctr, 52) + wsheet.Cells(ctr, 54)) / wsheet.Cells(ctr, 13)
    ElseIf (Combo1.ListIndex = 2) Then
    cost = wsheet.Cells(ctr, 16)
    ElseIf (Combo1.ListIndex = 3) Then
    cost = wsheet.Cells(ctr, 57)
    ElseIf (Combo1.ListIndex = 4) Then
    cost = wsheet.Cells(ctr, 138) + wsheet.Cells(ctr, 57) + (wsheet.Cells(ctr, 25) + wsheet.Cells(ctr, 27) + wsheet.Cells(ctr, 29) + wsheet.Cells(ctr, 31) + wsheet.Cells(ctr, 33) + wsheet.Cells(ctr, 35) + wsheet.Cells(ctr, 37) + wsheet.Cells(ctr, 39) + wsheet.Cells(ctr, 42) + wsheet.Cells(ctr, 45) + wsheet.Cells(ctr, 48) + wsheet.Cells(ctr, 50) + wsheet.Cells(ctr, 52) + wsheet.Cells(ctr, 54)) / wsheet.Cells(ctr, 13)
    Else
    cost = wsheet.Cells(ctr, 138)
    End If
    If wsheet.Cells(ctr, 4) = "UP" Or wsheet.Cells(ctr, 4) = "up" Or wsheet.Cells(ctr, 4) = "Up" Or wsheet.Cells(ctr, 4) = "Up" Then
    strtemp = "UP"
    ElseIf wsheet.Cells(ctr, 4) = "UT" Or wsheet.Cells(ctr, 4) = "ut" Or wsheet.Cells(ctr, 4) = "UA" Or wsheet.Cells(ctr, 4) = "Ua" Then
    strtemp = "UT"
    ElseIf wsheet.Cells(ctr, 4) = "BR" Or wsheet.Cells(ctr, 4) = "br" Or wsheet.Cells(ctr, 4) = "Br" Or wsheet.Cells(ctr, 4) = "bR" Then
    strtemp = "BR"
    Else
    strtemp = "Input data at" & (ctr + 2) & "," & "4 is undectable...."
    intmsg = MsgBox(strtemp, vbCritical, "HELP ??")
    End
    End If
    Form2.ProgressBar1.Value = CByte(ctr * 100 / Form1.total)
    Form2.Label2.Caption = CByte(ctr * 100 / Form1.total) & " % Completed  "
    Form2.Label2.Refresh
    costtemp = CLng(region(strtemp))
    cost = cost + costtemp
    region.Remove (strtemp)
    region.Add cost, strtemp
If (strtemp = "UP") Then
    uppopln = uppopln + 1
ElseIf (strtemp = "UT") Then
    utpopln = utpopln + 1
ElseIf (strtemp = "BR") Then
    bhpopln = bhpopln + 1
Else
    intmsg = MsgBox("Wrong case selection...")
End If
cost = 0
costtemp = 0
Next ctr
Form2.Hide
Form4.Show
If (uppopln <> 0) Then
x(1, 0) = CLng(region("UP")) / uppopln
Else
intmsg = MsgBox("No. of UP data set is: 0")
End If
If (utpopln <> 0) Then
x(1, 1) = CLng(region("UT")) / utpopln
Else
intmsg = MsgBox("No. of UT data set is: 0")
End If
If (bhpopln <> 0) Then
x(1, 2) = CLng(region("BR")) / bhpopln
Else
intmsg = MsgBox("No. of BH data set is: 0")
End If
MSChart1.Visible = True
x(0, 0) = "Uttar Pradesh"
x(0, 1) = "Uttranchal"
x(0, 2) = "Bihar"
MSChart1.ChartData = x
On Error Resume Next
appexcel.AlertBeforeOverwriting = False
If (Combo1.ListIndex = 0) Then
MSChart1.TitleText = "State Wise Cost of Construction Comparison"
ElseIf (Combo1.ListIndex = 1) Then
MSChart1.TitleText = "State Wise Cost of C/D Structure Construction Comparison"
ElseIf (Combo1.ListIndex = 2) Then
MSChart1.TitleText = "State Wise Cost Jungle Clearing Comparison"
Else
MSChart1.TitleText = "State Wise Cost of Construction Comparison"
End If
appexcel.Close (vbNo)
Set appexcel = Nothing
Set wsheet = Nothing
End Sub

Private Sub Command2_Click()
Form1.Show
Form1.Timer1.Interval = 1250
Form4.Hide
Form1.WindowState = 2
End Sub

Private Sub Form_Load()
Dim appexcel As excel.Workbook
Dim inti As Integer, ctr As Integer, rowno As Byte, fnum As Integer
Dim strtemp As String
Dim wsheet As Worksheet
Dim cntr As Integer
MSChart1.Visible = False
Form4.Width = screen.Width
Form4.Height = screen.Height
Form4.WindowState = 2
MSChart1.Height = screen.Height - 2000
MSChart1.Top = Form1.Top
MSChart1.Width = screen.Width
MSChart1.Left = 0
Command2.Top = screen.Height * 0.85
Command2.Left = (screen.Width - Command1.Width - screen.Width * 0.15) / 2
Command1.Left = (screen.Width - Command2.Width + screen.Width * 0.15) / 2
Command1.Top = screen.Height * 0.85
Combo1.Top = screen.Height * 0.85
Combo1.Left = screen.Width * 0.95 - Combo1.Width
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
fnum = FreeFile
Dim wBook As Workbook
strdst = ""
rowno = 2
Form2.Hide
Form4.Show
On Error Resume Next
appexcel.AlertBeforeOverwriting = False
appexcel.Close
Set appexcel = Nothing
Set wsheet = Nothing
Call Command
End Sub
Private Sub MSChart1_SeriesActivated(Series As Integer, MouseFlags As Integer, Cancel As Integer)
Dim intmsg As Integer
Dim str As String
Call Form10.frm10ld(4, CStr(x(0, Series - 1)))
Form10.Show
End Sub
