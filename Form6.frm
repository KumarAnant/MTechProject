VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form6 
   Caption         =   "COMPARISON"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form6"
   ScaleHeight     =   5865
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<-- &Back"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   5160
      Width           =   1455
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Index           =   1
      ItemData        =   "Form6.frx":0000
      Left            =   3480
      List            =   "Form6.frx":002E
      TabIndex        =   3
      Text            =   "Select Item"
      Top             =   5160
      Width           =   2655
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "Place 2"
      Top             =   3240
      Width           =   2655
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form6.frx":011B
      Left            =   3480
      List            =   "Form6.frx":011D
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Place 1"
      Top             =   2880
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form6.frx":011F
      Left            =   3480
      List            =   "Form6.frx":0129
      TabIndex        =   0
      Text            =   "Select Level"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "PLACES  TO  BE  COMPARED"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   840
      TabIndex        =   5
      Top             =   2520
      Width           =   5415
   End
   Begin VB.Frame Frame3 
      Caption         =   "ITEM  TO  BE  COMPARED"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   840
      TabIndex        =   6
      Top             =   4800
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      Caption         =   "LEVEL  OF  COMPARISON"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   5415
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4335
      Left            =   360
      OleObjectBlob   =   "Form6.frx":013E
      TabIndex        =   7
      Top             =   480
      Width           =   7455
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim appexcel As Workbook
Dim wsheet As Worksheet
Dim item1 As String, itemno1 As Integer
Dim item2 As String, itemno2 As Integer
Dim keystat As Boolean
Private Sub Combo1_Click()
Call secndnthrditm
Combo2.Text = "Place 1"
Combo3.Text = "Place 2"
End Sub
Private Sub Combo2_Click()
Dim ctr As Long
For ctr = 0 To Combo3.ListCount - 1
If Combo2.Text = Combo3.List(ctr) Then
    itemno1 = ctr
    Combo3.RemoveItem (ctr)
    Exit For
End If
Next ctr
If (IsEmpty(item1) = False And IsEmpty(itemno1) = False And item1 <> Combo2.Text And Combo2.Text <> "") Then
    Combo3.AddItem item1, itemno1
    item1 = Combo2.Text
End If
End Sub
Private Sub Combo3_Click()
Dim ctr As Long
For ctr = 0 To Combo2.ListCount - 1
    If Combo3.Text = Combo2.List(ctr) Then
    itemno2 = ctr
    Combo2.RemoveItem (ctr)
    Exit For
End If
Next ctr
If (IsEmpty(item2) = False And IsEmpty(itemno2) = False And item2 <> Combo3.Text And Combo3.Text <> "") Then
     Combo2.AddItem item2, itemno2
     item2 = Combo3.Text
 End If
End Sub
Private Sub Command1_Click()
If (keystat = False) Then
Me.Hide
Form1.Show
Form1.Timer1.Interval = 1250
ElseIf (keystat = True) Then
Combo1.Visible = True
Combo2.Visible = True
Combo3.Visible = True
Combo4(1).Visible = True
Frame1.Visible = True
Frame2.Visible = True
Frame3.Visible = True
Command3.Visible = True
MSChart1.Visible = False
keystat = False
End If
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Command3_Click()
Dim col As Integer, appexcel As Workbook
Dim wsheet As Worksheet
Dim item1 As Long, item2 As Long, popln1 As Integer, popln2 As Integer
Dim intmsg As Integer, ctr As Long
Dim charttitle As String
Dim x(1, 1) As Variant
Dim state As String, state2 As String
If (Combo4(1).ListIndex < 0) Then
    intmsg = MsgBox("FIRST SELECT AN ITEM FROM COMBO4 TO BE COMPARED")
    Exit Sub
End If
If (Combo1.ListIndex < 0) Then
    intmsg = MsgBox("FIRST SELECT AN ITEM FROM COMBO1 TO BE COMPARED")
    Exit Sub
End If
If (Combo2.ListIndex < 0) Then
    intmsg = MsgBox("FIRST SELECT AN ITEM FROM COMBO2 TO BE COMPARED")
    Exit Sub
End If
If (Combo3.ListIndex < 0) Then
    intmsg = MsgBox("FIRST SELECT AN ITEM FROM COMBO3 TO BE COMPARED")
    Exit Sub
End If
If ((UCase(Combo1.Text) = "STATE") And (UCase(Combo2.Text) = "BIHAR")) Then
state = "BR"
ElseIf ((UCase(Combo1.Text) = "STATE") And (UCase(Combo2.Text) = "UTTAR PRADESH")) Then
state = "UP"
ElseIf ((UCase(Combo1.Text) = "STATE") And (UCase(Combo2.Text) = "UTTRANCHAL")) Then
state = "UT"
ElseIf (UCase(Combo1.Text) = "DISTRICT") Then
state = Left(Combo2.Text, Len(Combo2.Text) - 5)
Else
intmsg = MsgBox("There is something wrong here")
End If
If ((UCase(Combo1.Text) = "STATE") And (UCase(Combo3.Text) = "BIHAR")) Then
state2 = "BR"
ElseIf ((UCase(Combo1.Text) = "STATE") And (UCase(Combo3.Text) = "UTTAR PRADESH")) Then
state2 = "UP"
ElseIf ((UCase(Combo1.Text) = "STATE") And (UCase(Combo3.Text) = "UTTRANCHAL")) Then
state2 = "UT"
ElseIf (UCase(Combo1.Text) = "DISTRICT") Then
state2 = Left(Combo3.Text, Len(Combo3.Text) - 5)
Else
intmsg = MsgBox("There is something wrong here")
End If
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
If Combo1.Text = "State" Then
    col = 4
ElseIf Combo1.Text = "District" Then
    col = 3
Else
    intmsg = MsgBox("Wrong combo box (1) selection")
End If
Combo1.Visible = False
Combo2.Visible = False
Combo3.Visible = False
Combo4(1).Visible = False
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Command3.Visible = False
MSChart1.Visible = True
keystat = True
item1 = 0
item2 = 0
popln1 = 0
popln2 = 0
Form2.Show
Form6.Hide
For ctr = 2 To Form1.total - 2
    If UCase(wsheet.Cells(ctr, col)) = UCase(state) Then
        popln1 = popln1 + 1
        Select Case Combo4(1).ListIndex
            Case 0
            item1 = item1 + wsheet.Cells(ctr, 11)
            Case 1
            item1 = item1 + wsheet.Cells(ctr, 16)
            Case 2
            item1 = item1 + wsheet.Cells(ctr, 18)
            Case 3
            item1 = item1 + wsheet.Cells(ctr, 20)
            Case 4
            item1 = item1 + (wsheet.Cells(ctr, 24) + wsheet.Cells(ctr, 26) + wsheet.Cells(ctr, 28) + wsheet.Cells(ctr, 30) + wsheet.Cells(ctr, 32) + wsheet.Cells(ctr, 34) + wsheet.Cells(ctr, 36) + wsheet.Cells(ctr, 38) + wsheet.Cells(ctr, 40) + wsheet.Cells(ctr, 43) + wsheet.Cells(ctr, 46) + wsheet.Cells(ctr, 51)) / wsheet.Cells(ctr, 13)
            Case 5
            item1 = item1 + (wsheet.Cells(ctr, 25) + wsheet.Cells(ctr, 27) + wsheet.Cells(ctr, 29) + wsheet.Cells(ctr, 31) + wsheet.Cells(ctr, 33) + wsheet.Cells(ctr, 35) + wsheet.Cells(ctr, 37) + wsheet.Cells(ctr, 39) + wsheet.Cells(ctr, 42) + wsheet.Cells(ctr, 45) + wsheet.Cells(ctr, 48) + wsheet.Cells(ctr, 50) + wsheet.Cells(ctr, 52)) / wsheet.Cells(ctr, 13)
            Case 6
            item1 = item1 + wsheet.Cells(ctr, 61)
            Case 7
            item1 = item1 + wsheet.Cells(ctr, 66)
            Case 8
            item1 = item1 + wsheet.Cells(ctr, 76)
            Case 9
            item1 = item1 + wsheet.Cells(ctr, 81)
            Case 10
            item1 = item1 + wsheet.Cells(ctr, 82)
            Case 11
            item1 = item1 + wsheet.Cells(ctr, 93) + wsheet.Cells(ctr, 100) + wsheet.Cells(ctr, 108) + wsheet.Cells(ctr, 116) + wsheet.Cells(ctr, 124) + wsheet.Cells(ctr, 132)
            Case 12
            item1 = item1 + wsheet.Cells(ctr, 138)
            Case 13
            item1 = item1 + wsheet.Cells(ctr, 68) + wsheet.Cells(ctr, 83) + wsheet.Cells(ctr, 101) + wsheet.Cells(ctr, 110) + wsheet.Cells(ctr, 118) + wsheet.Cells(ctr, 126) + wsheet.Cells(ctr, 134)
            Case Else
            intmsg = MsgBox("ITEM COULDN'T BE FOUND")
        End Select
    End If
        If UCase(wsheet.Cells(ctr, col)) = UCase(state2) Then
        popln2 = popln2 + 1
        Select Case Combo4(1).ListIndex
            Case 0
            item2 = item2 + wsheet.Cells(ctr, 11)
            Case 1
            item2 = item2 + wsheet.Cells(ctr, 16)
            Case 2
            item2 = item2 + wsheet.Cells(ctr, 18)
            Case 3
            item2 = item2 + wsheet.Cells(ctr, 20)
            Case 4
            item2 = item2 + (wsheet.Cells(ctr, 24) + wsheet.Cells(ctr, 26) + wsheet.Cells(ctr, 28) + wsheet.Cells(ctr, 30) + wsheet.Cells(ctr, 32) + wsheet.Cells(ctr, 34) + wsheet.Cells(ctr, 36) + wsheet.Cells(ctr, 38) + wsheet.Cells(ctr, 40) + wsheet.Cells(ctr, 43) + wsheet.Cells(ctr, 46) + wsheet.Cells(ctr, 51)) / wsheet.Cells(ctr, 13)
            Case 5
            item2 = item2 + (wsheet.Cells(ctr, 25) + wsheet.Cells(ctr, 27) + wsheet.Cells(ctr, 29) + wsheet.Cells(ctr, 31) + wsheet.Cells(ctr, 33) + wsheet.Cells(ctr, 35) + wsheet.Cells(ctr, 37) + wsheet.Cells(ctr, 39) + wsheet.Cells(ctr, 42) + wsheet.Cells(ctr, 45) + wsheet.Cells(ctr, 48) + wsheet.Cells(ctr, 50) + wsheet.Cells(ctr, 52)) / wsheet.Cells(ctr, 13)
            Case 6
            item2 = item2 + wsheet.Cells(ctr, 61)
            Case 7
            item2 = item2 + wsheet.Cells(ctr, 66)
            Case 8
            item2 = item2 + wsheet.Cells(ctr, 76)
            Case 9
            item2 = item2 + wsheet.Cells(ctr, 81)
            Case 10
            item1 = item2 + wsheet.Cells(ctr, 82)
            Case 11
            item2 = item2 + wsheet.Cells(ctr, 93) + wsheet.Cells(ctr, 100) + wsheet.Cells(ctr, 108) + wsheet.Cells(ctr, 116) + wsheet.Cells(ctr, 124) + wsheet.Cells(ctr, 132)
            Case 12
            item2 = item2 + wsheet.Cells(ctr, 138)
            Case 13
            item2 = item2 + wsheet.Cells(ctr, 68) + wsheet.Cells(ctr, 83) + wsheet.Cells(ctr, 101) + wsheet.Cells(ctr, 110) + wsheet.Cells(ctr, 118) + wsheet.Cells(ctr, 126) + wsheet.Cells(ctr, 134)
            Case Else
            intmsg = MsgBox("ITEM COULN'T BE FOUND")
        End Select
    End If
    Form2.Label2 = CByte(ctr * 100 / Form1.total) & " % Completed"
    Form2.Label2.Refresh
    Form2.ProgressBar1.Value = (ctr * 100 / Form1.total)
    Next ctr
Select Case Combo4(1).ListIndex
   Case 0
   charttitle = "CBR"
   Case 1
   charttitle = "COST OF JUNGLE CLEARING"
   Case 2
    charttitle = "EARTHWORK (Volm)"
    Case 3
    charttitle = "EARTHWORK (cost)"
    Case 4
    charttitle = "NO. OF C/D STR"
    Case 5
    charttitle = "C/D COST"
    Case 6
    charttitle = "SUB-BASE (Thickness)"
    Case 7
    charttitle = "SUB-BASE (Cost)"
    Case 8
    charttitle = "BASE (Thickness)"
    Case 9
    charttitle = "BASE (cost)"
    Case 10
    charttitle = "LEAD FROM QUARY"
    Case 11
    charttitle = "SURFACE COURSE COST"
    Case 12
    charttitle = "Total Cost"
    Case 13
    charttitle = "TRANSPORTATION cost"
    End Select
    MSChart1.TitleText = "Comparison for " & charttitle & " in " & Combo2.Text & " & " & Combo3.Text
    x(0, 0) = Combo2.Text
    x(0, 1) = Combo3.Text
    If (popln1 <> 0) Then
    x(1, 0) = item1 / popln1
    Else
    intmsg = MsgBox("There is no data set for " & charttitle & " in " & Combo2.Text)
    End If
    If (popln2 <> 0) Then
    x(1, 1) = item2 / popln2
    Else
    intmsg = MsgBox("There is no data set for " & charttitle & " in " & Combo3.Text)
    End If
    MSChart1.ChartData = x
Form2.Hide
Form6.Show
On Error Resume Next
appexcel.Close (vbNo)
Set appexcel = Nothing
Set wsheet = Nothing
End Sub
Private Sub Form_Load()
Command1.Top = screen.Height * 0.85
Command1.Left = screen.Width / 2 - 3375
Command2.Left = screen.Width / 2 + 2000
Command3.Left = screen.Width / 2 - 700
Command2.Top = screen.Height * 0.85
Command3.Top = screen.Height * 0.85
Command3.Caption = "Show &Graph"
Form6.WindowState = 2
MSChart1.Width = screen.Width
MSChart1.Height = screen.Height - 2000
MSChart1.Left = 0
MSChart1.Visible = False
On Error Resume Next
appexcel.Close (vbNo)
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
Combo1.ListIndex = 0
Combo1.Text = Combo1.List(0)
Call secndnthrditm
Combo2.Text = "Place 1"
Combo3.Text = "Place 2"
keystat = False
End Sub
Private Sub secndnthrditm()
Dim apxl As Workbook, wsheet As Worksheet
Dim ctr As Long, intmsg As Integer
Dim dist As String, disttemp As String
Set apxl = GetObject(Form1.filepath)
Set wsheet = apxl.Sheets(1)
If Combo1.ListIndex = 0 Then
    Combo2.Clear
    Combo3.Clear
    Combo3.AddItem "Uttar Pradesh"
    Combo2.AddItem "Uttar Pradesh"
    Combo3.AddItem "Uttranchal"
    Combo2.AddItem "Uttranchal"
    Combo3.AddItem "Bihar"
    Combo2.AddItem "Bihar"
ElseIf Combo1.ListIndex = 1 Then
    Combo2.Clear
    Combo3.Clear
    For ctr = 2 To Form1.total - 1 Step 1
    disttemp = wsheet.Cells(ctr, 3)
    If ((IsEmpty(dist) = True) Or (InStr(dist, disttemp) = 0)) Then
        Combo2.AddItem disttemp & " (" & wsheet.Cells(ctr, 4) & ")"
        Combo3.AddItem disttemp & " (" & wsheet.Cells(ctr, 4) & ")"
        dist = dist & disttemp
    End If
    Next ctr
Else
    intmsg = MsgBox("First select the level...from 1st combo box")
End If
item1 = Empty
item2 = Empty
itemno1 = Empty
itemno2 = Empty
On Error Resume Next
apxl.Close (vbNo)
Set apxl = Nothing
Set wsheet = Nothing
End Sub
