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
      Top             =   4200
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
      Top             =   1920
      Width           =   5535
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form11.frx":0000
         Left            =   1680
         List            =   "Form11.frx":000D
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
      OleObjectBlob   =   "Form11.frx":0027
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
Option Explicit
Private intmsg As Integer, keystat As Byte, x(2, 7) As Variant, hsbase As Double, hbase As Double, msa(1) As Double, rolling As VbMsgBoxResult, VDF(2) As Double
Private traffic As Double
Private appexcel As Workbook, wsheet As Worksheet, datapopln(7) As Double
Private col As Integer
Private Sub Combo1_Click()
Dim combo2text As String, combo2texttemp As String
Dim ctr As Long, intmsg As Integer
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
Combo2.Clear
Combo2.Text = "Select Place"
If Combo1.ListIndex = 0 Then
    col = 4
ElseIf Combo1.ListIndex = 1 Then
    col = 3
ElseIf UCase(Combo1.Text) = "ALL" Then
    col = 4
    ' do nothing
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
            Combo2.AddItem "Uttranchal"
        ElseIf (col = 4 And InStr(UCase(combo2texttemp), "BR") <> 0) Then
            Combo2.AddItem "Bihar"
        ElseIf (col = 3) Then
            Combo2.AddItem combo2texttemp & " (" & wsheet.Cells(ctr, 4) & ")"
        Else
            intmsg = MsgBox("There is something wrong here")
    End If
End If
Next ctr
End Sub
Private Sub Command1_Click()
If keystat = True Then
    Frame1.Visible = True
    Frame2.Visible = True
    Combo1.Visible = True
    Combo2.Visible = True
    MSChart1.Visible = False
    Command3.Visible = True
    keystat = False
Else
Me.Hide
On Error Resume Next
appexcel.Close (vbNo)
Set appexcel = Nothing
Set wsheet = Nothing
Form1.Show
Form1.Timer1.Interval = 1250
End If
End Sub
Private Sub Command2_Click()
On Error Resume Next
appexcel.Close (vbNo)
Set appexcel = Nothing
Set wsheet = Nothing
End
End Sub
Private Sub Command3_Click()
Dim ctr As Long, mthds As Byte, texttemp As String, cbr As Double
If (Combo1.ListIndex < 0) Then
    intmsg = MsgBox("First select U'r choice from LEVEL")
Exit Sub
End If
If (UCase(Combo2.Text) = "UTTAR PRADESH") Then
    texttemp = "UP"
ElseIf (UCase(Combo2.Text) = "UTTRANCHAL") Then
    texttemp = "UT"
ElseIf (UCase(Combo2.Text) = "BIHAR") Then
    texttemp = "BR"
ElseIf (Combo1.ListIndex = 2) Then
    texttemp = "UPUTBR"
ElseIf (Combo1.ListIndex = 1) Then
    texttemp = Left(Combo2.Text, Len(Combo2.Text) - 5)
ElseIf (Combo2.ListIndex < 0) Then
    intmsg = MsgBox("       First select place.........")
    Exit Sub
Else
    intmsg = MsgBox("There is somethign wrong here")
End If
On Error GoTo errorhandler
traffic = InputBox("Input Traffic Intensity (CV/d)", "Traffic ?", 0)
rolling = MsgBox("Is the site a Rolling/Plain area.", vbYesNo)
' ******* FUNCCTION FOR CALCULATION DIFFERENT COST STARTS HERE*********************
Form2.Show
Call VDFn
Call calcmsa
For ctr = 2 To Form1.total - 2
    Form2.ProgressBar1.Value = ctr * 100 / Form1.total
    Form2.Label2.Caption = CByte(ctr * 100 / Form1.total) & " % Completed"
    Form2.Label2.Refresh
    If (InStr(UCase(texttemp), UCase(wsheet.Cells(ctr, col))) <> 0) Then
    For cbr = 2 To 10
        For mthds = 1 To 3
            Call calcht(mthds, cbr)
            If (Int(cbr) >= 8 And Int(cbr) <= 10) Then
                    x(mthds - 1, 7) = (x(mthds - 1, 7) * datapopln(7) + wsheet.Cells(ctr, 16).Value + wsheet.Cells(ctr, 20).Value + (805 * wsheet.Cells(ctr, 69).Value + 4.05 * (hsbase - 100) * wsheet.Cells(ctr, 70).Value) + (281.25 * wsheet.Cells(ctr, 85) + (hbase - 75) * 3.75 * wsheet.Cells(ctr, 84)) + wsheet.Cells(ctr, 93) + wsheet.Cells(ctr, 100) + wsheet.Cells(ctr, 108) + wsheet.Cells(ctr, 116) + wsheet.Cells(ctr, 124) + wsheet.Cells(ctr, 132)) / (datapopln(7) + 1)
                    datapopln(7) = datapopln(7) + 1
            Else
                    x(mthds - 1, cbr - 1) = (x(mthds - 1, cbr - 1) * datapopln(cbr - 1) + wsheet.Cells(ctr, 16).Value + wsheet.Cells(ctr, 20).Value + (805 * wsheet.Cells(ctr, 69).Value + 4.05 * (hsbase - 100) * wsheet.Cells(ctr, 70).Value) + (281.25 * wsheet.Cells(ctr, 85).Value + (hbase - 75) * 3.75 * wsheet.Cells(ctr, 84).Value) + wsheet.Cells(ctr, 93).Value + wsheet.Cells(ctr, 100).Value + wsheet.Cells(ctr, 108).Value + wsheet.Cells(ctr, 116).Value + wsheet.Cells(ctr, 124).Value + wsheet.Cells(ctr, 132).Value) / (datapopln(cbr - 1) + 1)
                    datapopln(cbr - 1) = datapopln(cbr - 1) + 1
            End If
        Next
    Next cbr
    End If
Next
Form2.Hide
Command3.Visible = False
keystat = True
Frame1.Visible = False
Frame2.Visible = False
Combo1.Visible = False
Combo2.Visible = False
x(0, 0) = "IRC 37-2001"
x(1, 0) = "DRAFT MANUAL"
x(2, 0) = "IRC SP-20"
MSChart1.Visible = True
MSChart1.Refresh
MSChart1.ChartData = x
MSChart1.Refresh
For cbr = 2 To 8
    For mthds = 1 To 3
        If (Int(cbr) >= 8) Then
        Else
        End If
    Next
    
Next
On Error Resume Next
Exit Sub
errorhandler:
intmsg = MsgBox("U didn't enter traffic... Program unable to do the job")
Exit Sub
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
Public Sub calcht(key As Byte, cbr As Double)
Dim ht As Double
Select Case (key)
    Case 1
    Select Case (CInt(cbr))
        Case 2
            ht = 77.022 * Log(msa(0)) + 664.23
        Case 3
            ht = 94.015 * Log(msa(0)) + 544.32
        Case 4
            ht = 97.67 * Log(msa(0)) + 470.1
        Case 5
            ht = 96.125 * Log(msa(0)) + 423.58
        Case 6
            ht = 94.765 * Log(msa(0)) + 385.1
        Case 7
            ht = 93.11 * Log(msa(0)) + 356.84
        Case 8
            ht = 90.389 * Log(msa(0)) + 279.67
        Case 9
            ht = 90.389 * Log(msa(0)) + 279.67
        Case 10
            ht = 90.389 * Log(msa(0)) + 279.67
        Case Else
            intmsg = MsgBox("There is something wrong here")
        End Select
        If (ht < 375) Then
            hsbase = 150
            hbase = 225
        Else
            hbase = 225    ' 50 has been kept as the minimum prividable thickness of Sub-base
            hsbase = ht - 225 + (1 - ((ht - 225) / 50 - CInt((ht - 225) / 50))) * 50
        End If
    Case 2
        Select Case (cbr)
            Case 2
                If (msa(1) <= 0.2) Then
                    hbase = 225
                    hsbase = 300
                ElseIf (msa(1) <= 0.4) Then
                    hbase = 225
                    hsbase = 350
                ElseIf (msa(1) <= 0.8) Then
                    hbase = 225
                    hsbase = 350
                ElseIf (msa(1) <= 1) Then
                    hbase = 225
                    hsbase = 400
                ElseIf (msa(1) <= 2) Then
                    hbase = 225
                    hsbase = 400
                Else
                    intmsg = MsgBox("The design thickness is not available for this traffic in DRAFT MANUAL.. Aborting")
                    End
                End If
                
            Case 3
                If (msa(1) <= 0.2) Then
                    hbase = 150
                    hsbase = 250
                ElseIf (msa(1) <= 0.4) Then
                    hbase = 225
                    hsbase = 225
                ElseIf (msa(1) <= 0.8) Then
                    hbase = 225
                    hsbase = 225
                ElseIf (msa(1) <= 1) Then
                    hbase = 225
                   hsbase = 225
                ElseIf (msa(1) <= 2) Then
                    hbase = 225
                    hsbase = 225
                End If
            Case 4
                If (msa(1) <= 0.2) Then
                    hbase = 75
                    hsbase = 250
                ElseIf (msa(1) <= 0.4) Then
                    hbase = 150
                    hsbase = 250
                ElseIf (msa(1) <= 0.8) Then
                    hbase = 225
                    hsbase = 250
                ElseIf (msa(1) <= 1) Then
                    hbase = 225
                    hsbase = 250
                ElseIf (msa(1) <= 2) Then
                    hbase = 225
                    hsbase = 250
                End If
            Case 5
                If (msa(1) <= 0.2) Then
                    hbase = 150
                    hsbase = 150
                ElseIf (msa(1) <= 0.4) Then
                    hbase = 150
                    hsbase = 250
                ElseIf (msa(1) <= 0.8) Then
                    hbase = 150
                    hsbase = 250
                ElseIf (msa(1) <= 1) Then
                    hbase = 150
                    hsbase = 250
                ElseIf (msa(1) <= 2) Then
                    hbase = 225
                    hsbase = 250
                End If
            Case 6
                If (msa(1) <= 0.2) Then
                    hbase = 150
                    hsbase = 150
                ElseIf (msa(1) <= 0.4) Then
                    hbase = 150
                    hsbase = 150
                ElseIf (msa(1) <= 0.8) Then
                    hbase = 150
                    hsbase = 150
                ElseIf (msa(1) <= 1) Then
                    hbase = 150
                    hsbase = 250
                ElseIf (msa(1) <= 2) Then
                    hbase = 225
                    hsbase = 250
                End If
            Case 7
                If (msa(1) <= 0.2) Then
                    hbase = 150
                    hsbase = 100
                ElseIf (msa(1) <= 0.4) Then
                    hbase = 150
                    hsbase = 100
                ElseIf (msa(1) <= 0.8) Then
                    hbase = 150
                    hsbase = 150
                ElseIf (msa(1) <= 1) Then
                    hbase = 150
                    hsbase = 150
                ElseIf (msa(1) <= 2) Then
                    hbase = 150
                    hsbase = 225
                End If
            Case Else
            If (cbr = 10 Or cbr = 8 Or cbr = 9) Then
                If (msa(1) <= 0.2) Then
                    hbase = 150
                    hsbase = 100
                ElseIf (msa(1) <= 0.4) Then
                    hbase = 150
                    hsbase = 100
                ElseIf (msa(1) <= 0.8) Then
                    hbase = 150
                    hsbase = 150
                ElseIf (msa(1) <= 1) Then
                    hbase = 150
                    hsbase = 150
                ElseIf (msa(1) <= 2) Then
                    hbase = 150
                    hsbase = 150
                End If
            End If
            End Select
    Case 3
        Select Case (cbr)
            Case 2
                If (traffic < 15) Then
                    hbase = 150
                    hsbase = 375
                ElseIf (traffic <= 45) Then
                    hbase = 150
                    hsbase = 465
                ElseIf (traffic <= 150) Then
                    hbase = 225
                    hsbase = 470
                ElseIf (traffic < 4500) Then
                    hbase = 225
                    hsbase = 555
                Else
                    intmsg = MsgBox("design is not available for this TRAFFIC in SP-20")
                    End
                End If
            Case 3
                If (traffic < 15) Then
                    hbase = 150
                    hsbase = 250
                ElseIf (traffic <= 45) Then
                    hbase = 150
                    hsbase = 265
                ElseIf (traffic <= 150) Then
                    hbase = 150
                    hsbase = 330
                ElseIf (traffic < 4500) Then
                    hbase = 225
                    hsbase = 320
                End If
            Case 4
                If (traffic < 15) Then
                    hbase = 150
                    hsbase = 125
                ElseIf (traffic <= 45) Then
                    hbase = 150
                    hsbase = 200
                ElseIf (traffic <= 150) Then
                    hbase = 150
                    hsbase = 260
                ElseIf (traffic < 4500) Then
                    hbase = 150
                    hsbase = 315
                End If
            Case 5
                If (traffic < 15) Then
                    hbase = 150
                    hsbase = 100
                ElseIf (traffic <= 45) Then
                    hbase = 150
                    hsbase = 165
                ElseIf (traffic <= 150) Then
                    hbase = 150
                    hsbase = 210
                ElseIf (traffic < 4500) Then
                    hbase = 150
                    hsbase = 260
                End If
            Case 6
                If (traffic < 15) Then
                    hbase = 150
                    hsbase = 75
                ElseIf (traffic <= 45) Then
                    hbase = 150
                    hsbase = 125
                ElseIf (traffic <= 150) Then
                    hbase = 150
                    hsbase = 175
                ElseIf (traffic < 4500) Then
                    hbase = 150
                    hsbase = 225
                End If
            Case 7
                If (traffic < 15) Then
                    hbase = 150
                    hsbase = 60
                ElseIf (traffic <= 45) Then
                    hbase = 150
                    hsbase = 115
                ElseIf (traffic <= 150) Then
                    hbase = 150
                    hsbase = 150
                ElseIf (traffic < 4500) Then
                    hbase = 150
                    hsbase = 175
                End If
            Case Else
                If (cbr = 8 Or cbr = 9 Or cbr = 10) Then
                If (traffic < 15) Then
                    hbase = 150
                    hsbase = 50
                ElseIf (traffic <= 45) Then
                    hbase = 150
                    hsbase = 70
                ElseIf (traffic <= 150) Then
                    hbase = 150
                    hsbase = 95
                ElseIf (traffic < 4500) Then
                    hbase = 150
                    hsbase = 125
                End If
            End If
        End Select
    End Select
'CBR 3....y = 94.015Ln(x) + 544.32
'CBR 4....y = 97.67Ln(x) + 470.1
'CBR 5....y = 96.125Ln(x) + 423.58
'CBR 6....y = 94.765Ln(x) + 385.1
'CBR 7....y = 93.11Ln(x) + 356.84
'CBR 8....y = 94.16Ln(x) + 323.03
'CBR 9....y = 90.432Ln(x) + 301.31
'CBR 10....y = 90.389Ln(x) + 279.67

End Sub
Public Sub calcmsa()
            msa(0) = 365 * (1.06 * 1.06 * 1.06 * 1.06 * 1.06 * 1.06 * 1.06 * 1.06 * 1.06 * 1.06 - 1) / 1000000 / 0.06 * traffic * 1 * VDF(0)
            msa(1) = 365 * (1.06 * 1.06 * 1.06 * 1.06 * 1.06 * 1.06 * 1.06 * 1.06 * 1.06 * 1.06 - 1) / 1000000 / 0.06 * traffic * 1 * VDF(1)
End Sub
Public Function VDFn()
        If (traffic < 150) Then
            VDF(0) = IIf(rolling = vbYes, 1.5, 0.5)
        ElseIf (traffic < 1500) Then
            VDF(0) = IIf(rolling = vbYes, 3.5, 1.5)
        Else
            VDF(0) = IIf(rolling = vbYes, 4.5, 2.5)
        End If
        If (traffic <= 15) Then
            VDF(1) = 0.5
        ElseIf (traffic <= 50) Then
            VDF(1) = 1
        ElseIf (traffic <= 150) Then
            VDF(1) = 1.5
        Else
            VDF(1) = 2
        End If
End Function

' CBR 2... y = 77.022Ln(x) + 664.23
'CBR 3....y = 94.015Ln(x) + 544.32
'CBR 4....y = 97.67Ln(x) + 470.1
'CBR 5....y = 96.125Ln(x) + 423.58
'CBR 6....y = 94.765Ln(x) + 385.1
'CBR 7....y = 93.11Ln(x) + 356.84
'CBR 8....y = 94.16Ln(x) + 323.03
'CBR 9....y = 90.432Ln(x) + 301.31
'CBR 10....y = 90.389Ln(x) + 279.67

