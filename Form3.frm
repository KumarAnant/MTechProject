VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ovrvew"
   ClientHeight    =   8490
   ClientLeft      =   675
   ClientTop       =   150
   ClientWidth     =   10875
   FillColor       =   &H00FFFFFF&
   FillStyle       =   2  'Horizontal Line
   FontTransparent =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   10875
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<-- &BACK  "
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5880
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "DISTRICTS"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form3.frx":0000
      Left            =   5880
      List            =   "Form3.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "STATES"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   0
      Left            =   720
      TabIndex        =   17
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   1
      Left            =   720
      TabIndex        =   16
      Top             =   1710
      Width           =   135
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   2
      Left            =   720
      TabIndex        =   15
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   3
      Left            =   720
      TabIndex        =   14
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   4
      Left            =   720
      TabIndex        =   13
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000FF&
      Height          =   135
      Index           =   5
      Left            =   720
      TabIndex        =   12
      Top             =   5280
      Width           =   135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   1080
      TabIndex        =   8
      Top             =   5160
      Width           =   9840
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1080
      TabIndex        =   7
      Top             =   4320
      Width           =   9840
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   6
      Top             =   3480
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   5
      Top             =   2520
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List ofDistricts Analysed"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6240
      TabIndex        =   10
      Top             =   2400
      Width           =   1665
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of States Analysed"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6240
      TabIndex        =   9
      Top             =   720
      Width           =   1605
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim appexcel As excel.Workbook
Private Sub Combo1_Click()
    Call Form10.frm10ld(4, UCase(Combo1.Text))
    Form10.Show
End Sub
Private Sub Combo2_Click()
    Call Form10.frm10ld(3, UCase(Left((Combo2.Text), Len(Combo2.Text) - 7)))
    Form10.Show
End Sub
Private Sub Command1_Click()
Form1.Show
Me.Hide
Form1.Timer1.Interval = 1250
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
End
End Sub

Private Sub Form_Load()
Dim intmsg As Integer, intctr As Long
Dim wsheet As Worksheet
Dim dist As String, diststr As String, strtemp As String
Dim stat As String, statemp As String
Dim costmax As Long, costmin As Long, plncost As Long, hilcost As Long
costmin = 999999999
Dim bhno As Integer, upno As Integer, utno As Integer
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
Form3.WindowState = 2
diststr = ""
Combo1.Width = 2000
Combo2.Width = 3000
Combo1.Left = Form3.Width - 2000
Label7.Left = Form3.Width - 2000
Combo2.Left = Me.Width - 2000
Label8.Left = Form3.Width - 2000
Command1.Left = (screen.Width - Command1.Width - screen.Width * 0.15) / 2
Command2.Left = (screen.Width - Command2.Width + screen.Width * 0.15) / 2
Command1.Top = screen.Height * 0.8
Command2.Top = screen.Height * 0.8
intctr = 2
costmax = 0
plncost = hilcost = 0
Form3.Hide
Form2.Show
For intctr = 2 To Form1.total - 2
    Dim costtemp As Long
    Form2.ProgressBar1.Value = (intctr * 100 / Form1.total)
    Form2.Label2.Caption = CByte(intctr * 100 / Form1.total) & " % Completed"
    Form2.Label2.Refresh
    stat = wsheet.Cells(intctr, 4).Value
    dist = wsheet.Cells(intctr, 3).Value & " ( " & stat & " )"
    If (InStr(diststr, dist) = 0) Then
        Combo2.AddItem (dist)
        diststr = diststr & dist
    End If
    If (UCase(wsheet.Cells(intctr, 4).Value) = "BR") Then
        stat = "BIHAR"
    ElseIf (UCase(wsheet.Cells(intctr, 4).Value) = "UP") Then
        stat = "UTTAR PRADESH"
    ElseIf (UCase(wsheet.Cells(intctr, 4).Value) = "UT") Then
        stat = "UTTRANCHAL"
    End If
    If (InStr(statemp, stat) = 0) Then
        Combo1.AddItem stat
        statemp = statemp & stat
    End If
    costtemp = wsheet.Cells(intctr, 138)
    If (costtemp > costmax) Then
    costmax = costtemp
    End If
    If (costtemp < costmin And costtemp <> 0) Then
    costmin = costtemp
    End If
    If (wsheet.Cells(intctr, 4) = "UT") Then
        hilcost = hilcost + costtemp
        utno = utno + 1
    ElseIf (wsheet.Cells(intctr, 4) = "UP") Then
        plncost = plncost + costtemp
        upno = upno + 1
    ElseIf (wsheet.Cells(intctr, 4) = "BR") Then
        plncost = plncost + costtemp
        bhno = bhno + 1
    Else
        intmsg = MsgBox("Wrong data at" & intctr & " , " & 4)
    End If
Next intctr
Label1.Caption = "Total of " & intctr - 1 & " no. of proposed road analysed"
Label2.Caption = "Maximum cost of construction of road analysed (Rs/km): " & costmax
Label3.Caption = "Minimum cost of construction of road analysed (Rs/km): " & costmin
strtemp = CStr(CByte(((CDbl(hilcost) * CDbl(bhno + upno)) / CDbl(utno) / CDbl(plncost) - 1) * 100))
Label4.Caption = "Hilly Region is 33 % Costlier for road construction"
Label5.Caption = "The average cost of construction of rural roads under PMGSY in Bihar, Uttar Pradesh and Uttaranchal are Rs 20.7 Lakhs, 19.2 Lakhs and 26.6 Lakhs respectively"
Label6.Caption = "Cost of construction of Retaining Wall per Km in hilly area is Rs 2.47 lakhs whereas this requirement is almost nil in plain areas."
On Error GoTo errorhandler
Exit Sub
errorhandler:
intmsg = MsgBox("Won't Work....")
End Sub

