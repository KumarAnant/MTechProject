VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "An Analytical Study of roads under PMGSY "
   ClientHeight    =   6900
   ClientLeft      =   1995
   ClientTop       =   1140
   ClientWidth     =   8385
   ClipControls    =   0   'False
   FillStyle       =   2  'Horizontal Line
   BeginProperty Font 
      Name            =   "Bookman Old Style"
      Size            =   18
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   8385
   Begin VB.DriveListBox Drive1 
      Height          =   555
      Left            =   600
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   3600
      Top             =   2760
   End
   Begin VB.CommandButton Command3 
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
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "AN ANALYTICAL STUDY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   8415
   End
   Begin VB.Menu gnrl 
      Caption         =   "&General"
      Begin VB.Menu ovrvew 
         Caption         =   "Over &View"
         Shortcut        =   {F3}
      End
      Begin VB.Menu vldn 
         Caption         =   "Data chec&k"
      End
      Begin VB.Menu dgnmthd 
         Caption         =   "Design Method"
      End
   End
   Begin VB.Menu location 
      Caption         =   "&Location"
      Begin VB.Menu region 
         Caption         =   "&Region"
      End
      Begin VB.Menu district 
         Caption         =   "&District"
      End
   End
   Begin VB.Menu c_compnt 
      Caption         =   "&Cost-Component"
      Begin VB.Menu matnlab 
         Caption         =   "Material && Labour"
      End
      Begin VB.Menu layer 
         Caption         =   "Laye&rs && C/D"
      End
   End
   Begin VB.Menu infrence 
      Caption         =   "In&ference"
      Begin VB.Menu compare 
         Caption         =   "Com&pare"
      End
      Begin VB.Menu Lfctr 
         Caption         =   "Local &Factors"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim str As String, str2 As String
Dim inti As Integer, intmsg As Integer
Public total As Long, filepath As String
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Sub compare_Click()
Form1.Hide
Form6.Show
Form1.Timer1.Interval = 0
End Sub
Private Sub dgnmthd_Click()
Form1.Hide
Form1.Timer1.Interval = 0
FORM11.Show
End Sub
Private Sub district_Click()
Form1.Hide
Form1.Timer1.Interval = 0
Form5.Show
End Sub
Private Sub Form_Load()
Dim cntr As Long
filepath = App.Path & "\Data\Anant2.xls"
If (GetDriveType(Left(filepath, 2)) <> 5) Then
    intmsg = MsgBox("         Run Program from Correct CD !!          ", vbCritical, " Error  Loading  Program.....")
    End
End If
Form1.Width = screen.Width
Form1.Left = 0
Form1.Height = screen.Height - 300
Form1.Top = 0

Set appexcel = CreateObject("excel.application")
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
On Error GoTo errorhandler
Form1.Hide
frmSplash.Show
str = "I M MAD"
cntr = 2
total = 2
frmSplash.Refresh
Do While (str <> "")
        str = wsheet.Cells(total, 4).Text
    total = total + 1
    frmSplash.ProgressBar1.Value = total * 100 / 510
Loop
frmSplash.ProgressBar1.Value = 99
frmSplash.Hide
Form1.Show
Command3.Top = Form1.Top + Form1.Height - 1070 - Command3.Height
Command3.Left = (Form1.Width - Command3.Width) / 2
Label1.Width = Form1.Width - 1000
Label1.Left = Form1.Left + 500
Label1.Height = screen.Height - 2250
Label1.Top = Form1.Top + 400
Label1.Caption = "                                                                                                                                                                                                                                                                                                                     AN  ANALYTICAL STUDY                               ON EFFECT OF REGIONAL VARIATION     ON COST OF RURAL ROADS UNDER PMGSY"
Label1.FontSize = 32
Label1.FontBold = True
Set appexcel = Nothing
Exit Sub
errorhandler:
inti = MsgBox("SORRY..... WON'T WORK...")
End
End Sub
Private Sub Command3_Click()
End
End Sub
Private Sub layer_Click()
Form9.Show
Form1.Hide
Timer1.Interval = 0
End Sub
Private Sub Lfctr_Click()
Form7.Show
Form1.Hide
Form1.Timer1.Interval = 0
End Sub

Private Sub matnlab_Click()
Form8.Show
Form1.Hide
Timer1.Interval = 0
End Sub
Private Sub ovrvew_Click()
Timer1.Interval = 0
Form1.Hide
Form3.Show
Form3.WindowState = 2
End Sub
Private Sub region_Click()
Timer1.Interval = 0
Form1.Hide
Form1.WindowState = 1
Form4.Show
End Sub
Private Sub Timer1_Timer()
Form1.BackColor = RGB(Rnd() * 255, Rnd() * 255, Rnd() * 255)
End Sub
