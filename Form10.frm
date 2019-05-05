VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H80000013&
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808000&
   LinkTopic       =   "Form10"
   ScaleHeight     =   8850
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H80000013&
      Caption         =   "Conveyance"
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   840
      TabIndex        =   29
      Top             =   6720
      Width           =   8055
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label30"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   33
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Label29"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   32
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Conveyance Cost (Rs.):"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   31
         Top             =   720
         Width           =   3045
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Query Distance (Km):"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   30
         Top             =   360
         Width           =   2190
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "Cost Components"
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   840
      TabIndex        =   2
      Top             =   3360
      Width           =   8055
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label26"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   16
         Top             =   2640
         Width           =   840
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         TabIndex        =   15
         Top             =   2265
         Width           =   840
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label24"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2520
         TabIndex        =   14
         Top             =   1875
         Width           =   840
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label23"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2640
         TabIndex        =   13
         Top             =   1500
         Width           =   840
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label22"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   12
         Top             =   1125
         Width           =   840
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label21"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   11
         Top             =   735
         Width           =   840
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   10
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Site Clearing (Rs.):"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   1965
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C/D Str (Rs.):"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   8
         Top             =   1125
         Width           =   1365
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Earthwork (Rs):"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   7
         Top             =   1500
         Width           =   1590
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Base (Rs.)"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   6
         Top             =   1875
         Width           =   1590
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retaining Wall (Rs.):"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   5
         Top             =   735
         Width           =   2145
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base (Rs.):"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   4
         Top             =   2265
         Width           =   1170
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Surface (Rs.):"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   3
         Top             =   2640
         Width           =   1425
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "O&K"
      Default         =   -1  'True
      DownPicture     =   "Form10.frx":0000
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
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Road Details"
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   8055
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Place:"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   28
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avg CBR (%):"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   27
         Top             =   720
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avg Annual Rainfall (mm):"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   26
         Top             =   1440
         Width           =   2670
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avg No. of C/D Str:"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   25
         Top             =   1800
         Width           =   1950
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avg. Cost of Const/KM (Rs.):"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   24
         Top             =   2160
         Width           =   2910
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1680
         TabIndex        =   23
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         TabIndex        =   22
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3360
         TabIndex        =   21
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. of roads Analysed:"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   840
         TabIndex        =   20
         Top             =   1080
         Width           =   2385
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Label17"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   19
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Label18"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   18
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label19"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3840
         TabIndex        =   17
         Top             =   2160
         Width           =   840
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Form10.Hide
End Sub
Private Sub Form_Load()
Command1.Left = (Form10.Width - Command1.Width) / 2
Form10.Top = (screen.Height - Form10.Height) / 2 - 100
Form10.Left = (screen.Width - Form10.Width) / 2
End Sub
Public Sub frm10ld(key As Integer, name As String)
Dim ctr As Long, datapopln As Integer, intmsg As Integer, state As String, ctrstat As Long
Dim wsheet As Worksheet
Dim appexcel As excel.Workbook
Dim cbr As Integer, anlrnfl As Double, cdpopln As Double, avgcost As Long, stclrng As Long, rwall As Long, cdstr As Long, ework As Long, sbase As Long, base As Long, surface As Long
Dim qdist As Double, convcost As Long
Set appexcel = GetObject(Form1.filepath)
Set wsheet = appexcel.Sheets(1)
Form10.Caption = name
If (key = 4) Then
    Label14.Caption = name
End If
If (UCase(name) = "UTTAR PRADESH") Then
    name = "UP"
ElseIf ((UCase(name) = "UTTRANCHAL")) Then
    name = "UT"
ElseIf ((UCase(name) = "BIHAR")) Then
    name = "BR"
End If
Form2.Show
datapopln = 0
For ctr = 2 To Form1.total Step 1
    Form2.ProgressBar1.Value = CByte(ctr * 100 / Form1.total)
    Form2.Label2.Caption = CByte(ctr * 100 / Form1.total) & " % Completed  "
    Form2.Label2.Refresh
    If (UCase(wsheet.Cells(ctr, key)) = UCase(name)) Then
    datapopln = datapopln + 1
    If (key = 3 And UCase(wsheet.Cells(ctr, 4) = "UP")) Then
        state = " (Uttar Pradesh)"
    ElseIf (key = 3 And UCase(wsheet.Cells(ctr, 4) = "UT")) Then
        state = " (Uttranchal)"
    ElseIf (key = 3 And UCase(wsheet.Cells(ctr, 4) = "BR")) Then
        state = " (Bihar)"
    ElseIf (key = 3) Then
        intmsg = MsgBox("There is something wrong here")
    End If
    cbr = cbr + wsheet.Cells(ctr, 11)
    anlrnfl = anlrnfl + wsheet.Cells(ctr, 12)
    cdpopln = cdpopln + (wsheet.Cells(ctr, 24) + wsheet.Cells(ctr, 26) + wsheet.Cells(ctr, 28) + wsheet.Cells(ctr, 30) + wsheet.Cells(ctr, 32) + wsheet.Cells(ctr, 34) + wsheet.Cells(ctr, 36) + wsheet.Cells(ctr, 38) + wsheet.Cells(ctr, 40) + wsheet.Cells(ctr, 43) + wsheet.Cells(ctr, 46) + wsheet.Cells(ctr, 51) + wsheet.Cells(ctr, 53)) / wsheet.Cells(ctr, 13)
    avgcost = avgcost + wsheet.Cells(ctr, 138) + (wsheet.Cells(ctr, 25) + wsheet.Cells(ctr, 27) + wsheet.Cells(ctr, 29) + wsheet.Cells(ctr, 31) + wsheet.Cells(ctr, 33) + wsheet.Cells(ctr, 35) + wsheet.Cells(ctr, 37) + wsheet.Cells(ctr, 39) + wsheet.Cells(ctr, 42) + wsheet.Cells(ctr, 45) + wsheet.Cells(ctr, 48) + wsheet.Cells(ctr, 50) + wsheet.Cells(ctr, 52) + wsheet.Cells(ctr, 54)) / wsheet.Cells(ctr, 13) + wsheet.Cells(ctr, 57)
    stclrng = stclrng + wsheet.Cells(ctr, 16)
    rwall = rwall + wsheet.Cells(ctr, 57)
    cdstr = cdstr + (wsheet.Cells(ctr, 25) + wsheet.Cells(ctr, 27) + wsheet.Cells(ctr, 29) + wsheet.Cells(ctr, 31) + wsheet.Cells(ctr, 33) + wsheet.Cells(ctr, 35) + wsheet.Cells(ctr, 37) + wsheet.Cells(ctr, 39) + wsheet.Cells(ctr, 42) + wsheet.Cells(ctr, 45) + wsheet.Cells(ctr, 48) + wsheet.Cells(ctr, 50) + wsheet.Cells(ctr, 52) + wsheet.Cells(ctr, 54)) / wsheet.Cells(ctr, 13)
    ework = ework + wsheet.Cells(ctr, 20)
    sbase = sbase + wsheet.Cells(ctr, 66)
    base = base + wsheet.Cells(ctr, 81)
    surface = surface + wsheet.Cells(ctr, 93) + wsheet.Cells(ctr, 100) + wsheet.Cells(ctr, 108) + wsheet.Cells(ctr, 116) + wsheet.Cells(ctr, 124) + wsheet.Cells(ctr, 132)
    qdist = qdist + wsheet.Cells(ctr, 151)
    convcost = convcost + wsheet.Cells(ctr, 68) + wsheet.Cells(ctr, 83) + wsheet.Cells(ctr, 102) + wsheet.Cells(ctr, 110) + wsheet.Cells(ctr, 118) + wsheet.Cells(ctr, 126) + wsheet.Cells(ctr, 134)
    ctrstat = ctr
    End If
Next ctr
If (datapopln <> 0) Then
    cbr = cbr / datapopln
    anlrnfl = anlrnfl / datapopln
    cdpopln = cdpopln / datapopln
    avgcost = avgcost / datapopln
    stclrng = stclrng / datapopln
    rwall = rwall / datapopln
    cdstr = cdstr / datapopln
    ework = ework / datapopln
    sbase = sbase / datapopln
    base = base / datapopln
    surface = surface / datapopln
    qdist = qdist / datapopln
    convcost = convcost / datapopln
    Label17.Caption = anlrnfl
    Label16 = datapopln
    Label18.Caption = cdpopln
    Label19.Caption = avgcost
    Label20.Caption = stclrng
    Label21.Caption = rwall
    Label22.Caption = cdstr
    Label23 = ework
    Label24.Caption = sbase
    Label25.Caption = base
    Label26.Caption = surface
    Label15.Caption = cbr
    Label29.Caption = qdist
    Label30.Caption = convcost
    If (key = 3) Then
    Label14.Caption = name & state
    End If
    ElseIf (datapopln = 0) Then
    intmsg = MsgBox("THERE IS SOMETHING WRONG HERE")
End If
On Error Resume Next
    appexcel.Close (vbNo)
    Set appexcel = Nothing
    Form2.Hide
End Sub
