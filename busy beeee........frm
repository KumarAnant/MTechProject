VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Busy Bee...."
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   FillColor       =   &H00FFFFFF&
   FontTransparent =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form2"
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   41.54
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   183.886
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   11
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   7575
      TabIndex        =   2
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "            U'r COMMAND IS BEING PROCESSED..........      Plz. Wait ! ! ! ! !      "
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1020
      TabIndex        =   0
      Top             =   360
      Width           =   7695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form2.Left = (screen.Width - Form2.Width) / 2
Form2.Top = (screen.Height - Form2.Height) / 2
End Sub
