VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   2640
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2007"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "License To: Quality Concrete Sdn Bhd"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Version:1.1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "INFORMATION SYSTEMS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "VEHICLE "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "QUALITY CONCRETE "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   4215
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
PB.Min = 1
PB.Max = 200
PB.Value = 1
End Sub


Private Sub Timer1_Timer()
PB.Value = PB.Value + 10
If PB.Value > 188 Then
    
    frmLogin.Show
    Unload Me
End If
End Sub
