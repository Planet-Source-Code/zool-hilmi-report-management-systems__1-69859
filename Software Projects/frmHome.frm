VERSION 5.00
Begin VB.Form frmHome 
   Appearance      =   0  'Flat
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14355
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10350
   ScaleWidth      =   14355
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "OPTION MENU"
      Height          =   5895
      Left            =   8160
      TabIndex        =   4
      Top             =   1320
      Width           =   4215
      Begin VB.CommandButton cmdDeletePassword 
         Caption         =   "DELETE USER"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         TabIndex        =   7
         Top             =   4080
         Width           =   2295
      End
      Begin VB.CommandButton cmdChangePassword 
         Caption         =   "CHANGE PASSWORD"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         TabIndex        =   6
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton cmdAddPassword 
         Caption         =   "ADD NEW USER"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "INSURANCE INFORMATION"
      Height          =   5895
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
      Begin VB.CommandButton cmdReport 
         Caption         =   "REPORT GENERATOR"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   8
         Top             =   4080
         Width           =   2535
      End
      Begin VB.CommandButton cmdNonMotorInsurance 
         Appearance      =   0  'Flat
         Caption         =   "NON MOTOR INSURANCE"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   3
         Top             =   2520
         Width           =   2655
      End
      Begin VB.CommandButton cmdMotorInsurance 
         Appearance      =   0  'Flat
         Caption         =   "MOTOR INSURANCE"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE SELECT AN OPTION"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   0
      Top             =   720
      Width           =   5175
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAddPassword_Click()
Unload Me

frmAddUser.Show
End Sub

Private Sub cmdMotorInsurance_Click()
Unload Me
frmMotorInsurance.Show

End Sub

