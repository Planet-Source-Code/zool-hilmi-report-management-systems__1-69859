VERSION 5.00
Begin VB.Form FormLPKP 
   Caption         =   "Information for LPKP"
   ClientHeight    =   6360
   ClientLeft      =   1110
   ClientTop       =   -60
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   9210
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   6
      Top             =   4200
      Width           =   5655
      Begin VB.CommandButton Command3 
         Caption         =   "Remove Info"
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update Info"
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add New Info"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   5
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Tempoh Lesen"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "No Rujukan Fail"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "No:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "FormLPKP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
