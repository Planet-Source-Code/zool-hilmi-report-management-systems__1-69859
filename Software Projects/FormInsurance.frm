VERSION 5.00
Begin VB.Form FormInsurance 
   Caption         =   "INFORMATION FOR INSURANCE"
   ClientHeight    =   9840
   ClientLeft      =   360
   ClientTop       =   330
   ClientWidth     =   15090
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   15090
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Non Motor Insurance"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   8040
      TabIndex        =   13
      Top             =   720
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Motor Insurance"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   4935
      Begin VB.CommandButton Command3 
         Caption         =   "SEARCH"
         Height          =   495
         Left            =   1680
         TabIndex        =   16
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "DELETE"
         Height          =   495
         Left            =   2640
         TabIndex        =   15
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD RECORD"
         Height          =   495
         Left            =   840
         TabIndex        =   14
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "N.C.B."
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
         Left            =   240
         TabIndex        =   6
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Basic Premium Acces"
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
         Left            =   240
         TabIndex        =   5
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Excess"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Policy Period"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Vehicle Sum"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Policy No"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "FormInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
