VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FormInsurance 
   Caption         =   "Insurance Information"
   ClientHeight    =   8385
   ClientLeft      =   660
   ClientTop       =   150
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   9720
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab Ta 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Motor Insurance"
      TabPicture(0)   =   "Form_Insurance.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Non-Motor Insurance"
      TabPicture(1)   =   "Form_Insurance.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3360
         TabIndex        =   12
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3360
         TabIndex        =   11
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3360
         TabIndex        =   10
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3360
         TabIndex        =   9
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3360
         TabIndex        =   8
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3360
         TabIndex        =   7
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "N.C.B"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Basic Premium Access"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Excess"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Vehicle Sum Insured"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Policy Period"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Policy No."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   960
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FormInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
