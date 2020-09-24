VERSION 5.00
Begin VB.Form FormJPJ 
   Caption         =   "Information For Jabatan Pengangkutan Jalan"
   ClientHeight    =   8745
   ClientLeft      =   1350
   ClientTop       =   -180
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   11235
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   2640
      TabIndex        =   26
      Top             =   6720
      Width           =   6135
      Begin VB.CommandButton Command3 
         Caption         =   "Remove Vehicle Info"
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
         Left            =   4200
         TabIndex        =   29
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update Vehicle Info"
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
         Left            =   2160
         TabIndex        =   28
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add New Vehicle"
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
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.TextBox Text13 
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
      Left            =   8400
      TabIndex        =   25
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text12 
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
      Left            =   8400
      TabIndex        =   24
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text11 
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
      Left            =   8400
      TabIndex        =   23
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text10 
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
      Left            =   8400
      TabIndex        =   22
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text9 
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
      Left            =   8400
      TabIndex        =   21
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text8 
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
      Left            =   8400
      TabIndex        =   20
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text7 
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
      Left            =   3840
      TabIndex        =   19
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox Text6 
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
      Left            =   3840
      TabIndex        =   18
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text5 
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
      Left            =   3840
      TabIndex        =   17
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text4 
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
      Left            =   3840
      TabIndex        =   16
      Top             =   3120
      Width           =   1935
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
      Left            =   3840
      TabIndex        =   15
      Top             =   2400
      Width           =   1935
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
      Left            =   3840
      TabIndex        =   14
      Top             =   1680
      Width           =   1935
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
      Left            =   3840
      TabIndex        =   13
      Top             =   960
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10800
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label13 
      Caption         =   "Kadar Lesen"
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
      Left            =   6240
      TabIndex        =   12
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "D.T.M"
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
      Left            =   6240
      TabIndex        =   11
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "B.D.M / B.A.K"
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
      Left            =   6240
      TabIndex        =   10
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "Tarikh Pendaftaran"
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
      Left            =   6240
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Tahub Di Buat"
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
      Left            =   6240
      TabIndex        =   8
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Jenis Badan"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Bahan Bakar"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Keupayaan Model"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Nama Model"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Buatan"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "No.Casis"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "No.Enjin"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "No Siri"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "FormJPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
