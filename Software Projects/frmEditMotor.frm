VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditMotor 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   165
   ClientWidth     =   15195
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9855
   ScaleWidth      =   15195
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboRegistration 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   5280
      MouseIcon       =   "frmEditMotor.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmEditMotor.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Edit record"
      Top             =   8040
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   1335
      Left            =   4920
      TabIndex        =   60
      Top             =   7680
      Width           =   3735
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   1440
         MouseIcon       =   "frmEditMotor.frx":06F7
         MousePointer    =   99  'Custom
         Picture         =   "frmEditMotor.frx":0849
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Save record"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   2400
         MouseIcon       =   "frmEditMotor.frx":0DE1
         MousePointer    =   99  'Custom
         Picture         =   "frmEditMotor.frx":0F33
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Cancel"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   64
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   63
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   62
         Top             =   960
         Width           =   615
      End
   End
   Begin TabDlg.SSTab SSTabMotorInsurance 
      Height          =   5775
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Green Card ( JPJ )"
      TabPicture(0)   =   "frmEditMotor.frx":14B3
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Line1(0)"
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(2)=   "Label17"
      Tab(0).Control(3)=   "Label16"
      Tab(0).Control(4)=   "Label15"
      Tab(0).Control(5)=   "Label14"
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(7)=   "Label12"
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(10)=   "Label9"
      Tab(0).Control(11)=   "Label6"
      Tab(0).Control(12)=   "Label5"
      Tab(0).Control(13)=   "Label4"
      Tab(0).Control(14)=   "Label3"
      Tab(0).Control(15)=   "Label2"
      Tab(0).Control(16)=   "Label1"
      Tab(0).Control(17)=   "Label30"
      Tab(0).Control(18)=   "Label38"
      Tab(0).Control(19)=   "txtStatusJPJ"
      Tab(0).Control(20)=   "txtSeriesJPJ"
      Tab(0).Control(21)=   "txtDriveJPJ"
      Tab(0).Control(22)=   "txtBTMJPJ"
      Tab(0).Control(23)=   "txtBDMBGKJPJ"
      Tab(0).Control(24)=   "txtLicenseDuration"
      Tab(0).Control(25)=   "txtYearRegisJPJ"
      Tab(0).Control(26)=   "txtYearManuJPJ"
      Tab(0).Control(27)=   "txtEnginePowerJPJ"
      Tab(0).Control(28)=   "txtModelJPJ"
      Tab(0).Control(29)=   "txtManufacturerJPJ"
      Tab(0).Control(30)=   "txtCasiJPJ"
      Tab(0).Control(31)=   "txtEngineJPJ"
      Tab(0).Control(32)=   "cboFuelTypeJPJ"
      Tab(0).Control(33)=   "cboVehicleTypeJPJ"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cboLocationJPJ"
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "LPKP ( JPJ )"
      TabPicture(1)   =   "frmEditMotor.frx":14CF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label21"
      Tab(1).Control(1)=   "Line2"
      Tab(1).Control(2)=   "Line1(2)"
      Tab(1).Control(3)=   "Line1(1)"
      Tab(1).Control(4)=   "Label20"
      Tab(1).Control(5)=   "Label19"
      Tab(1).Control(6)=   "Label7"
      Tab(1).Control(7)=   "Label39"
      Tab(1).Control(8)=   "txtLicenseExpire"
      Tab(1).Control(9)=   "txtLicRefLPKP"
      Tab(1).Control(10)=   "txtFileRefLPKP"
      Tab(1).Control(11)=   "txtformLPKP"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Insurance (Jerneh)"
      TabPicture(2)   =   "frmEditMotor.frx":14EB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtPayable"
      Tab(2).Control(1)=   "txtPolicyNoJerneh"
      Tab(2).Control(2)=   "txtPeriodJerneh"
      Tab(2).Control(3)=   "txtVehicleSumJerneh"
      Tab(2).Control(4)=   "txtExcessJerneh"
      Tab(2).Control(5)=   "txtBasicJerneh"
      Tab(2).Control(6)=   "txtNCTJerneh"
      Tab(2).Control(7)=   "Label8"
      Tab(2).Control(8)=   "Line3"
      Tab(2).Control(9)=   "Label22"
      Tab(2).Control(10)=   "Label23"
      Tab(2).Control(11)=   "Label24"
      Tab(2).Control(12)=   "Label25"
      Tab(2).Control(13)=   "Label26"
      Tab(2).Control(14)=   "Label27"
      Tab(2).ControlCount=   15
      TabCaption(3)   =   "Puspacom"
      TabPicture(3)   =   "frmEditMotor.frx":1507
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label29"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label28"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Line6"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label36"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label37"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtTimePuspacom"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtDatePuspacom"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      Begin VB.ComboBox cboLocationJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66240
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4080
         Width           =   2655
      End
      Begin VB.ComboBox cboVehicleTypeJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71760
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   4560
         Width           =   2655
      End
      Begin VB.ComboBox cboFuelTypeJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71760
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   4080
         Width           =   2655
      End
      Begin VB.TextBox txtPayable 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   65
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox txtEngineJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71760
         TabIndex        =   2
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtCasiJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71760
         TabIndex        =   3
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtManufacturerJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71760
         TabIndex        =   4
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtModelJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71760
         TabIndex        =   5
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txtEnginePowerJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71760
         TabIndex        =   6
         Top             =   3600
         Width           =   2655
      End
      Begin VB.TextBox txtYearManuJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71760
         TabIndex        =   9
         Top             =   5040
         Width           =   2655
      End
      Begin VB.TextBox txtYearRegisJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66240
         TabIndex        =   10
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtLicenseDuration 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66240
         TabIndex        =   11
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtBDMBGKJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66240
         TabIndex        =   12
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtBTMJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66240
         TabIndex        =   13
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txtDriveJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66240
         TabIndex        =   14
         Top             =   3600
         Width           =   2655
      End
      Begin VB.TextBox txtSeriesJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66240
         TabIndex        =   16
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox txtStatusJPJ 
         DataField       =   "Status"
         DataSource      =   "Qc_Adodc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -66240
         TabIndex        =   17
         Top             =   5040
         Width           =   2655
      End
      Begin VB.TextBox txtformLPKP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   18
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtFileRefLPKP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   19
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtLicRefLPKP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   20
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtLicenseExpire 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   21
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txtPolicyNoJerneh 
         DataSource      =   "Qc_Adodc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   22
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtPeriodJerneh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   23
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtVehicleSumJerneh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   24
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtExcessJerneh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   25
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txtBasicJerneh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   26
         Top             =   3600
         Width           =   2655
      End
      Begin VB.TextBox txtNCTJerneh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   30
         Top             =   4080
         Width           =   2655
      End
      Begin VB.TextBox txtDatePuspacom 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtTimePuspacom 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   28
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label39 
         Caption         =   "( e.g 14 sep 2007 )"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69240
         TabIndex        =   73
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label38 
         Caption         =   "( e.g. 6 or12 Months)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -63480
         TabIndex        =   72
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label37 
         Caption         =   "( e.g 14-sep-2007 )"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   71
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label36 
         Caption         =   "( e.g 08:30 )"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   70
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label30 
         Caption         =   "( e.g. 1999 )"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -63480
         TabIndex        =   67
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Total Payable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   66
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Series No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68520
         TabIndex        =   58
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Engine No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   57
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Casi No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   56
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Manufacturer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   55
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Model Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   54
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Fuel Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   53
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Year Of Manufacturer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   52
         Top             =   5040
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "Year of Registration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68520
         TabIndex        =   51
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "B.D.M / B.G.K"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68520
         TabIndex        =   50
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "B.T.M"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68520
         TabIndex        =   49
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "License Duration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68520
         TabIndex        =   48
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "Driver Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68520
         TabIndex        =   47
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68520
         TabIndex        =   46
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Vehicle Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   45
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68520
         TabIndex        =   44
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Engine Horsepower"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   43
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   -75000
         X2              =   -59760
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label7 
         Caption         =   "File Reference No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   42
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "License Expiry Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   41
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   "Form No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   40
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   -75000
         X2              =   -62640
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   -75000
         X2              =   -62640
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         X1              =   -75000
         X2              =   -59760
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label21 
         Caption         =   "License Reference No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   39
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Line Line3 
         X1              =   -75000
         X2              =   -59760
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label22 
         Caption         =   "Policy No Insurance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   38
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label23 
         Caption         =   "Policy Period"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   37
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label24 
         Caption         =   "Vehicle Sum"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   36
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label25 
         Caption         =   "Excess Applicable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   35
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label26 
         Caption         =   "Basic premium"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   34
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label27 
         Caption         =   "NCT (DTT)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74040
         TabIndex        =   33
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   15240
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label28 
         Caption         =   "Inspection Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   32
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label29 
         Caption         =   "Inspection Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   31
         Top             =   2280
         Width           =   2295
      End
   End
   Begin VB.Label Label34 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quick Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   69
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Edit Motor Insurance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   68
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmEditMotor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cnn As ADODB.Connection
Dim qc As ADODB.Recordset
Dim FuelType As ADODB.Recordset
Dim VehicleType As ADODB.Recordset
Dim location As ADODB.Recordset



Private Sub cboregistration_Click()
If cboRegistration.ListIndex = -1 Then
    cmdEdit.Enabled = False
    Exit Sub
End If
cmdEdit.Enabled = True
 With qc
  .Requery
  .MoveLast
  .MoveFirst
    While Not .EOF = True
        If cboRegistration.List(cboRegistration.ListIndex) = .Fields(0) Then
            txtEngineJPJ.Text = .Fields(1)
            txtCasiJPJ.Text = .Fields(2)
            txtManufacturerJPJ.Text = .Fields(3)
            txtModelJPJ.Text = .Fields(4)
            txtEnginePowerJPJ.Text = .Fields(5)
            cboFuelTypeJPJ.Text = .Fields(6)
            cboVehicleTypeJPJ.Text = .Fields(7)
            txtYearManuJPJ.Text = .Fields(8)
            txtYearRegisJPJ.Text = .Fields(9)
            txtLicenseDuration.Text = .Fields(10)
            txtBDMBGKJPJ.Text = .Fields(11)
            txtBTMJPJ.Text = .Fields(12)
            txtDriveJPJ.Text = .Fields(13)
            cboLocationJPJ.Text = .Fields(14)
            txtSeriesJPJ.Text = .Fields(15)
            txtStatusJPJ.Text = .Fields(16)
            txtformLPKP.Text = .Fields(17)
            txtFileRefLPKP.Text = .Fields(18)
            txtLicRefLPKP.Text = .Fields(19)
            txtLicenseExpire.Text = .Fields(20)
            txtPolicyNoJerneh.Text = .Fields(21)
            txtPeriodJerneh.Text = .Fields(22)
            txtVehicleSumJerneh.Text = .Fields(23)
            txtExcessJerneh.Text = .Fields(24)
            txtBasicJerneh.Text = .Fields(25)
            txtNCTJerneh.Text = .Fields(26)
            txtPayable.Text = .Fields(27)
            txtDatePuspacom.Text = .Fields(28)
            txtTimePuspacom.Text = .Fields(29)
        End If
   .MoveNext
  Wend
 End With
End Sub



Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdEdit_Click()



'*******ERROR Prevention*****
cmdEdit.Enabled = False
cmdUpdate.Enabled = True
cmdCancel.Enabled = True
unlocked_editmotor
txtEngineJPJ.SetFocus

End Sub

Private Sub cmdUpdate_Click()





        Dim StrSql As String
          With qc
            .MoveLast
            .MoveFirst
              While .EOF = False
                    If cboRegistration.List(cboRegistration.ListIndex) = .Fields(0) Then
                    StrSql = "UPDATE qc_motor SET EngineNo = '" & txtEngineJPJ.Text & "'," _
                    & "CasisNo = '" & txtCasiJPJ.Text & "'," & "Manufacturer = '" & txtManufacturerJPJ.Text & "'," _
                    & "ModelName = '" & txtModelJPJ.Text & "'," & "EngineHorsepower = '" & txtEnginePowerJPJ.Text & "'," _
                    & "FuelType= '" & cboFuelTypeJPJ & "'," & "VehicleType = '" & cboVehicleTypeJPJ & "'," _
                    & "YearManufacturer= '" & txtYearManuJPJ.Text & "'," & "YearRegistration = '" & txtYearRegisJPJ.Text & "'," _
                    & "LicenseDuration= '" & txtLicenseDuration.Text & "'," & "BDMBGK = '" & txtBDMBGKJPJ.Text & "'," _
                    & "BTM= '" & txtBTMJPJ.Text & "'," & "DriverName = '" & txtDriveJPJ.Text & "'," _
                    & "Location= '" & cboLocationJPJ & "'," & "SeriesNo = '" & txtSeriesJPJ.Text & "'," _
                    & "Status= '" & txtStatusJPJ.Text & "'," & "LPKPFormNo = '" & txtformLPKP.Text & "'," _
                    & "FileReferenceNo= '" & txtFileRefLPKP.Text & "'," & "LicenseReferenceNo = '" & txtLicRefLPKP.Text & "'," _
                    & "LicenseExpire= '" & txtLicenseExpire.Text & "'," & "PolicyNoInsurance = '" & txtPolicyNoJerneh.Text & "'," _
                    & "PolicyPeriod= '" & txtPeriodJerneh.Text & "'," & "VehicleSumInsured = '" & txtVehicleSumJerneh.Text & "'," _
                    & "ExcessApplicable= '" & txtExcessJerneh.Text & "'," & "BasicPremium = '" & txtBasicJerneh.Text & "'," _
                    & "NCTDTT= '" & txtNCTJerneh.Text & "'," & "TotalPayable= '" & txtPayable.Text & "'," & "InspectioDate= '" & txtDatePuspacom.Text & "'," _
                    & "InspectionTime = '" & txtTimePuspacom.Text & "' " _
                    & "WHERE RegistrationNo = '" & cboRegistration.List(cboRegistration.ListIndex) & "';"
                
                    Cnn.Execute StrSql
                    .Update
                   MsgBox "Record Has Been Updated Successfully", vbOKOnly + vbInformation, "Correct"
                    .MoveNext
                        Else
                        .MoveNext
                    End If
            Wend
    End With
    '***********Error Prevetion*****'
    clear_editmotor
    locked_editmotor
    cmdEdit.Enabled = False
    cmdUpdate.Enabled = False
    cmdCancel.Enabled = True
    cboRegistration.SetFocus
    cboRegistration.ListIndex = -1
    qc.Requery

End Sub

Private Sub Form_Load()
Set Cnn = New ADODB.Connection
Set qc = New ADODB.Recordset
Set FuelType = New ADODB.Recordset
Set VehicleType = New ADODB.Recordset
Set location = New ADODB.Recordset

'open connection and data source
Cnn.CursorLocation = adUseClient
Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
         & App.Path & "\data.mdb;Persist Security Info=False"

qc.Open "select * from qc_motor", Cnn, adOpenDynamic, adLockPessimistic
FuelType.Open "select * from Fuel", Cnn, adOpenDynamic, adLockPessimistic
VehicleType.Open "select * from Vehicle_Body", Cnn, adOpenDynamic, adLockPessimistic
location.Open "select * from Location", Cnn, adOpenDynamic, adLockPessimistic

'*************Populate list from database********************'
With qc
    While Not .EOF
    cboRegistration.AddItem .Fields(0)
    .MoveNext
Wend
End With

With FuelType
    While Not .EOF
    cboFuelTypeJPJ.AddItem .Fields(0)
    .MoveNext
Wend
End With

With VehicleType
    While Not .EOF
    cboVehicleTypeJPJ.AddItem .Fields(0)
    .MoveNext
Wend
End With

With location
    While Not .EOF
    cboLocationJPJ.AddItem .Fields(0)
    .MoveNext
Wend
End With

'****************************load first tab****************'
SSTabMotorInsurance.Tab = 0


'****Error prevention***'
locked_editmotor
cmdEdit.Enabled = False
cmdUpdate.Enabled = False
cmdCancel.Enabled = True

End Sub


Public Sub locked_editmotor()

txtEngineJPJ.Locked = True
txtCasiJPJ.Locked = True
txtManufacturerJPJ.Locked = True
txtModelJPJ.Locked = True
txtEnginePowerJPJ.Locked = True
cboFuelTypeJPJ.Locked = True
cboVehicleTypeJPJ.Locked = True
txtYearManuJPJ.Locked = True
txtYearRegisJPJ.Locked = True
txtLicenseDuration.Locked = True
txtBDMBGKJPJ.Locked = True
txtBTMJPJ.Locked = True
txtDriveJPJ.Locked = True
cboLocationJPJ.Locked = True
txtSeriesJPJ.Locked = True
txtStatusJPJ.Locked = True
txtformLPKP.Locked = True
txtFileRefLPKP.Locked = True
txtLicRefLPKP.Locked = True
txtLicenseExpire.Locked = True
txtPolicyNoJerneh.Locked = True
txtPeriodJerneh.Locked = True
txtVehicleSumJerneh.Locked = True
txtExcessJerneh.Locked = True
txtBasicJerneh.Locked = True
txtNCTJerneh.Locked = True
txtDatePuspacom.Locked = True
txtTimePuspacom.Locked = True
End Sub
Public Sub unlocked_editmotor()

txtEngineJPJ.Locked = False
txtCasiJPJ.Locked = False
txtManufacturerJPJ.Locked = False
txtModelJPJ.Locked = False
txtEnginePowerJPJ.Locked = False
cboFuelTypeJPJ.Locked = False
cboVehicleTypeJPJ.Locked = False
txtYearManuJPJ.Locked = False
txtYearRegisJPJ.Locked = False
txtLicenseDuration.Locked = False
txtBDMBGKJPJ.Locked = False
txtBTMJPJ.Locked = False
txtDriveJPJ.Locked = False
cboLocationJPJ.Locked = False
txtSeriesJPJ.Locked = False
txtStatusJPJ.Locked = False
txtformLPKP.Locked = False
txtFileRefLPKP.Locked = False
txtLicRefLPKP.Locked = False
txtLicenseExpire.Locked = False
txtPolicyNoJerneh.Locked = False
txtPeriodJerneh.Locked = False
txtVehicleSumJerneh.Locked = False
txtExcessJerneh.Locked = False
txtBasicJerneh.Locked = False
txtNCTJerneh.Locked = False
txtPayable.Locked = False
txtDatePuspacom.Locked = False
txtTimePuspacom.Locked = False
End Sub
Public Sub clear_editmotor()
cboRegistration.ListIndex = -1
txtEngineJPJ.Text = ""
txtCasiJPJ.Text = ""
txtManufacturerJPJ.Text = ""
txtModelJPJ.Text = ""
txtEnginePowerJPJ.Text = ""
cboFuelTypeJPJ.ListIndex = -1
cboVehicleTypeJPJ.ListIndex = -1
txtYearManuJPJ.Text = ""
txtYearRegisJPJ.Text = ""
txtLicenseDuration.Text = ""
txtBDMBGKJPJ.Text = ""
txtBTMJPJ.Text = ""
txtDriveJPJ.Text = ""
cboLocationJPJ.ListIndex = -1
txtSeriesJPJ.Text = ""
txtStatusJPJ.Text = ""
txtformLPKP.Text = ""
txtFileRefLPKP.Text = ""
txtLicRefLPKP.Text = ""
txtLicenseExpire.Text = ""
txtPolicyNoJerneh.Text = ""
txtPeriodJerneh.Text = ""
txtVehicleSumJerneh.Text = ""
txtExcessJerneh.Text = ""
txtBasicJerneh.Text = ""
txtNCTJerneh.Text = ""
txtPayable.Text = ""
txtDatePuspacom.Text = ""
txtTimePuspacom.Text = ""
End Sub


Private Sub txtDatePuspacom_LostFocus()
txtDatePuspacom.Text = Format(txtDatePuspacom.Text, "dd-mmm-yyyy")
End Sub
Private Sub txtLicenseDuration_Validate(Cancel As Boolean)
If Not IsNumeric(txtLicenseDuration.Text) Then
MsgBox "Please Enter Numerical Month Value", vbOKOnly + vbInformation, "Invalid"
txtLicenseDuration.Text = ""
txtLicenseDuration.SetFocus
End If
End Sub
Private Sub txtLicenseExpire_LostFocus()
txtLicenseExpire.Text = Format(txtLicenseExpire.Text, "dd-mmm-yyyy")
End Sub

Private Sub txtTimePuspacom_LostFocus()
txtTimePuspacom.Text = Format(txtTimePuspacom.Text, "hh:mm")
End Sub

Private Sub txtYearManuJPJ_Validate(Cancel As Boolean)
If Not IsNumeric(txtYearManuJPJ.Text) Then
MsgBox "Please Enter A Numerical Month Value", vbOKOnly + vbInformation, "Invalid"
txtYearManuJPJ.Text = ""
txtYearManuJPJ.SetFocus
End If
End Sub

Private Sub txtYearRegisJPJ_Validate(Cancel As Boolean)
If Len(txtYearRegisJPJ.Text) > 4 Or Not IsNumeric(txtYearRegisJPJ.Text) Then
MsgBox "Please Enter a Numerical 4 Digit Year", vbOKOnly + vbInformation, "Invalid"
txtYearRegisJPJ.Text = ""
txtYearRegisJPJ.SetFocus
End If
End Sub
