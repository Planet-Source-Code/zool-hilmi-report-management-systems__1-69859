VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMotorInsurance 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   1455
      Left            =   5040
      TabIndex        =   58
      Top             =   7920
      Width           =   3615
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   480
         MaskColor       =   &H8000000F&
         MouseIcon       =   "frmMotorInsurance.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmMotorInsurance.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Add new record"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   2400
         MouseIcon       =   "frmMotorInsurance.frx":0730
         MousePointer    =   99  'Custom
         Picture         =   "frmMotorInsurance.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Cancel"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   1440
         MouseIcon       =   "frmMotorInsurance.frx":0E02
         MousePointer    =   99  'Custom
         Picture         =   "frmMotorInsurance.frx":0F54
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Save record"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label32 
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
         TabIndex        =   61
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Save "
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
         TabIndex        =   60
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Add "
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
         Left            =   600
         TabIndex        =   59
         Top             =   1080
         Width           =   495
      End
   End
   Begin TabDlg.SSTab SSTabMotorInsurance 
      Height          =   6735
      Left            =   240
      TabIndex        =   21
      Top             =   840
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   4210816
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
      TabPicture(0)   =   "frmMotorInsurance.frx":14EC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(9)=   "Label12"
      Tab(0).Control(10)=   "Label13"
      Tab(0).Control(11)=   "Label14"
      Tab(0).Control(12)=   "Label15"
      Tab(0).Control(13)=   "Label16"
      Tab(0).Control(14)=   "Label17"
      Tab(0).Control(15)=   "Label8"
      Tab(0).Control(16)=   "Label18"
      Tab(0).Control(17)=   "Line1(0)"
      Tab(0).Control(18)=   "Label33"
      Tab(0).Control(19)=   "Label39"
      Tab(0).Control(20)=   "txtRegistrationJPJ"
      Tab(0).Control(21)=   "txtEngineJPJ"
      Tab(0).Control(22)=   "txtCasiJPJ"
      Tab(0).Control(23)=   "txtManufacturerJPJ"
      Tab(0).Control(24)=   "txtModelJPJ"
      Tab(0).Control(25)=   "txtEnginePowerJPJ"
      Tab(0).Control(26)=   "txtYearManuJPJ"
      Tab(0).Control(27)=   "txtYearRegisJPJ"
      Tab(0).Control(28)=   "txtLicenseDuration"
      Tab(0).Control(29)=   "txtBDMBGKJPJ"
      Tab(0).Control(30)=   "txtBTMJPJ"
      Tab(0).Control(31)=   "txtDriveJPJ"
      Tab(0).Control(32)=   "txtSeriesJPJ"
      Tab(0).Control(33)=   "txtStatusJPJ"
      Tab(0).Control(34)=   "CboFuelType"
      Tab(0).Control(35)=   "cboVehicleBody"
      Tab(0).Control(36)=   "cboLocation"
      Tab(0).ControlCount=   37
      TabCaption(1)   =   "LPKP ( JPJ )"
      TabPicture(1)   =   "frmMotorInsurance.frx":1508
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Label19"
      Tab(1).Control(2)=   "Label20"
      Tab(1).Control(3)=   "Line1(1)"
      Tab(1).Control(4)=   "Line2"
      Tab(1).Control(5)=   "Label21"
      Tab(1).Control(6)=   "Label38"
      Tab(1).Control(7)=   "txtformLPKP"
      Tab(1).Control(8)=   "txtFileRefLPKP"
      Tab(1).Control(9)=   "txtLicRefLPKP"
      Tab(1).Control(10)=   "txtLicenseExpire"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Insurance (Jerneh)"
      TabPicture(2)   =   "frmMotorInsurance.frx":1524
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Line3"
      Tab(2).Control(1)=   "Label22"
      Tab(2).Control(2)=   "Label23"
      Tab(2).Control(3)=   "Label24"
      Tab(2).Control(4)=   "Label25"
      Tab(2).Control(5)=   "Label26"
      Tab(2).Control(6)=   "Label27"
      Tab(2).Control(7)=   "Label34"
      Tab(2).Control(8)=   "txtPolicyNoJerneh"
      Tab(2).Control(9)=   "txtPeriodJerneh"
      Tab(2).Control(10)=   "txtVehicleSumJerneh"
      Tab(2).Control(11)=   "txtExcessJerneh"
      Tab(2).Control(12)=   "txtBasicJerneh"
      Tab(2).Control(13)=   "txtNCTJerneh"
      Tab(2).Control(14)=   "txtPayable"
      Tab(2).ControlCount=   15
      TabCaption(3)   =   "Puspacom"
      TabPicture(3)   =   "frmMotorInsurance.frx":1540
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Line6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label28"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label29"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label37"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label36"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtDatePuspacom"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtTimePuspacom"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
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
         TabIndex        =   68
         Top             =   4560
         Width           =   2655
      End
      Begin VB.ComboBox cboLocation 
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
         TabIndex        =   65
         Top             =   4080
         Width           =   2655
      End
      Begin VB.ComboBox cboVehicleBody 
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
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   5040
         Width           =   2655
      End
      Begin VB.ComboBox CboFuelType 
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
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   4560
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
         TabIndex        =   29
         Top             =   2160
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
         TabIndex        =   28
         Top             =   1680
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
         TabIndex        =   27
         Top             =   4080
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
      Begin VB.TextBox txtPolicyNoJerneh 
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
         TabIndex        =   19
         Top             =   3120
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
         TabIndex        =   18
         Top             =   2640
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
         TabIndex        =   17
         Top             =   2160
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
         TabIndex        =   16
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtStatusJPJ 
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
         TabIndex        =   15
         Top             =   5040
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
         TabIndex        =   14
         Top             =   4560
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
         TabIndex        =   13
         Top             =   3600
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
         TabIndex        =   12
         Top             =   3120
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
         TabIndex        =   11
         Top             =   2640
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
         TabIndex        =   10
         Top             =   2160
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
         TabIndex        =   9
         Top             =   1680
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
         Left            =   -71880
         TabIndex        =   8
         Top             =   5520
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
         Left            =   -71880
         TabIndex        =   5
         Top             =   4080
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
         Left            =   -71880
         TabIndex        =   4
         Top             =   3600
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
         Left            =   -71880
         TabIndex        =   3
         Top             =   3120
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
         Left            =   -71880
         TabIndex        =   2
         Top             =   2640
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
         Left            =   -71880
         TabIndex        =   1
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtRegistrationJPJ 
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
         Left            =   -71880
         TabIndex        =   0
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label38 
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
         Left            =   -69120
         TabIndex        =   73
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label39 
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
         Height          =   255
         Left            =   -63480
         TabIndex        =   72
         Top             =   2280
         Width           =   1815
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
         TabIndex        =   71
         Top             =   2280
         Width           =   1935
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
         TabIndex        =   70
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label34 
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
         TabIndex        =   67
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label33 
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
         Height          =   255
         Left            =   -63480
         TabIndex        =   66
         Top             =   1800
         Width           =   1215
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
         TabIndex        =   57
         Top             =   2280
         Width           =   2295
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
         TabIndex        =   56
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   15120
         Y1              =   1200
         Y2              =   1200
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
         TabIndex        =   55
         Top             =   4080
         Width           =   1455
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
         TabIndex        =   54
         Top             =   3600
         Width           =   1695
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
         TabIndex        =   53
         Top             =   3120
         Width           =   1575
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
         TabIndex        =   52
         Top             =   2640
         Width           =   1455
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
         TabIndex        =   51
         Top             =   2160
         Width           =   1695
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
         TabIndex        =   50
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Line Line3 
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
         TabIndex        =   49
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Line Line2 
         X1              =   -75000
         X2              =   -59880
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   -75000
         X2              =   -62640
         Y1              =   0
         Y2              =   0
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
         TabIndex        =   48
         Top             =   1680
         Width           =   1815
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
         TabIndex        =   47
         Top             =   3120
         Width           =   1815
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
         TabIndex        =   46
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   -75000
         X2              =   -60240
         Y1              =   1200
         Y2              =   1200
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
         TabIndex        =   45
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Registration No"
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
         TabIndex        =   44
         Top             =   1680
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
         TabIndex        =   43
         Top             =   5040
         Width           =   1335
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
         TabIndex        =   42
         Top             =   5040
         Width           =   1335
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
         TabIndex        =   41
         Top             =   4080
         Width           =   1095
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
         TabIndex        =   40
         Top             =   3600
         Width           =   1695
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
         TabIndex        =   39
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label12 
         Caption         =   "BTM"
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
         TabIndex        =   38
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "BDM / BGK"
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
         TabIndex        =   37
         Top             =   2640
         Width           =   1815
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
         TabIndex        =   36
         Top             =   1680
         Width           =   1695
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
         TabIndex        =   35
         Top             =   5520
         Width           =   2415
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
         TabIndex        =   34
         Top             =   4560
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
         TabIndex        =   33
         Top             =   3600
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
         TabIndex        =   32
         Top             =   3120
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
         TabIndex        =   31
         Top             =   2640
         Width           =   1215
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
         TabIndex        =   30
         Top             =   2160
         Width           =   1335
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
         TabIndex        =   20
         Top             =   4560
         Width           =   1455
      End
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add New Record Motor Insurance"
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
      Left            =   5400
      TabIndex        =   69
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmMotorInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cnn As ADODB.Connection
Dim FuelType As ADODB.Recordset
Dim qc As ADODB.Recordset
Dim VehicleType As ADODB.Recordset
Dim location As ADODB.Recordset

Private Sub cmdAdd_Click()
enable
txtRegistrationJPJ.SetFocus
cmdAdd.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True


End Sub

Private Sub cmdCancel_Click()
cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
Disable
Unload Me
End Sub



Private Sub cmdSave_Click()
   '***************adding value to database*****************'
   
If CboFuelType.ListIndex = -1 Or cboVehicleBody.ListIndex = -1 Or cboLocation.ListIndex = -1 Then
MsgBox "Please Ensure FuelType/VehicleBody/Location is Fill", vbOKOnly + vbInformation, "Invalid"
Exit Sub
End If

With qc
         .AddNew
         .Fields(0) = txtRegistrationJPJ.Text
         .Fields(1) = txtEngineJPJ.Text
         .Fields(2) = txtCasiJPJ.Text
         .Fields(3) = txtManufacturerJPJ.Text
         .Fields(4) = txtModelJPJ.Text
         .Fields(5) = txtEnginePowerJPJ.Text
         .Fields(6) = CboFuelType.List(CboFuelType.ListIndex)
         .Fields(7) = cboVehicleBody.List(cboVehicleBody.ListIndex)
         .Fields(8) = txtYearManuJPJ.Text
         .Fields(9) = txtYearRegisJPJ.Text
         .Fields(10) = txtLicenseDuration.Text
         .Fields(11) = txtBDMBGKJPJ.Text
         .Fields(12) = txtBTMJPJ.Text
         .Fields(13) = txtDriveJPJ.Text
         .Fields(14) = cboLocation.List(cboLocation.ListIndex)
         .Fields(15) = txtSeriesJPJ.Text
         .Fields(16) = txtStatusJPJ.Text
         .Fields(17) = txtformLPKP.Text
         .Fields(18) = txtFileRefLPKP.Text
         .Fields(19) = txtLicRefLPKP.Text
         .Fields(20) = txtLicenseExpire.Text
         .Fields(21) = txtPolicyNoJerneh.Text
         .Fields(22) = txtPeriodJerneh.Text
         .Fields(23) = txtVehicleSumJerneh.Text
         .Fields(24) = txtExcessJerneh.Text
         .Fields(25) = txtBasicJerneh.Text
         .Fields(26) = txtNCTJerneh.Text
         .Fields(27) = txtPayable.Text
         .Fields(28) = txtDatePuspacom.Text
         .Fields(29) = txtTimePuspacom.Text
        .Update
        MsgBox "Record Has Been Entered Sucessfully", vbOKOnly + vbInformation, "Correct"
 End With

clear_info
Disable

cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = True

    



End Sub

Private Sub Form_Load()

'**declare and set connection
Set Cnn = New ADODB.Connection
Set FuelType = New ADODB.Recordset
Set VehicleType = New ADODB.Recordset
Set location = New ADODB.Recordset
Set qc = New ADODB.Recordset

'open connection and data source
Cnn.CursorLocation = adUseClient
Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
         & App.Path & "\data.mdb;Persist Security Info=False"

FuelType.Open "SELECT * FROM Fuel", Cnn, adOpenDynamic, adLockPessimistic
qc.Open "select * from qc_motor", Cnn, adOpenDynamic, adLockPessimistic
VehicleType.Open "select * from vehicle_body", Cnn, adOpenDynamic, adLockPessimistic
location.Open "select * from location", Cnn, adOpenDynamic, adLockPessimistic


'**cal from file data then select fuel table and display into form**'
With FuelType
  While Not .EOF
   CboFuelType.AddItem .Fields(0)
   .MoveNext
  Wend
 End With
 
 With location
  While Not .EOF
   cboLocation.AddItem .Fields(0)
   .MoveNext
  Wend
 End With


'***********cal from file data then select vehicle_body table and display into form**********'
With VehicleType
    While Not .EOF
    cboVehicleBody.AddItem .Fields(0)
   .MoveNext
Wend
End With

'********************load first tab************************'
SSTabMotorInsurance.Tab = 0






 Disable
cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = True

End Sub

    
Public Sub enable()
txtRegistrationJPJ.Enabled = True
txtEngineJPJ.Enabled = True
txtCasiJPJ.Enabled = True
txtManufacturerJPJ.Enabled = True
txtModelJPJ.Enabled = True
txtEnginePowerJPJ.Enabled = True
CboFuelType.Enabled = True
cboVehicleBody.Enabled = True
txtYearManuJPJ.Enabled = True
txtYearRegisJPJ.Enabled = True
txtLicenseDuration.Enabled = True
txtBDMBGKJPJ.Enabled = True
txtBTMJPJ.Enabled = True
txtDriveJPJ.Enabled = True
cboLocation.Enabled = True
txtSeriesJPJ.Enabled = True
txtStatusJPJ.Enabled = True
txtformLPKP.Enabled = True
txtFileRefLPKP.Enabled = True
txtLicRefLPKP.Enabled = True
txtLicenseExpire.Enabled = True
txtPolicyNoJerneh.Enabled = True
txtPeriodJerneh.Enabled = True
txtVehicleSumJerneh.Enabled = True
txtExcessJerneh.Enabled = True
txtBasicJerneh.Enabled = True
txtNCTJerneh.Enabled = True
txtPayable.Enabled = True
txtDatePuspacom.Enabled = True
txtTimePuspacom.Enabled = True
End Sub

Public Sub Disable()
txtRegistrationJPJ.Enabled = False
txtEngineJPJ.Enabled = False
txtCasiJPJ.Enabled = False
txtManufacturerJPJ.Enabled = False
txtModelJPJ.Enabled = False
txtEnginePowerJPJ.Enabled = False
CboFuelType.Enabled = False
cboVehicleBody.Enabled = False
txtYearManuJPJ.Enabled = False
txtYearRegisJPJ.Enabled = False
txtLicenseDuration.Enabled = False
txtBDMBGKJPJ.Enabled = False
txtBTMJPJ.Enabled = False
txtDriveJPJ.Enabled = False
cboLocation.Enabled = False
txtSeriesJPJ.Enabled = False
txtStatusJPJ.Enabled = False
txtformLPKP.Enabled = False
txtFileRefLPKP.Enabled = False
txtLicRefLPKP.Enabled = False
txtLicenseExpire.Enabled = False
txtPolicyNoJerneh.Enabled = False
txtPeriodJerneh.Enabled = False
txtVehicleSumJerneh.Enabled = False
txtExcessJerneh.Enabled = False
txtBasicJerneh.Enabled = False
txtNCTJerneh.Enabled = False
txtPayable.Enabled = False
txtDatePuspacom.Enabled = False
txtTimePuspacom.Enabled = False
End Sub

Public Sub clear_info()
txtRegistrationJPJ.Text = ""
txtEngineJPJ.Text = ""
txtCasiJPJ.Text = ""
txtManufacturerJPJ.Text = ""
txtModelJPJ.Text = ""
txtEnginePowerJPJ.Text = ""
CboFuelType.Clear
cboVehicleBody.Clear
txtYearManuJPJ.Text = ""
txtYearRegisJPJ.Text = ""
txtLicenseDuration.Text = ""
txtBDMBGKJPJ.Text = ""
txtBTMJPJ.Text = ""
txtDriveJPJ.Text = ""
cboLocation.Clear
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
MsgBox "Pease Enter A Nuermical Month Value", vbOKOnly + vbInformation, "Invalid"
txtLicenseDuration.Text = ""
txtLicenseDuration.SetFocus
End If
End Sub





Private Sub txtRegistrationJPJ_Validate(Cancel As Boolean)
If txtRegistrationJPJ.Text = "" Then
MsgBox "Please Enter Registration No", vbOKOnly + vbInformation, "Invalid"
txtRegistrationJPJ.SetFocus
End If
End Sub

Private Sub txtTimePuspacom_LostFocus()
txtTimePuspacom.Text = Format(txtTimePuspacom.Text, "hh:mm")
End Sub
Private Sub txtYearManuJPJ_Validate(Cancel As Boolean)
If Len(txtYearManuJPJ.Text) > 4 Or Not IsNumeric(txtYearManuJPJ.Text) Then
MsgBox "Please Enter A 4 Digit Numerical Year Value", vbOKOnly + vbInformation, "Invalid"
txtYearManuJPJ.Text = ""
txtYearManuJPJ.SetFocus
End If

End Sub

Private Sub txtYearRegisJPJ_Validate(Cancel As Boolean)
If Len(txtYearRegisJPJ.Text) < 4 Or Not IsNumeric(txtYearRegisJPJ.Text) Then
MsgBox "Please Enter a Numerical 4 Digit Year", vbOKOnly + vbInformation, "Invalid"
txtYearRegisJPJ.Text = ""
txtYearRegisJPJ.SetFocus
End If
End Sub
