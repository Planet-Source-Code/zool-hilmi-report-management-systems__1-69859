VERSION 5.00
Begin VB.Form test 
   BackColor       =   &H000000C0&
   Caption         =   "Search"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   11505
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   65
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   64
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   63
      Top             =   720
      Width           =   2415
   End
   Begin VB.ListBox lstEmpName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   240
      TabIndex        =   62
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Caption         =   "Puspacom"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1575
      Index           =   3
      Left            =   8880
      TabIndex        =   30
      Top             =   6720
      Width           =   4815
      Begin VB.TextBox txtTimePuspacom 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   61
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtDatePuspacom 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   60
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Inspection Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inspection Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Caption         =   "Insurance (Jerneh)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2775
      Index           =   2
      Left            =   8880
      TabIndex        =   23
      Top             =   3840
      Width           =   4815
      Begin VB.TextBox txtNCTJerneh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   59
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox txtBasicJerneh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   58
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txtExcessJerneh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   57
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtVehicleSumJerneh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   56
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtPeriodJerneh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   55
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtPolicyNoJerneh 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   54
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NCT (DTT)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Basic premium"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Excess Applicable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Sum"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Policy Period"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Policy No Insurance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Caption         =   "LPKP (JPJ)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2175
      Index           =   1
      Left            =   8880
      TabIndex        =   18
      Top             =   1560
      Width           =   4815
      Begin VB.TextBox txtLicDurLPKP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   53
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtLicRefLPKP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   52
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtFileRefLPKP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   51
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtformLPKP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   50
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Form No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "File Reference No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "License Reference No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "License Duration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Caption         =   "GreenCard(JPJ)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   6735
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   1560
      Width           =   5535
      Begin VB.TextBox txtStatusJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   49
         Top             =   6240
         Width           =   2175
      End
      Begin VB.TextBox txtSeriesJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   48
         Top             =   5880
         Width           =   2175
      End
      Begin VB.TextBox txtLocationJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   47
         Top             =   5520
         Width           =   2175
      End
      Begin VB.TextBox txtDriveJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   46
         Top             =   5160
         Width           =   2175
      End
      Begin VB.TextBox txtBTMJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   45
         Top             =   4800
         Width           =   2175
      End
      Begin VB.TextBox txtBDMBGKJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   44
         Top             =   4440
         Width           =   2175
      End
      Begin VB.TextBox txtLicenseJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   43
         Top             =   4080
         Width           =   2175
      End
      Begin VB.TextBox txtYearRegisJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   42
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox txtYearManuJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   41
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txtVehicleTypeJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   40
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtFuelTypeJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   39
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox txtEnginePowerJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   38
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtModelJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   37
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtManufacturerJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   36
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtCasiJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   35
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtEngineJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   34
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtRegistrationJPJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   33
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Series No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Driver Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "B.T.M"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "B.D.M / B.G.K"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "License Duration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Year Registered"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Year Manufacturer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle/Body Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine power"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Model Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacturer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Casis No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label Label18 
      BackColor       =   &H000000C0&
      Caption         =   "Search By Manufacturer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   67
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label lblcount 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   8400
      TabIndex        =   66
      Top             =   720
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0FF&
      X1              =   120
      X2              =   14280
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnn As ADODB.Connection
Dim Search As ADODB.Recordset
Dim StrSql As String

Public Sub FindX()

    ' connection and recordset variables
    Dim Cnxn As New ADODB.Connection
    Dim rstLocation As New ADODB.Recordset
    Dim strCnxn As String
    Dim strSQLManuName As String
    
     ' record variables
    Dim mark As Variant
    Dim count As Integer
    
     ' open connection
     lblcount.Caption = ""
    Set Cnxn = New ADODB.Connection
    strCnxn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
              & App.Path & "\data.mdb;Persist Security Info=False"
              
    Cnxn.Open strCnxn
       
    ' open recordset with default parameters which are
    ' sufficient to search forward through a Recordset
    Set rstLocation = New ADODB.Recordset
    strSQLManuName = "SELECT Location FROM QC where location='" & txtname.Text & "%'"
    rstLocation.Open strSQLManuName, Cnxn, adOpenStatic, adLockReadOnly, adCmdText

    count = 0
    Do While Not rstLocation.EOF
        'continue if last find succeeded
       lstEmpName.AddItem rstLocation!location
        'count the last title found
       count = count + 1
        ' note current position
       mark = rstLocation.Bookmark
    
        ' above code skips current record to avoid finding the same row repeatedly;
        ' last arg (bookmark) is redundant because Find searches from current position
    Loop
    If count = 0 Then
     MsgBox "No Match Found", vbOKOnly + vbInformation, "Information"
     txtname.SetFocus
    Else
     lblcount.Caption = "Total Matches found " & count
    End If
     ' clean up
    rstLocation.Close
    Cnxn.Close
    Set rstLocation = Nothing
    Set Cnxn = Nothing

End Sub

Private Sub cmdGo_Click()
lstEmpName.clear

If txtname.Text = "" Then
 MsgBox "Enter the name", vbOKOnly + vbCritical, "Error"
 txtname.SetFocus
 Exit Sub
End If

FindX
End Sub

Private Sub Form_Load()
Set cnn = New ADODB.Connection

Set Search = New ADODB.Recordset

cnn.CursorLocation = adUseClient
cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " _
         & App.Path & "\data.mdb;Persist Security Info = False"
         
Search.Open "SELECT * FROM qc ", cnn, adOpenDynamic, adLockPessimistic


lstEmpName.clear

End Sub

Private Sub lstEmpName_Click()
With Search
 .MoveFirst
 While Not .EOF
  If lstEmpName.List(lstEmpName.ListIndex) = .Fields(14) Then
        txtRegistrationJPJ.Text = .Fields(0)
        txtEngineJPJ.Text = .Fields(1)
        txtCasiJPJ.Text = .Fields(2)
        txtManufacturerJPJ.Text = .Fields(3)
        txtModelJPJ.Text = .Fields(4)
        txtEnginePowerJPJ.Text = .Fields(5)
        txtFuelTypeJPJ.Text = .Fields(6)
        txtVehicleTypeJPJ.Text = .Fields(7)
        txtYearManuJPJ.Text = .Fields(8)
        txtYearRegisJPJ.Text = .Fields(9)
        txtLicenseJPJ.Text = .Fields(10)
        txtBDMBGKJPJ.Text = .Fields(11)
        txtBTMJPJ.Text = .Fields(12)
        txtDriveJPJ.Text = .Fields(13)
        txtLocationJPJ.Text = .Fields(14)
        txtSeriesJPJ.Text = .Fields(15)
        txtStatusJPJ.Text = .Fields(16)
        txtformLPKP.Text = .Fields(17)
        txtFileRefLPKP.Text = .Fields(18)
        txtLicRefLPKP.Text = .Fields(19)
        txtLicDurLPKP.Text = .Fields(20)
        txtPolicyNoJerneh.Text = .Fields(21)
        txtPeriodJerneh.Text = .Fields(22)
        txtVehicleSumJerneh.Text = .Fields(23)
        txtExcessJerneh.Text = .Fields(24)
        txtBasicJerneh.Text = .Fields(25)
        txtNCTJerneh.Text = .Fields(26)
        txtDatePuspacom.Text = .Fields(27)
        txtTimePuspacom.Text = .Fields(28)
  End If
  .MoveNext
 Wend
 End With
End Sub
