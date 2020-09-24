VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "QUALITY CONCRETE VEHICLE INFORMATION SYSTEMS"
   ClientHeight    =   9180
   ClientLeft      =   165
   ClientTop       =   915
   ClientWidth     =   12825
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   327682
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Left            =   4680
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4800
      Top             =   3600
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8805
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Object.Tag             =   ""
            Object.ToolTipText     =   "User"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "***This Project is developed by Zool and Jeak Yii**"
            TextSave        =   "***This Project is developed by Zool and Jeak Yii**"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Author"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Capital Letter Status"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Numerical Pad Status"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "9/29/2007"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Date Today"
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   1
            TextSave        =   "9:42 PM"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Time"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuInsurance 
      Caption         =   "&Insurance"
      WindowList      =   -1  'True
      Begin VB.Menu mnuMotorInsurance 
         Caption         =   "Motor Insurance"
         Begin VB.Menu mnuAddNewRecordMotor 
            Caption         =   "&Add New Record (Motor)"
         End
         Begin VB.Menu mnuEditRecordMotor 
            Caption         =   "&Edit Record (Motor)"
         End
         Begin VB.Menu mnuSearchRecordMotor 
            Caption         =   "&Search Record (Motor)"
         End
      End
      Begin VB.Menu mnuNonMotorInsurance 
         Caption         =   "Non Motor Insurance"
         Begin VB.Menu mnuAddRecordNonMotor 
            Caption         =   "Add New Record (Non Motor)"
         End
         Begin VB.Menu mnuEditRecordNonMotor 
            Caption         =   "Edit Record (Non Motor)"
         End
         Begin VB.Menu mnuSearchRecordNonMotor 
            Caption         =   "Search Record (Non Motor)"
         End
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuMotorReport 
         Caption         =   "Motor"
         Begin VB.Menu mnuMotorDetailed 
            Caption         =   "DetailedReport"
         End
         Begin VB.Menu mnuMotorFilterReport 
            Caption         =   "FilterReport"
         End
      End
      Begin VB.Menu mnuNonMotorReport 
         Caption         =   "NonMotor"
         Begin VB.Menu mnuNonMotorDetailedReport 
            Caption         =   "DetailedReport"
         End
         Begin VB.Menu mnuNonMotorFilterReport 
            Caption         =   "FilterReport"
         End
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Option"
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuDeletePassword 
         Caption         =   "Delete User"
      End
   End
   Begin VB.Menu mnuLogOut 
      Caption         =   "&Log Out"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
      (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
      ByVal lpParameters As String, ByVal lpDirectory As String, _
      ByVal nShowCmd As Long) As Long
      Const SW_SHOWNORMAL = 1




Private Sub MDIForm_Load()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Programmer: Zool Hilmi Osman  kenet18m@yahoo.com          ''
''           : Yii Soon Jeak                                 ''
''                                                           ''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'-------------------------------------------------------------------------
'This project is Developed as Part of Our Final Semester Project For HIT3034
'Bachelor of Business Information Systems
'Client:Quality Concrete Sdn Bhd
'------------------------------------------------------------------------

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuAddNewRecordMotor_Click()
Forms (1)
End Sub

Private Sub mnuAddRecordNonMotor_Click()
Forms (4)
End Sub

Private Sub mnuChangePassword_Click()
Forms (11)
End Sub

Private Sub mnuDeletePassword_Click()
Forms (12)
End Sub

Private Sub mnuEditRecordMotor_Click()
Forms (2)
End Sub

Private Sub mnuEditRecordNonMotor_Click()
Forms (5)
End Sub

Private Sub mnuHelp_Click()
Call ShellExecute(Me.HWnd, "open", "yea.chm", "", 0, SW_SHOWNORMAL)
End Sub

Private Sub mnuLogOut_Click()
If MsgBox("Are You Sure ?", vbYesNo + vbInformation, "Warning") = vbYes Then
    End
End If
End Sub

Private Sub mnuMotorDetailed_Click()



Set cnnreport = New ADODB.Connection
Set reportmotor = New ADODB.Recordset
Set reportnonmotor = New ADODB.Recordset

'open connection and data source
cnnreport.CursorLocation = adUseClient
cnnreport.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
         & App.Path & "\data.mdb;Persist Security Info=False"

reportmotor.Open "select * from qc_motor", cnnreport, adOpenDynamic, adLockPessimistic
Set DetailedReportMotor.DataSource = reportmotor

reportnonmotor.Open "select * from qc_nonmotor", cnnreport, adOpenDynamic, adLockPessimistic
Set DetailedReportNonmotor.DataSource = reportnonmotor
Forms (7)
End Sub

Private Sub mnuMotorFilterReport_Click()
Forms (8)
End Sub

Private Sub mnuNonMotorDetailedReport_Click()
Set cnnreport = New ADODB.Connection
Set reportmotor = New ADODB.Recordset
Set reportnonmotor = New ADODB.Recordset

'open connection and data source
cnnreport.CursorLocation = adUseClient
cnnreport.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
         & App.Path & "\data.mdb;Persist Security Info=False"

reportmotor.Open "select * from qc_motor", cnnreport, adOpenDynamic, adLockPessimistic
Set DetailedReportMotor.DataSource = reportmotor

reportnonmotor.Open "select * from qc_nonmotor", cnnreport, adOpenDynamic, adLockPessimistic
Set DetailedReportNonmotor.DataSource = reportnonmotor

Forms (9)
End Sub

Private Sub mnuNonMotorFilterReport_Click()
Forms (10)
End Sub

Private Sub mnuSearchRecordMotor_Click()
Forms (3)
End Sub

Private Sub mnuSearchRecordNonMotor_Click()
Forms (6)
End Sub

Private Sub Timer1_Timer()
StatusBar.Panels(2).Text = Right(StatusBar.Panels(2).Text, Len(StatusBar.Panels(2).Text) - 1) & Left(StatusBar.Panels(2).Text, 1)
End Sub








