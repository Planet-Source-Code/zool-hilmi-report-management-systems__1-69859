VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Log in"
   ClientHeight    =   3300
   ClientLeft      =   3345
   ClientTop       =   3375
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1949.748
   ScaleMode       =   0  'User
   ScaleWidth      =   5450.582
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboUsername 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "frmLogin.frx":0000
      Left            =   1680
      List            =   "frmLogin.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      MouseIcon       =   "frmLogin.frx":0004
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":0156
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancel to Abort"
      Top             =   2400
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      MouseIcon       =   "frmLogin.frx":06D6
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":0828
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Click to submit"
      Top             =   2400
      Width           =   945
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   3045
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   225.347
      X2              =   5070.309
      Y1              =   638.099
      Y2              =   638.099
   End
   Begin VB.Label Label3 
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
      Left            =   3720
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Log In"
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
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   360
      Picture         =   "frmLogin.frx":0D9A
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter Your Username and Password To Log In."
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
      Height          =   735
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&User Name:"
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
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1470
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Password:"
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
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1860
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cnn As ADODB.Connection
Dim login As ADODB.Recordset
Dim UserName, Pas As String



Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

Private Sub cmdOK_Click()

If cboUsername.ListIndex = -1 Then
 MsgBox "Select User Name", vbOKOnly + vbInformation, "Error"
End If
UserName = cboUsername.List(cboUsername.ListIndex)

With login
  .MoveLast
  .MoveFirst
  While Not .EOF
    If UserName = .Fields(0).Value And _
      txtPassword.Text = .Fields(1).Value Then
      MsgBox "Access Granted", vbOKOnly + vbInformation, "Correct"
      If UserName = "Admin" Then
        MDIForm1.StatusBar.Panels(1) = "User Name :- " & UserName
        
      ElseIf UserName = "Staff" Then
        MDIForm1.StatusBar.Panels(1) = "User Name :- " & UserName
        MDIForm1.mnuDeletePassword.Enabled = False
    
       End If
      MDIForm1.Show
      Unload Me
      Exit Sub
    ElseIf cboUsername.List(cboUsername.ListIndex) = .Fields(0).Value And _
      Not txtPassword.Text = .Fields(1).Value Then
      MsgBox "Invalid Password", vbOKOnly + vbInformation, "Error"
      txtPassword.Text = ""
      txtPassword.SetFocus
      Exit Sub
    Else
     .MoveNext
    End If
  Wend
End With
End Sub

Private Sub Form_Load()


Set Cnn = New ADODB.Connection
Set login = New ADODB.Recordset

Cnn.CursorLocation = adUseClient
Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
        & App.Path & "\data.mdb;Persist Security Info=False"

login.Open "SELECT * FROM password_table", Cnn, adOpenDynamic, adLockPessimistic

With login
 .MoveLast
 .MoveFirst
  While Not .EOF
   cboUsername.AddItem .Fields(0).Value
   .MoveNext
  Wend
End With
End Sub

