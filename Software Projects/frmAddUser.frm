VERSION 5.00
Begin VB.Form frmAddUser 
   BackColor       =   &H000000C0&
   Caption         =   "New User Option"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   6360
      MouseIcon       =   "frmAddUser.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmAddUser.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Save record"
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4920
      Picture         =   "frmAddUser.frx":06EA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtPassword2 
      Enabled         =   0   'False
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
      Left            =   6000
      TabIndex        =   3
      Top             =   3840
      Width           =   3255
   End
   Begin VB.TextBox txtUsername 
      Enabled         =   0   'False
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
      Left            =   6000
      TabIndex        =   1
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox txtPassword1 
      Enabled         =   0   'False
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
      Left            =   6000
      TabIndex        =   2
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   3960
      TabIndex        =   8
      Top             =   4440
      Width           =   5415
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   3840
         MouseIcon       =   "frmAddUser.frx":0B9D
         MousePointer    =   99  'Custom
         Picture         =   "frmAddUser.frx":0CEF
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
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
         Left            =   3960
         TabIndex        =   13
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H000000C0&
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
         Left            =   2640
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Add User"
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
         Left            =   1080
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   5400
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   2880
      X2              =   10200
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select the username using Alpahbateical Character Only.Password can be be a mix of Alpabateical and Number"
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
      Height          =   975
      Left            =   6000
      TabIndex        =   6
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   3960
      Picture         =   "frmAddUser.frx":1269
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Retype password"
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
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Index           =   0
      Left            =   3720
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Username"
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
      Height          =   495
      Index           =   0
      Left            =   3720
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnn As ADODB.Connection
Dim login As ADODB.Recordset


Private Sub cmdAdd_Click()
cmdSave.Enabled = True
txtUsername.Enabled = True
txtPassword1.Enabled = True
txtPassword2.Enabled = True
txtUsername.SetFocus
cmdAdd.Enabled = False
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
With adminpass


'check the character less than 3 or not
If Len(txtUsername.Text) <= 3 Then
   MsgBox "Please Select Username With More than 3 character"
    txtUsername.SetFocus




'check password if less than or equal to 3
Else
If Len(txtPassword1.Text) <= 3 Then
    MsgBox "Please Select Password with more than 3 Character"
     txtPassword1.SetFocus
     
     
     
'check password equal or not
Else
If txtPassword1.Text <> txtPassword2.Text Then
    MsgBox "The Password Does Not Match"
    txtPassword1.Text = ""
    txtPassword2.Text = ""
    txtPassword1.SetFocus
    
    
'check for emptied field
Else
If txtUsername.Text = "" Or txtPassword1.Text = "" Or txtPassword2.Text = "" Then
    MsgBox "Please Complete The information"
    txtUsername.SetFocus



'if username exist or not
Else
If txtUsername.Text = .Fields(0).Value Then
    MsgBox "Username You have selected already exist"
    txtUsername = ""
    txtPassword1.Text = ""
    txtPassword2.Text = ""
    txtUsername.SetFocus


'if everything is fine then add to database
Else
    .AddNew
    .Fields(0) = txtUsername.Text
    .Fields(1) = txtPassword1.Text
    .Update
    MsgBox "New User Has been Created Succesfully"
txtUsername.Text = ""
txtPassword1.Text = ""
txtPassword2.Text = ""
txtUsername.Enabled = False
txtPassword1.Enabled = False
txtPassword2.Enabled = False
cmdSave.Enabled = False
cmdAdd.Enabled = True
End If
End If
End If
End If
End If
End With


End Sub

Private Sub Form_Load()
Set cnn = New ADODB.Connection
Set login = New ADODB.Recordset

cnn.CursorLocation = adUseClient
cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
        & App.Path & "\data.mdb;Persist Security Info=False"

login.Open "SELECT * FROM password_login", cnn, adOpenDynamic, adLockPessimistic


End Sub


