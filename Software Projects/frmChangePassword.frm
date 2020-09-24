VERSION 5.00
Begin VB.Form frmChangePassword 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Delete Username"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtConfirmpassword 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3840
      TabIndex        =   5
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox txtNewPassword 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3840
      TabIndex        =   4
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtOldPassword 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3840
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   1455
      Left            =   1560
      TabIndex        =   8
      Top             =   3600
      Width           =   5055
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   3480
         MouseIcon       =   "frmChangePassword.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmChangePassword.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   960
         MouseIcon       =   "frmChangePassword.frx":06CC
         MousePointer    =   99  'Custom
         Picture         =   "frmChangePassword.frx":081E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Edit record"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   2280
         MouseIcon       =   "frmChangePassword.frx":0DC3
         MousePointer    =   99  'Custom
         Picture         =   "frmChangePassword.frx":0F15
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Save record"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Left            =   3600
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
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
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Edit "
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
         Left            =   1200
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Change Password Username"
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
      Left            =   1560
      TabIndex        =   15
      Top             =   120
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   1560
      X2              =   6600
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   1440
      Picture         =   "frmChangePassword.frx":14AD
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblUsername 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label lblconfirmpassword 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
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
      Height          =   360
      Left            =   1680
      TabIndex        =   2
      Top             =   3000
      Width           =   1560
   End
   Begin VB.Label lblnewpassword 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
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
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label lbloldpassword 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
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
      Height          =   360
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   1170
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cnn As ADODB.Connection
Dim login As ADODB.Recordset
Dim i As Integer
Dim str, UserName, StrSql As String



Private Sub cmdCancel_Click()
Unload Me



End Sub

Private Sub cmdEdit_Click()
cmdSave.Enabled = True
txtOldPassword.Enabled = True
txtNewPassword.Enabled = True
txtConfirmpassword.Enabled = True
txtOldPassword.SetFocus
cmdEdit.Enabled = False
End Sub

Private Sub cmdSave_Click()
'if oldpassword is left blank
 If txtOldPassword.Text = "" Then
    MsgBox "Enter Complete Information", vbOKOnly + vbInformation, "Information"
    txtOldPassword.SetFocus
    Exit Sub
 End If
 
 
 'if new password left blank
 If txtNewPassword.Text = "" Then
    MsgBox "Enter Complete Information", vbOKOnly + vbInformation, "Information"
    txtNewPassword.SetFocus
    Exit Sub
 End If
 
 'if new confirmpassword left blank
 If txtConfirmpassword.Text = "" Then
     MsgBox "Enter Complete Information", vbOKOnly + vbInformation, "Information"
     txtConfirmpassword.SetFocus
    Exit Sub
 End If
 
 With login
 .MoveLast
 .MoveFirst
 While Not .EOF
  If .Fields(0) = UserName Then
   If .Fields(1) <> txtOldPassword.Text Then
    MsgBox "Old Password Doesn't Match", vbOKOnly + vbCritical, "Error"
    txtOldPassword.SetFocus
    Exit Sub
   ElseIf txtNewPassword.Text <> txtConfirmpassword.Text Then
    MsgBox "New Password Doesn't Match", vbOKOnly + vbCritical, "Error"
    txtNewPassword.SetFocus
    Exit Sub
   Else
    StrSql = "UPDATE Password_Table SET Pass = '" & txtConfirmpassword.Text & "' " _
             & "WHERE UserName = '" & UserName & "';"
    Cnn.Execute StrSql
    .Update
    MsgBox "Password Has Been Changed Succesfully", vbOKOnly + vbInformation, "Correct"
    Exit Sub
   End If
  End If
  .MoveNext
 Wend
End With
   
   
End Sub


Private Sub Form_Load()

str = MDIForm1.StatusBar.Panels(1).Text

UserName = Mid(str, 14, Len(str))

 Set Cnn = New ADODB.Connection
 Set login = New ADODB.Recordset
 
 
 Cnn.CursorLocation = adUseClient
 Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
          & App.Path & "\data.mdb"
 
login.Open "SELECT * FROM password_table", Cnn, adOpenDynamic, adLockOptimistic
 lblUsername.Caption = "-" & UserName
 
 
 cmdSave.Enabled = False


End Sub







