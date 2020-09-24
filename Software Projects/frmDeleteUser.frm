VERSION 5.00
Begin VB.Form frmDeleteUser 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3750
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   1455
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   3375
      Begin VB.CommandButton cmd_cancel 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   2040
         MouseIcon       =   "frmDeleteUser.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmDeleteUser.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancel"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmd_delete 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   840
         MouseIcon       =   "frmDeleteUser.frx":06D2
         MousePointer    =   99  'Custom
         Picture         =   "frmDeleteUser.frx":0824
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Delete record"
         Top             =   480
         Width           =   735
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
         Left            =   2040
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
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
         Left            =   840
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
   End
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
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delete Username"
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
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   360
      X2              =   6240
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Username"
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
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmDeleteUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cnn As ADODB.Connection
Dim login As ADODB.Recordset
Dim UserName, Pas As String




Private Sub cmd_cancel_Click()
Unload Me
End Sub

Private Sub cmd_delete_Click()
If cboUsername.List(cboUsername.ListIndex) = "Admin" Then
 MsgBox "Sorry,You Cannot Delete Admin", vbOKOnly + vbInformation, "Error"
 Else
    With login
    .MoveLast
    .MoveFirst
        While Not .EOF
            .Delete
             MsgBox "Username Has Been Deleted", vbOKOnly + vbInformation, "Confirm"
             cboUsername.Enabled = False
             cmd_delete.Enabled = False
            Exit Sub
            .MoveNext
     
        Wend
   End With
End If



End Sub

Private Sub Form_Load()
Set Cnn = New ADODB.Connection
Set login = New ADODB.Recordset

Cnn.CursorLocation = adUseClient
Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
        & App.Path & "\data.mdb;Persist Security Info=False"

login.Open "SELECT * FROM password_table", Cnn, adOpenDynamic, adLockPessimistic
cboUsername.Enabled = True

With login
 .MoveLast
 .MoveFirst
  While Not .EOF
   cboUsername.AddItem .Fields(0).Value
   .MoveNext
  Wend
End With
End Sub


