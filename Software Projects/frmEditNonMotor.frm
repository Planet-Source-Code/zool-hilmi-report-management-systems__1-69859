VERSION 5.00
Begin VB.Form frmEditNonMotor 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   12270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   1575
      Left            =   3240
      TabIndex        =   19
      Top             =   7920
      Width           =   6135
      Begin VB.CommandButton cmdEdit 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   1560
         MaskColor       =   &H8000000F&
         MouseIcon       =   "frmEditNonMotor.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmEditNonMotor.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Add new record"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   3960
         MouseIcon       =   "frmEditNonMotor.frx":0730
         MousePointer    =   99  'Custom
         Picture         =   "frmEditNonMotor.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   2760
         MouseIcon       =   "frmEditNonMotor.frx":0E02
         MousePointer    =   99  'Custom
         Picture         =   "frmEditNonMotor.frx":0F54
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Save record"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
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
         Left            =   2760
         TabIndex        =   21
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
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
         Left            =   3960
         TabIndex        =   20
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.ComboBox cboNo 
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
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   5775
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   10575
      Begin VB.TextBox txtLocation 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3240
         Width           =   4095
      End
      Begin VB.ComboBox cboInsurance 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox txtPolicyNo 
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
         Left            =   1800
         TabIndex        =   3
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox txtPremium 
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
         Left            =   7200
         TabIndex        =   7
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtRate 
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
         Left            =   7200
         TabIndex        =   8
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox txtExpiryDate 
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
         Left            =   1800
         TabIndex        =   4
         Top             =   2640
         Width           =   4095
      End
      Begin VB.TextBox txtSumInsured 
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
         Left            =   1800
         TabIndex        =   6
         Top             =   4680
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Record ID"
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
         Left            =   2040
         TabIndex        =   24
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Type Insurance"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Policy No"
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
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date"
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
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sum Insured"
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
         Left            =   240
         TabIndex        =   14
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Premium "
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
         Left            =   6240
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Left            =   6360
         TabIndex        =   12
         Top             =   2040
         Width           =   735
      End
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Edit Non Motor Insurance"
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
      Left            =   3600
      TabIndex        =   23
      Top             =   600
      Width           =   4935
   End
End
Attribute VB_Name = "frmEditNonMotor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim InsuranceType As ADODB.Recordset



Private Sub cboNo_Click()
If cboNo.ListIndex = -1 Then
 cmdEdit.Enabled = False
Exit Sub
End If



 cmdEdit.Enabled = True
 With rs
    .Requery
    .MoveLast
    .MoveFirst
                    'you reload info update it
        While Not .EOF = True
            If cboNo.List(cboNo.ListIndex) = .Fields(0) Then
                cboInsurance.Text = .Fields(1)
                txtPolicyNo.Text = .Fields(2)
                txtExpiryDate.Text = .Fields(3)
                txtLocation.Text = .Fields(4)
                txtSumInsured.Text = .Fields(5)
                txtPremium.Text = .Fields(6)
                txtRate.Text = .Fields(7)
            End If
              .MoveNext
        Wend
 End With
 
 
 
 End Sub

Private Sub cmdCancel_Click()
Unload Me


End Sub

Private Sub cmdEdit_Click()

'***********Error Prevention***********'
unlocked_nonmotor
cmdEdit.Enabled = False
cmdUpdate.Enabled = True
cmdCancel.Enabled = True
cboInsurance.SetFocus

End Sub

Private Sub cmdUpdate_Click()

Dim StrSql As String


 With rs
  .MoveLast
  .MoveFirst
  While .EOF = False
    If cboNo.List(cboNo.ListIndex) = .Fields(0) Then
       StrSql = "UPDATE qc_nonmotor SET TypeInsurance = '" & cboInsurance.Text & "'," _
                & "PolicyNo = '" & txtPolicyNo.Text & "'," _
                & "ExpiryDate = '" & txtExpiryDate.Text & "'," _
                & "Location = '" & txtLocation.Text & "'," _
                & "SumInsured = '" & txtSumInsured.Text & "'," _
                & "Premium = '" & txtPremium.Text & "'," _
                & "Rate = '" & txtRate.Text & "' " _
                & "WHERE RecordID = '" & cboNo.List(cboNo.ListIndex) & "';"
            
        Cnn.Execute StrSql
      .Update
    MsgBox "Information Has Been Entered Succesfully", vbOKOnly + vbInformation, "Correct"

    .MoveNext
   Else
   .MoveNext
   End If
  Wend
 End With
            
            
            
'**************Error prevention**********'

Clear

cmdEdit.Enabled = True
cmdUpdate.Enabled = False
cmdCancel.Enabled = True


                
End Sub

Private Sub Form_Load()
Set Cnn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set InsuranceType = New ADODB.Recordset

'open connection and data source
Cnn.CursorLocation = adUseClient
Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
         & App.Path & "\data.mdb;Persist Security Info=False"

rs.Open "select * from qc_nonmotor", Cnn, adOpenDynamic, adLockPessimistic
InsuranceType.Open "SELECT * FROM insurance_type", Cnn, adOpenDynamic, adLockPessimistic

'cal from database to quick link
With rs
    While Not .EOF
    cboNo.AddItem .Fields(0)
    .MoveNext
Wend
End With
'To populate insurance
With InsuranceType
    While Not .EOF
    cboInsurance.AddItem .Fields(0)
    .MoveNext
Wend
End With


'******************Error Prevention********'

lock_nonmotor
 cmdEdit.Enabled = False
 cmdUpdate.Enabled = False
 cmdCancel.Enabled = True

End Sub

Private Sub lock_nonmotor()

cboInsurance.Locked = True
txtPolicyNo.Locked = True
txtExpiryDate.Locked = True
txtLocation.Locked = True
txtSumInsured.Locked = True
txtPremium.Locked = True
txtRate.Locked = True
End Sub

Private Sub unlocked_nonmotor()

cboInsurance.Locked = False
txtPolicyNo.Locked = False
txtExpiryDate.Locked = False
txtLocation.Locked = False
txtSumInsured.Locked = False
txtPremium.Locked = False
txtRate.Locked = False
End Sub

Private Sub Clear()
cboNo.ListIndex = -1
cboInsurance.ListIndex = -1
txtPolicyNo.Text = ""
txtExpiryDate.Text = ""
txtLocation.Text = ""
txtSumInsured.Text = ""
txtPremium.Text = ""
txtRate.Text = ""
End Sub





Private Sub txtExpiryDate_LostFocus()
txtExpiryDate.Text = Format(txtExpiryDate.Text, "dd-mmm-yyyy")
End Sub
