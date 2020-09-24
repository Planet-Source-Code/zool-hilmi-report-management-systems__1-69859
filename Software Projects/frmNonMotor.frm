VERSION 5.00
Begin VB.Form frmNonMotor 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   -105
   ClientWidth     =   12300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   12300
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   5280
      MouseIcon       =   "frmNonMotor.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmNonMotor.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Save record"
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   6240
      MouseIcon       =   "frmNonMotor.frx":06EA
      MousePointer    =   99  'Custom
      Picture         =   "frmNonMotor.frx":083C
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Cancel"
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   4320
      MaskColor       =   &H8000000F&
      MouseIcon       =   "frmNonMotor.frx":0DBC
      MousePointer    =   99  'Custom
      Picture         =   "frmNonMotor.frx":0F0E
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Add new record"
      Top             =   7200
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00FFC0C0&
      Height          =   1575
      Left            =   3720
      TabIndex        =   17
      Top             =   6840
      Width           =   4695
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
         Left            =   2520
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
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
         Left            =   1560
         TabIndex        =   19
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
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
         TabIndex        =   18
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   5535
      Left            =   480
      TabIndex        =   8
      Top             =   840
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
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3240
         Width           =   4095
      End
      Begin VB.TextBox txtRecordNo 
         BackColor       =   &H00FFC0C0&
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
         Left            =   1680
         TabIndex        =   0
         Top             =   840
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1440
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
         Left            =   1680
         TabIndex        =   5
         Top             =   4680
         Width           =   4095
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
         Left            =   1680
         TabIndex        =   3
         Top             =   2640
         Width           =   4095
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
         TabIndex        =   7
         Top             =   1440
         Width           =   2895
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
         TabIndex        =   6
         Top             =   840
         Width           =   2895
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
         Left            =   1680
         TabIndex        =   2
         Top             =   2040
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "( e.g. 17-Sep-2007 )"
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
         TabIndex        =   25
         Top             =   2760
         Width           =   1815
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
         Left            =   6480
         TabIndex        =   16
         Top             =   1440
         Width           =   735
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
         Left            =   6360
         TabIndex        =   15
         Top             =   840
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
         TabIndex        =   13
         Top             =   3240
         Width           =   975
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
         TabIndex        =   12
         Top             =   2640
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
         TabIndex        =   11
         Top             =   2040
         Width           =   1455
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
         TabIndex        =   10
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add New Record Non-Motor Insurance"
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
      Left            =   3120
      TabIndex        =   21
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmNonMotor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cnn As ADODB.Connection
Dim qc As ADODB.Recordset
Dim Insurance As ADODB.Recordset
Dim num As Integer
Dim strMonth, strYear, str As String
Private Sub cmdAdd_Click()


'********error prevention************'
unlocked_nonmotor
cmdAdd.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = False
cboInsurance.SetFocus
End Sub

Private Sub cmdCancel_Click()
Unload Me
Me.Visible = False
End Sub

Private Sub cmdSave_Click()
If cboInsurance.ListIndex = -1 Then
    MsgBox "Please Select Insurance Type", vbOKOnly + vbInformation, "Correct"
    Exit Sub
End If
With qc
    '**adding value to database**'
     .AddNew
     .Fields(0) = txtRecordNo.Text
     .Fields(1) = cboInsurance.List(cboInsurance.ListIndex)
     .Fields(2) = txtPolicyNo.Text
     .Fields(3) = txtExpiryDate.Text
     .Fields(4) = txtLocation.Text
     .Fields(5) = txtSumInsured.Text
     .Fields(6) = txtPremium.Text
     .Fields(7) = txtRate.Text
     .Update
     MsgBox "Record Has Been Entered Sucessfully", vbOKOnly + vbInformation, "Correct"
     
    Clear
    lock_nonmotor
     num = 100 + .RecordCount + 1
      txtRecordNo.Text = "ID" + CStr(strMonth) _
                       + "-" + CStr(num) + "-" + CStr(strYear)
 End With
 
 cmdAdd.Enabled = True
 cmdSave.Enabled = False
 cmdCancel.Enabled = True
 
 
End Sub

Private Sub Form_Load()
'establish connection

Set Cnn = New ADODB.Connection
Set qc = New ADODB.Recordset
Set Insurance = New ADODB.Recordset

'open connection and data source
Cnn.CursorLocation = adUseClient
Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
         & App.Path & "\data.mdb;Persist Security Info=False"


qc.Open "select * from qc_nonmotor", Cnn, adOpenDynamic, adLockPessimistic
Insurance.Open "select * from insurance_type", Cnn, adOpenDynamic, adLockPessimistic



'******load fields from database**************'
With Insurance
    While Not .EOF
    cboInsurance.AddItem .Fields(0)
    .MoveNext
Wend
End With

'************Generate A New Id******************************'

num = 100
strMonth = Month(Date)
strYear = Year(Date)
str = "ID"

With qc
 If .RecordCount = 0 Then
  txtRecordNo.Text = txtRecordNo.Text + CStr(strMonth) _
                   + "-" + CStr(num) + "-" + CStr(strYear)
    Else
   num = num + .RecordCount + 1
   txtRecordNo.Text = str + CStr(strMonth) _
                   + "-" + CStr(num) + "-" + CStr(strYear)
   
 End If
End With

'*************error prevention**************'

lock_nonmotor

cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = True
End Sub



Private Sub lock_nonmotor()
txtRecordNo.Locked = True
cboInsurance.Locked = True
txtPolicyNo.Locked = True
txtExpiryDate.Locked = True
txtLocation.Locked = True
txtSumInsured.Locked = True
txtPremium.Locked = True
txtRate.Locked = True
End Sub

Private Sub unlocked_nonmotor()
txtRecordNo.Locked = False
cboInsurance.Locked = False
txtPolicyNo.Locked = False
txtExpiryDate.Locked = False
txtLocation.Locked = False
txtSumInsured.Locked = False
txtPremium.Locked = False
txtRate.Locked = False
End Sub

Private Sub Clear()
cboInsurance.ListIndex = -1
txtRecordNo.Text = ""
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

