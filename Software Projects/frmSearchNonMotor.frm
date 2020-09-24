VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSearchNonMotor 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   14745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10245
   ScaleWidth      =   14745
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Search By:"
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
      Height          =   3615
      Left            =   5520
      TabIndex        =   8
      Top             =   5640
      Width           =   2655
      Begin VB.OptionButton optLocation 
         BackColor       =   &H00FFC0C0&
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
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optPolicyNo 
         BackColor       =   &H00FFC0C0&
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
         Left            =   480
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton optExpiryDate 
         BackColor       =   &H00FFC0C0&
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
         Left            =   480
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
      End
      Begin VB.OptionButton optInsuranType 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Insurance Type"
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
         Left            =   480
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "(e.g. 12-mar-2007)"
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
         Left            =   720
         TabIndex        =   17
         Top             =   2400
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<<"
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
      Left            =   2400
      TabIndex        =   7
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   ">>"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   6600
      Width           =   855
   End
   Begin VB.ListBox lstTo 
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
      Height          =   3420
      Left            =   3480
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
   End
   Begin VB.ListBox lstFrom 
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
      Height          =   3420
      Left            =   480
      TabIndex        =   4
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Enter Search Info"
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
      Height          =   3615
      Left            =   8400
      TabIndex        =   0
      Top             =   5640
      Width           =   2175
      Begin VB.TextBox txtname 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   2400
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4815
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Non Motor Insurance Search Engine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Available Colum:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Display Column:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   5520
      Width           =   2415
   End
End
Attribute VB_Name = "frmSearchNonMotor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cnn As ADODB.Connection
Dim qc_info As ADODB.Recordset
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long



Private Sub cmdClear_Click()

MSFlexGrid1.Clear
optInsuranType.Value = False
optPolicyNo.Value = False
optExpiryDate.Value = False
optLocation.Value = False
txtname.Text = ""
End Sub


Private Sub cmdSearch_Click()
If optInsuranType.Value = False And _
    optPolicyNo.Value = False And _
    optExpiryDate.Value = False And _
    optLocation.Value = False Then
    MsgBox "Please check any search by field to continue!", vbOKOnly + vbInformation, "Error"
Exit Sub
End If
    
    Dim rs As New ADODB.Recordset
    Dim itemNo As Integer
    Dim item() As String
    Dim i As Integer
    Dim j As Integer
    Dim X As Integer
  

    itemNo = lstTo.ListCount
    ReDim item(itemNo) As String
  
    For i = 0 To lstTo.ListCount - 1
        item(i) = lstTo.List(i)
    Next i
    MSFlexGrid1.Rows = 2
'    MSFlexGrid1.Cols = 1
'    MSFlexGrid1.Clear

    
    'Insurancetype
         If optInsuranType.Value = True Then
            rs.Open "Select * from qc_nonmotor where TypeInsurance='" & txtname.Text & "' ", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
   'search by Policy
         ElseIf optPolicyNo.Value = True Then
            rs.Open "Select * from qc_nonmotor where PolicyNo='" & txtname.Text & "' ", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
     'search by Expiry
        ElseIf optExpiryDate.Value = True Then
            rs.Open "Select * from qc_nonmotor where ExpiryDate='" & txtname.Text & "' ", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
      'search by location
        ElseIf optLocation.Value = True Then
            rs.Open "Select * from qc_nonmotor where Location='" & txtname.Text & "' ", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
     
     
     End If
     
         
    
    'if record not found then
       If rs.EOF Then
            MsgBox "No Record Found!", vbOKOnly + vbInformation, "Info"
            Exit Sub
        End If
 
    
    
    
    
'if record record found continue
    
  With MSFlexGrid1
        .Clear
        .Cols = itemNo
    
        For j = 0 To itemNo - 1
            .Row = 0
            .Col = j
            .Clip = item(j)
        Next
      
        Select Case itemNo
        Case 1:
            Do While Not rs.EOF
                .AddItem rs(item(0))
                rs.MoveNext
            Loop
            .RemoveItem 1
        
        Case 2:
            Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 3:
            Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 4:
             Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 5:
             Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 6:
            Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 7:
            Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 8:
            Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7))
                rs.MoveNext
            Loop
            .RemoveItem 1
        
                   End Select
                   
      End With
      
  
    
    
    'Disable control for error prevention
 
End Sub


Private Sub Command1_Click()

End Sub


Private Sub lstFrom_Click()
    lstTo.ListIndex = -1
End Sub

Private Sub lstTo_Click()
    lstFrom.ListIndex = -1
End Sub

Private Sub lstFrom_DblClick()
    lstTo.AddItem lstFrom.List(lstFrom.ListIndex)
    lstFrom.RemoveItem lstFrom.ListIndex

End Sub

Private Sub lstTo_dblClick()
    lstFrom.AddItem lstTo.List(lstTo.ListIndex)
    lstTo.RemoveItem lstTo.ListIndex
End Sub

Private Sub cmdAdd_Click()
    If lstFrom.ListIndex <> -1 Then
        lstTo.AddItem lstFrom.List(lstFrom.ListIndex)
        lstFrom.RemoveItem lstFrom.ListIndex
    End If
End Sub
Private Sub cmdRemove_Click()
    If lstTo.ListIndex <> -1 Then
        lstFrom.AddItem lstTo.List(lstTo.ListIndex)
        lstTo.RemoveItem lstTo.ListIndex
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo err
    Set Cnn = New ADODB.Connection
    Cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                           "Data Source=" & App.Path & "\data.mdb;" & _
                           "Persist Security Info=False"
                           
                           
                           
With lstFrom
      .Clear
      .AddItem "RecordID"
      .AddItem "TypeInsurance"
      .AddItem "PolicyNo"
      .AddItem "ExpiryDate"
      .AddItem "Location"
      .AddItem "SumInsured"
      .AddItem "Premium"
      .AddItem "Rate"

   End With
    Cnn.Open
    Exit Sub
    
err:
    Dim result As Integer
    result = MsgBox(err.Description, vbDefaultButton1 + vbExclamation, err.Source & " Error")
    

End Sub





