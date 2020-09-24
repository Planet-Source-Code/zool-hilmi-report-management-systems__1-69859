VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSearchMotor 
   BackColor       =   &H00FFC0C0&
   Caption         =   "MotorSearch engine"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
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
      Height          =   4455
      Left            =   5520
      TabIndex        =   13
      Top             =   5640
      Width           =   2655
      Begin VB.OptionButton optDriver 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Driver"
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
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
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
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optLicenseDuration 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   2175
      End
      Begin VB.OptionButton optStatus 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optLicenseExpire 
         BackColor       =   &H00FFC0C0&
         Caption         =   "License Expiry"
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
         Left            =   240
         TabIndex        =   15
         Top             =   3480
         Width           =   2175
      End
      Begin VB.OptionButton optManufacturer 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   2175
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
         Left            =   480
         TabIndex        =   21
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "(e.g. Kuching/Landeh)"
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
         Left            =   480
         TabIndex        =   20
         Top             =   600
         Width           =   1815
      End
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
      TabIndex        =   8
      Top             =   5640
      Width           =   2175
      Begin VB.CommandButton cmdCancel 
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
         Height          =   495
         Left            =   480
         TabIndex        =   12
         Top             =   2520
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
         TabIndex        =   11
         Top             =   1920
         Width           =   1335
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
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
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
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
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
      TabIndex        =   4
      Top             =   5760
      Width           =   1815
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
      TabIndex        =   3
      Top             =   6600
      Width           =   855
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
      TabIndex        =   2
      Top             =   7200
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   8493
      _Version        =   393216
      Rows            =   1
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
      TabIndex        =   7
      Top             =   5520
      Width           =   2415
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
      TabIndex        =   6
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Motor Search Engine"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmSearchMotor"
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



Private Sub LoadFields()


'load information from data
   With lstFrom
      .Clear
      .AddItem "RegistrationNo"
      .AddItem "EngineNo"
      .AddItem "CasisNo"
      .AddItem "Manufacturer"
      .AddItem "ModelName"
      .AddItem "EngineHorsepower"
      .AddItem "FuelType"
      .AddItem "VehicleType"
      .AddItem "YearManufacturer"
      .AddItem "YearRegistration"
      .AddItem "LicenseDuration"
      .AddItem "BDMBGK"
      .AddItem "BTM"
      .AddItem "DriverName"
      .AddItem "Location"
      .AddItem "SeriesNo"
      .AddItem "Status"
      .AddItem "LPKPFormNo"
      .AddItem "FileReferenceNo"
      .AddItem "LicenseReferenceNo"
      .AddItem "LicenseExpire"
      .AddItem "PolicyNoInsurance"
      .AddItem "PolicyPeriod"
      .AddItem "VehicleSumInsured"
       .AddItem "ExcessApplicable"
      .AddItem "BasicPremium"
      .AddItem "NCTDTT"
      .AddItem "TotalPayable"
      .AddItem "InspectioDate"
      .AddItem "InspectionTime"
   End With
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClear_Click()
MSFlexGrid1.Clear
txtname.Text = ""
optLocation.Value = False
optDriver.Value = False
optManufacturer.Value = False
optLicenseDuration.Value = False
optStatus.Value = False
optLicenseExpire.Value = False

End Sub


Private Sub cmdSearch_Click()
MSFlexGrid1.Clear
If optLocation.Value = False And optDriver.Value = False And optManufacturer.Value = False And _
    optLicenseDuration.Value = False And optStatus.Value = False And optLicenseExpire.Value = False Then
      MsgBox "Please Select At Least One Search Option", vbOKOnly + vbInformation, "Invalid"

    Else
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
 


         If optLocation.Value = True Then
           rs.Open "Select * from qc_motor WHERE location='" & txtname.Text & "' ", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
       
               ElseIf optDriver.Value = True Then
               rs.Open "Select * from QC_motor where DriverName='" & txtname.Text & "' ", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
     
               ElseIf optManufacturer.Value = True Then
               rs.Open "Select * from QC_motor where Manufacturer='" & txtname.Text & "' ", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
        
               ElseIf optLicenseDuration.Value = True Then
               rs.Open "Select * from QC_motor where LicenseDuration='" & txtname.Text & "' ", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
         
               ElseIf optStatus.Value = True Then
               rs.Open "Select * from QC_motor where Status='" & txtname.Text & "' ", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
        
               ElseIf optLicenseExpire.Value = True Then
               rs.Open "Select * from QC_motor where LicenseExpire='" & txtname.Text & "' ", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
          End If
          
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
        Case 9:
            Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 10:
            Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 11:
            Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 12:
            Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 13:
            Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 14:
            Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 15:
        Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                 vbTab & rs(item(14))
                rs.MoveNext
            Loop
            .RemoveItem 1
            Case 16
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 17
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                 vbTab & rs(item(16))
                rs.MoveNext
            Loop
            .RemoveItem 1
       Case 18
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                vbTab & rs(item(16)) & _
                vbTab & rs(item(17))
                rs.MoveNext
            Loop
            .RemoveItem 1
                   Case 19
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                vbTab & rs(item(16)) & _
                vbTab & rs(item(17)) & _
                vbTab & rs(item(18))
                rs.MoveNext
            Loop
            .RemoveItem 1
            
            Case 20
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                vbTab & rs(item(16)) & _
                vbTab & rs(item(17)) & _
                vbTab & rs(item(18)) & _
                vbTab & rs(item(19))
                rs.MoveNext
            Loop
            .RemoveItem 1
             Case 21
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                vbTab & rs(item(16)) & _
                vbTab & rs(item(17)) & _
                vbTab & rs(item(18)) & _
                vbTab & rs(item(19)) & _
                vbTab & rs(item(20))
                rs.MoveNext
            Loop
            .RemoveItem 1
    Case 22
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                vbTab & rs(item(16)) & _
                vbTab & rs(item(17)) & _
                vbTab & rs(item(18)) & _
                vbTab & rs(item(19)) & _
                vbTab & rs(item(20)) & _
                vbTab & rs(item(21))
                rs.MoveNext
            Loop
            .RemoveItem 1
              Case 23
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                vbTab & rs(item(16)) & _
                vbTab & rs(item(17)) & _
                vbTab & rs(item(18)) & _
                vbTab & rs(item(19)) & _
                vbTab & rs(item(20)) & _
                vbTab & rs(item(21)) & _
                vbTab & rs(item(22))
                rs.MoveNext
            Loop
            .RemoveItem 1
        Case 24
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                vbTab & rs(item(16)) & _
                vbTab & rs(item(17)) & _
                vbTab & rs(item(18)) & _
                vbTab & rs(item(19)) & _
                vbTab & rs(item(20)) & _
                vbTab & rs(item(21)) & _
                vbTab & rs(item(22)) & _
                vbTab & rs(item(23))
                rs.MoveNext
            Loop
            .RemoveItem 1
            
                    Case 25
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                vbTab & rs(item(16)) & _
                vbTab & rs(item(17)) & _
                vbTab & rs(item(18)) & _
                vbTab & rs(item(19)) & _
                vbTab & rs(item(20)) & _
                vbTab & rs(item(21)) & _
                vbTab & rs(item(22)) & _
                vbTab & rs(item(23)) & _
                vbTab & rs(item(24))
                rs.MoveNext
            Loop
            .RemoveItem 1
            
            Case 26
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                vbTab & rs(item(16)) & _
                vbTab & rs(item(17)) & _
                vbTab & rs(item(18)) & _
                vbTab & rs(item(19)) & _
                vbTab & rs(item(20)) & _
                vbTab & rs(item(21)) & vbTab & rs(item(22)) & vbTab & rs(item(23)) & _
                vbTab & rs(item(24)) & vbTab & rs(item(25))
                rs.MoveNext
            Loop
            .RemoveItem 1
            
                        Case 27
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                vbTab & rs(item(16)) & _
                vbTab & rs(item(17)) & _
                vbTab & rs(item(18)) & _
                vbTab & rs(item(19)) & _
                vbTab & rs(item(20)) & _
                vbTab & rs(item(21)) & vbTab & rs(item(22)) & vbTab & rs(item(23)) & _
                vbTab & rs(item(24)) & vbTab & rs(item(25)) & vbTab & rs(item(26))
                rs.MoveNext
            Loop
            .RemoveItem 1
            
        Case 28
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                vbTab & rs(item(16)) & _
                vbTab & rs(item(17)) & _
                vbTab & rs(item(18)) & _
                vbTab & rs(item(19)) & _
                vbTab & rs(item(20)) & _
                vbTab & rs(item(21)) & vbTab & rs(item(22)) & vbTab & rs(item(23)) & _
                vbTab & rs(item(24)) & vbTab & rs(item(25)) & vbTab & rs(item(26)) & vbTab & rs(item(27))
                rs.MoveNext
            Loop
            .RemoveItem 1
            
             Case 29
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                vbTab & rs(item(16)) & _
                vbTab & rs(item(17)) & _
                vbTab & rs(item(18)) & _
                vbTab & rs(item(19)) & _
                vbTab & rs(item(20)) & _
                vbTab & rs(item(21)) & vbTab & rs(item(22)) & vbTab & rs(item(23)) & _
                vbTab & rs(item(24)) & vbTab & rs(item(25)) & vbTab & rs(item(26)) & vbTab & rs(item(27)) & vbTab & rs(item(28))
                rs.MoveNext
            Loop
            .RemoveItem 1
            
               
             Case 30
    Do While Not rs.EOF
                .AddItem rs(item(0)) & _
                vbTab & rs(item(1)) & _
                vbTab & rs(item(2)) & _
                vbTab & rs(item(3)) & _
                vbTab & rs(item(4)) & _
                vbTab & rs(item(5)) & _
                vbTab & rs(item(6)) & _
                vbTab & rs(item(7)) & _
                vbTab & rs(item(8)) & _
                vbTab & rs(item(9)) & _
                vbTab & rs(item(10)) & _
                vbTab & rs(item(11)) & _
                vbTab & rs(item(12)) & _
                vbTab & rs(item(13)) & _
                vbTab & rs(item(14)) & _
                vbTab & rs(item(15)) & _
                vbTab & rs(item(16)) & _
                vbTab & rs(item(17)) & _
                vbTab & rs(item(18)) & _
                vbTab & rs(item(19)) & _
                vbTab & rs(item(20)) & _
                vbTab & rs(item(21)) & vbTab & rs(item(22)) & vbTab & rs(item(23)) & _
                vbTab & rs(item(24)) & vbTab & rs(item(25)) & vbTab & rs(item(26)) & vbTab & rs(item(27)) & vbTab & rs(item(28)) & vbTab & rs(item(29))
                rs.MoveNext
            Loop
            .RemoveItem 1
            
           
                   End Select
                
      End With
      
End If
    

 
End Sub


Private Sub lstFrom_DblClick()
     lstTo.ListIndex = -1
    lstTo.AddItem lstFrom.List(lstFrom.ListIndex)
    lstFrom.RemoveItem lstFrom.ListIndex
    

End Sub

Private Sub lstTo_dblClick()
        lstFrom.ListIndex = -1
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
  qc_cnn
'load connection called qc_cnn
Set qc_info = New ADODB.Recordset
qc_info.Open "select * from location", Cnn, adOpenDynamic, adLockPessimistic

'cal from database to quick link


    qc_cnn
    LoadFields
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 2
    
    
End Sub


Private Sub qc_cnn()
    On Error GoTo err
    Set Cnn = New ADODB.Connection
    Cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                           "Data Source=" & App.Path & "\data.mdb;" & _
                           "Persist Security Info=False"
    Cnn.Open
    Exit Sub
    
err:
    Dim result As Integer
    result = MsgBox(err.Description, vbDefaultButton1 + vbExclamation, err.Source & " Error")
    
End Sub

Private Sub ExitInit()
    Cnn.Close
End Sub

