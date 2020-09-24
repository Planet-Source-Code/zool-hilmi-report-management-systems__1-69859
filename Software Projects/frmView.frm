VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFilterReportMotor 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Report By Location"
   ClientHeight    =   9945
   ClientLeft      =   360
   ClientTop       =   450
   ClientWidth     =   14175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   14175
   WindowState     =   2  'Maximized
   Begin VB.CheckBox checkLocation 
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
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   6840
      TabIndex        =   12
      Top             =   9000
      Width           =   1575
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
      Height          =   3180
      ItemData        =   "frmView.frx":0000
      Left            =   3240
      List            =   "frmView.frx":0007
      MultiSelect     =   1  'Simple
      TabIndex        =   10
      Top             =   6120
      Width           =   1935
   End
   Begin VB.ComboBox cbovehicle 
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
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
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
      Left            =   6840
      TabIndex        =   8
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export To Excel"
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
      Left            =   6840
      TabIndex        =   7
      Top             =   8520
      Width           =   1575
   End
   Begin VB.ComboBox cbolocation 
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
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   7320
      Width           =   1575
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
      Left            =   2280
      TabIndex        =   2
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
      Left            =   2280
      TabIndex        =   1
      Top             =   6720
      Width           =   855
   End
   Begin VB.ListBox lstFrom 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Location:"
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
      Height          =   1335
      Left            =   5640
      TabIndex        =   4
      Top             =   6480
      Width           =   4215
      Begin VB.CheckBox checkVehicle 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Vehicle Type"
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
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5775
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   10186
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
      BackColor       =   &H00FFC0C0&
      Caption         =   "(To Return Result without Filtering,leave two check boxes unchecked)"
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
      Left            =   5760
      TabIndex        =   15
      Top             =   6120
      Width           =   5535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   5640
      X2              =   9720
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Display Column:"
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
      Left            =   3240
      TabIndex        =   6
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Available Colum:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   1935
   End
End
Attribute VB_Name = "frmFilterReportMotor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cnn As ADODB.Connection
Dim qc_location As ADODB.Recordset
Dim qc_vehicle As ADODB.Recordset
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook






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

Private Sub cmdExport_Click()
FlexToExcel
End Sub



Private Sub cmdView_Click()

If lstTo.ListCount <= 0 Then
   MsgBox "Please Select At Least One Column", vbOKOnly + vbInformation, "Invalid"

    
        
                    
            Else
            
            Dim rs As New ADODB.Recordset
            Dim itemNo As Integer
            Dim item() As String
            Dim i As Integer
            Dim j As Integer
            Dim X As Integer
            Dim Z As Long
          
        
            itemNo = lstTo.ListCount
            ReDim item(itemNo) As String
          
            For i = 0 To lstTo.ListCount
                item(i) = lstTo.List(i)
            Next i
            MSFlexGrid1.Rows = 2

        
           'if user select one option(location)
            If checkLocation.Value = 1 And checkVehicle.Value = 0 Then
                rs.Open "Select * from QC_motor where location='" & cbolocation.Text & "' ", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
                    
            'if user select one option(Vehicle)
                ElseIf checkVehicle.Value = 1 And checkLocation.Value = 0 Then
                    rs.Open "Select * from QC_motor where VehicleType='" & cbovehicle.Text & "'", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
                 
             'if user select two option(Location and vehicle)
                    ElseIf checkVehicle.Value = 1 And checkLocation.Value = 1 Then
                    rs.Open "Select * from QC_motor where VehicleType='" & cbovehicle.Text & "'AND Location='" & cbolocation.Text & "' ", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
                    
                    ElseIf checkVehicle.Value = 0 And checkLocation.Value = 0 Then
                                If MsgBox("This Will Return A Result Without Filter ?", vbYesNo + vbInformation, "Warning") = vbYes Then
                                cbolocation.ListIndex = -1
                                cbovehicle.ListIndex = -1
                                rs.Open "Select * from QC_motor  ", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
                                    Else
                                    MsgBox "Please Check At Least One Option (Location/Vehicle Types)", vbOKOnly + vbInformation, "Correct"

                                    
                                    Exit Sub
                                End If
                                
                    
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
                 Z = 0
                  Select Case itemNo
                  Case 1:
                    Do While Not rs.EOF
                        Z = Z + 1
                           .AddItem Z
                          rs.MoveNext
                      Loop
                      .RemoveItem 1
                  
                  Case 2:
                      Do While Not rs.EOF
                           Z = Z + 1
                          .AddItem Z & _
                          vbTab & rs(item(1))
                          rs.MoveNext
                      Loop
                      .RemoveItem 1
                  Case 3:
                    Do While Not rs.EOF
                            Z = Z + 1
                          .AddItem Z & _
                          vbTab & rs(item(1)) & _
                          vbTab & rs(item(2))
                          rs.MoveNext
                      Loop
                      .RemoveItem 1
                  Case 4:
                    Do While Not rs.EOF
                              Z = Z + 1
                          .AddItem Z & _
                          vbTab & rs(item(1)) & _
                          vbTab & rs(item(2)) & _
                          vbTab & rs(item(3))
                          rs.MoveNext
                      Loop
                      .RemoveItem 1
                  Case 5:
                    Do While Not rs.EOF
                              Z = Z + 1
                          .AddItem Z & _
                          vbTab & rs(item(1)) & _
                          vbTab & rs(item(2)) & _
                          vbTab & rs(item(3)) & _
                          vbTab & rs(item(4))
                          rs.MoveNext
                      Loop
                      .RemoveItem 1
                  Case 6:
                    Do While Not rs.EOF
                             Z = Z + 1
                          .AddItem Z & _
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
                              Z = Z + 1
                          .AddItem Z & _
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
                              Z = Z + 1
                          .AddItem Z & _
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
                             Z = Z + 1
                          .AddItem Z & _
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
                              Z = Z + 1
                          .AddItem Z & _
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
                              Z = Z + 1
                          .AddItem Z & _
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
                              Z = Z + 1
                          .AddItem Z & _
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
                              Z = Z + 1
                          .AddItem Z & _
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
                             Z = Z + 1
                          .AddItem Z & _
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
                              Z = Z + 1
                          .AddItem Z & _
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
                Case 16:
                  Do While Not rs.EOF
                             Z = Z + 1
                          .AddItem Z & _
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
                  Case 17:
                    Do While Not rs.EOF
                              Z = Z + 1
                          .AddItem Z & _
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
                 Case 18:
                    Do While Not rs.EOF
                              Z = Z + 1
                          .AddItem Z & _
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
                    Case 19:
                      Do While Not rs.EOF
                            Z = Z + 1
                          .AddItem Z & _
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
                      
                  Case 20:
                     Do While Not rs.EOF
                            Z = Z + 1
                          .AddItem Z & _
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
                Case 21:
                  Do While Not rs.EOF
                            Z = Z + 1
                          .AddItem Z & _
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
                 Case 22:
                   Do While Not rs.EOF
                            Z = Z + 1
                          .AddItem Z & _
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
                  Case 23:
                    Do While Not rs.EOF
                           Z = Z + 1
                          .AddItem Z & _
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
                  Case 24:
                    Do While Not rs.EOF
                            Z = Z + 1
                          .AddItem Z & _
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
                      
                  Case 25:
                    Do While Not rs.EOF
                           Z = Z + 1
                          .AddItem Z & _
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
                      
                  Case 26:
                    Do While Not rs.EOF
                            Z = Z + 1
                          .AddItem Z & _
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
                      
                    Case 27:
                      Do While Not rs.EOF
                            Z = Z + 1
                          .AddItem Z & _
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
                      
                  Case 28:
                    Do While Not rs.EOF
                            Z = Z + 1
                          .AddItem Z & _
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
                          vbTab & rs(item(24)) & vbTab & rs(item(25)) & vbTab & rs(item(26)) & _
                          vbTab & rs(item(27))
                          rs.MoveNext
                      Loop
                      .RemoveItem 1
                      
                    Case 29:
                     Do While Not rs.EOF
                           Z = Z + 1
                          .AddItem Z & _
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
                          vbTab & rs(item(24)) & vbTab & rs(item(25)) & vbTab & rs(item(26)) & vbTab & rs(item(27)) & _
                          vbTab & rs(item(28))
                          rs.MoveNext
                     
                        Loop
                        .RemoveItem 1
                        
                    Case 30:
                     Do While Not rs.EOF
                           Z = Z + 1
                          .AddItem Z & _
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
                          vbTab & rs(item(24)) & vbTab & rs(item(25)) & vbTab & rs(item(26)) & vbTab & rs(item(27)) & _
                          vbTab & rs(item(28)) & vbTab & rs(item(29))
                          rs.MoveNext
                        
                        Loop
                        .RemoveItem 1
                        
                       Case 31:
                     Do While Not rs.EOF
                           Z = Z + 1
                          .AddItem Z & _
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
                          vbTab & rs(item(24)) & vbTab & rs(item(25)) & vbTab & rs(item(26)) & vbTab & rs(item(27)) & _
                          vbTab & rs(item(28)) & vbTab & rs(item(29)) & vbTab & rs(item(30))
                         rs.MoveNext
                           Loop
                        .RemoveItem 1
                          End Select
            End With
End If
End Sub

Private Sub Label4_Click()

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
Set qc_location = New ADODB.Recordset
Set qc_vehicle = New ADODB.Recordset

qc_location.Open "select * from location", Cnn, adOpenDynamic, adLockPessimistic
qc_vehicle.Open "select * from vehicle_body", Cnn, adOpenDynamic, adLockPessimistic

'cal location table from database
With qc_location
    While Not .EOF
    cbolocation.AddItem .Fields(0)
    .MoveNext
Wend
End With


'cal vehicle table from databse
With qc_vehicle
    While Not .EOF
    cbovehicle.AddItem .Fields(0)
    .MoveNext
Wend
End With

    qc_cnn
    LoadFields
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 2

    
'********Error Prevetion********'



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
Private Sub FlexToExcel()

'************Code Quoted from vbcode.com without modification*********

      Dim xlObject    As Excel.Application
      Dim xlWB        As Excel.Workbook
          Set xlObject = New Excel.Application
          'This Adds a new woorkbook, you could open the workbook from file also
          Set xlWB = xlObject.Workbooks.Add
          Clipboard.Clear 'Clear the Clipboard
          With MSFlexGrid1
              'Select Full Contents (You could also select partial content)
              .Col = 0               'From first column
              .Row = 0               'From first Row (header)
            .ColSel = .Cols - 1    'Select all columns
              .RowSel = .Rows - 1    'Select all rows
              Clipboard.SetText .Clip 'Send to Clipboard
          End With

          With xlObject.ActiveWorkbook.ActiveSheet
              .Range("A1").Select 'Select Cell A1 (will paste from here, to different cells)
              .Paste              'Paste clipboard contents
          End With

         
          ' This makes Excel visible
          xlObject.Visible = True
End Sub


