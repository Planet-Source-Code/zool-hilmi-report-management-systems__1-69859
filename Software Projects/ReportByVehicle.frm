VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFilterReportNonmotor 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   10440
   ClientLeft      =   75
   ClientTop       =   75
   ClientWidth     =   13290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10440
   ScaleWidth      =   13290
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5055
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   8916
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
      Left            =   6000
      TabIndex        =   6
      Top             =   7320
      Width           =   1575
   End
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
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   7800
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
      Left            =   6000
      TabIndex        =   4
      Top             =   6840
      Width           =   1575
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
      Height          =   3180
      Left            =   360
      TabIndex        =   3
      Top             =   6240
      Width           =   1935
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
      Left            =   3360
      TabIndex        =   2
      Top             =   6240
      Width           =   1935
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
      TabIndex        =   1
      Top             =   7200
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
      TabIndex        =   0
      Top             =   7800
      Width           =   855
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   5760
      TabIndex        =   7
      Top             =   6240
      Width           =   2055
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
      Left            =   3360
      TabIndex        =   10
      Top             =   6000
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
      Left            =   360
      TabIndex        =   9
      Top             =   6000
      Width           =   1935
   End
End
Attribute VB_Name = "frmFilterReportNonmotor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cnn As ADODB.Connection
Dim rs As ADODB.Recordset



Private Sub cmdAdd_Click()
   If lstFrom.ListIndex <> -1 Then
        lstTo.AddItem lstFrom.List(lstFrom.ListIndex)
        lstFrom.RemoveItem lstFrom.ListIndex
    End If
End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdExport_Click()
FlexToExcel
End Sub

Private Sub cmdRemove_Click()
    If lstTo.ListIndex <> -1 Then
        lstFrom.AddItem lstTo.List(lstTo.ListIndex)
        lstTo.RemoveItem lstTo.ListIndex
    End If
End Sub

Private Sub cmdView_Click()
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
    MSFlexGrid1.Rows = 3

   
    rs.Open "Select * from qc_nonmotor", Cnn, adOpenDynamic, adLockOptimistic, adCmdText
              
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
                   End Select
                   .RemoveItem 1
                   
      End With
End Sub

Private Sub Form_Load()
'populate fields to listbox
 LoadFields
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



'populate fields to listbox
 LoadFields


End Sub






Private Sub LoadFields()


'**********'load information from data
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

