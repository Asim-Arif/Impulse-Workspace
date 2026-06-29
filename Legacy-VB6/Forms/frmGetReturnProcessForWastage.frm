VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmGetReturnProcessForWastage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Process to Return to"
   ClientHeight    =   7545
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10290
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetReturnProcessForWastage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptChargeTo 
      Caption         =   "Charge to No One"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   4635
      TabIndex        =   7
      Top             =   165
      Width           =   2250
   End
   Begin VB.OptionButton OptChargeTo 
      Caption         =   "Charge to Employee"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   2355
      TabIndex        =   6
      Top             =   165
      Width           =   2250
   End
   Begin VB.OptionButton OptChargeTo 
      Caption         =   "Charge to Maker"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   75
      TabIndex        =   5
      Top             =   165
      Value           =   -1  'True
      Width           =   2250
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   75
      TabIndex        =   4
      Top             =   585
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Enabled         =   0   'False
      Appearance      =   0
   End
   Begin VB.CheckBox chkChargeToEmployee 
      Caption         =   "Charge to Employee"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4335
      TabIndex        =   3
      Top             =   7215
      Visible         =   0   'False
      Width           =   2040
   End
   Begin MSComctlLib.ListView LV 
      Height          =   6150
      Left            =   75
      TabIndex        =   2
      Top             =   855
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   10848
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Process Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Process"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Maker"
         Object.Width           =   10134
      EndProperty
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   9000
      TabIndex        =   1
      Top             =   7065
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   435
      Left            =   7710
      TabIndex        =   0
      Top             =   7065
      Width           =   1215
   End
End
Attribute VB_Name = "frmGetReturnProcessForWastage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim lEntryID As Long, lVendID As Long
Dim strLotNo As String, strEmpID As String

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub chkChargeToEmployee_Click()
'
'    If chkChargeToEmployee.Value = vbChecked Then
'        cmbEmp.Enabled = True
'    Else
'        cmbEmp.Enabled = False
'    End If
'
'    If cmbEmp.Tag = "" Then
'        cmbEmp.AddVals con, "'{'+EmpID+'} ' + Name + ' ' + Rel + ' ' + FName", "Employees", "EmpID"
'        cmbEmp.Tag = "Filled"
'    End If
'
End Sub

Private Sub LV_DblClick()
'    lEntryID = Val(LV.SelectedItem.key)
'    lVendID = Val(LV.SelectedItem.ListSubItems(2).Tag)
'    bCancelled = False
'    Unload Me
End Sub

Private Sub OKButton_Click()
     
     
    Dim i As Integer
     
    If OptChargeTo(1).Value Then
        If cmbEmp.MatchFound = False Then
            MsgBox "Please Select Employee.", vbInformation
            Exit Sub
        End If
    ElseIf OptChargeTo(0).Value Then
        If LV.SelectedItem Is Nothing Then
            MsgBox "Please Select Vender to Charge to.", vbInformation
            Exit Sub
        End If
    End If
    bCancelled = False
    
    If OptChargeTo(2).Value Then
        strEmpID = ""
        lEntryID = 0
        lVendID = 0
    ElseIf OptChargeTo(1).Value Then
        strEmpID = cmbEmp.ID
        lEntryID = 0
        lVendID = 0
    Else
        lEntryID = Val(LV.SelectedItem.key)
        lVendID = Val(LV.SelectedItem.ListSubItems(2).Tag)
        strEmpID = ""
    End If
    
    Unload Me
    
End Sub

Public Function getData(p_strLotNo As String, ByRef p_lEntryID As Long, ByRef p_lVendID As Long, ByRef p_strEmpID As String) As Boolean
    
    strLotNo = p_strLotNo
    
    Dim i As Integer, strMakerList As String
    Dim rs As New ADODB.Recordset, ITM As ListItem, strMakerName As String, lForeColor As Long
    With rs
        
        '.Open "SELECT * FROM VVendRcvdDetail WHERE LotNo='" & strLotNo & "' AND EntryID<=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT * FROM VVendRcvdDetail WHERE LotNo IN(SELECT LotNo FROM dbo.GetAllLots_F('" & strLotNo & "')) AND EntryID<=" & p_lEntryID, con_ServerSide, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        Do Until .EOF
            Set ITM = LV.ListItems.add(, Val(!EntryID & "") & "'", !code & "")
            ITM.ListSubItems.add , , !Description & ""
            If Val(!VendID & "") = iFactoryMakerID Then
                strMakerName = con.Execute("SELECT dbo.VendReceived_Employees_F(" & Val(!ReceivedEntryID) & ")").Fields(0).Value & "" '"(" & !EmpId & ") " & !EmployeeName
                lForeColor = vbRed
            Else
                strMakerName = !VenderName & ""
                lForeColor = vbBlue
            End If
            
            ITM.ListSubItems.add(, , strMakerName).Tag = Val(!VendID & "")
            ITM.ForeColor = lForeColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = lForeColor
            Next
            
            strMakerList = strMakerList & Val(!VendID & "") & ","
            .MoveNEXT
        Loop
        .Close
        If strMakerList <> "" Then
            strMakerList = Left(strMakerList, Len(strMakerList) - 1)
        End If
        
        .Open "SELECT * FROM Makers WHERE VendID NOT IN(" & strMakerList & ") ORDER BY VenderName", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set ITM = LV.ListItems.add(, "", "-")
            ITM.ListSubItems.add , , "-"
            
            strMakerName = !VenderName & ""
            lForeColor = RGB(25, 175, 225)
            
            ITM.ListSubItems.add(, , strMakerName).Tag = Val(!VendID & "")
            ITM.ForeColor = lForeColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = lForeColor
            Next
           
            .MoveNEXT
        Loop
        .Close
        
    End With
    Set rs = Nothing
   
    
    Me.Show 1
    
    p_lEntryID = lEntryID
    p_lVendID = lVendID
    p_strEmpID = strEmpID
    
    getData = Not bCancelled
    
End Function

Private Sub OptChargeTo_Click(Index As Integer)

    If OptChargeTo(1).Value Then
        cmbEmp.Enabled = True
        If cmbEmp.Tag = "" Then
            cmbEmp.AddVals con, "'{'+EmpID+'} ' + Name + ' ' + Rel + ' ' + FName", "Employees", "EmpID"
            cmbEmp.Tag = "Filled"
        End If
    Else
        cmbEmp.Enabled = False
    End If
   
End Sub
