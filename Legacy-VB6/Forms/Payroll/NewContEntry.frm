VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form NewContEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance Sheet..."
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleMode       =   0  'User
   ScaleWidth      =   6516.315
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbCont 
      Height          =   285
      Left            =   2850
      TabIndex        =   0
      Top             =   930
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   75
      TabIndex        =   11
      Top             =   0
      Width           =   7155
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Contractor's Entry"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Index           =   1
         Left            =   1560
         TabIndex        =   12
         Top             =   120
         Width           =   3720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Contractor's Entry"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Index           =   2
         Left            =   1575
         TabIndex        =   13
         Top             =   135
         Width           =   3720
      End
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   375
      TabIndex        =   7
      Top             =   930
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   5250
      TabIndex        =   6
      Top             =   960
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "ddd-MMM-yyyy"
      Format          =   61276163
      CurrentDate     =   37055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5730
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NewContEntry.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NewContEntry.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5370
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame FLV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      Left            =   90
      TabIndex        =   4
      Top             =   1260
      Width           =   7110
      Begin VB.TextBox txtQtyEdit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   3630
         TabIndex        =   18
         Top             =   1680
         Visible         =   0   'False
         Width           =   1035
      End
      Begin ComboList.Usercontrol1 cmbItem 
         Height          =   285
         Left            =   330
         TabIndex        =   1
         Top             =   450
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   503
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
      End
      Begin MSComctlLib.ListView LV 
         Height          =   3210
         Left            =   90
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   885
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   5662
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15199215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Unit"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Rate"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amt"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSForms.Label Label2 
         Height          =   240
         Index           =   1
         Left            =   345
         TabIndex        =   17
         Top             =   210
         Width           =   5670
         BackColor       =   11517387
         Caption         =   "  Item                                       Unit               Rate         Qty           Amt"
         Size            =   "10001;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtUnit 
         Height          =   285
         Left            =   2475
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
         VariousPropertyBits=   679495711
         Size            =   "1508;503"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtRate 
         Height          =   285
         Left            =   3375
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
         VariousPropertyBits=   679495711
         Size            =   "1508;503"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtQty 
         Height          =   285
         Left            =   4275
         TabIndex        =   2
         Top             =   480
         Width           =   855
         VariousPropertyBits=   679495707
         Size            =   "1508;503"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAmt 
         Height          =   285
         Left            =   5175
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   855
         VariousPropertyBits=   679495711
         Size            =   "1508;503"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdAdd 
         Height          =   600
         Left            =   6075
         TabIndex        =   3
         Top             =   180
         Width           =   585
         ForeColor       =   0
         Caption         =   "Add"
         Size            =   "1032;1058"
         Picture         =   "NewContEntry.frx":0564
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdSave 
         Height          =   360
         Left            =   4080
         TabIndex        =   9
         Top             =   4170
         Width           =   1425
         ForeColor       =   0
         Caption         =   "  Save       "
         PicturePosition =   327683
         Size            =   "2514;635"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   360
         Left            =   5565
         TabIndex        =   5
         Top             =   4170
         Width           =   1425
         ForeColor       =   0
         Caption         =   " Cancel     "
         PicturePosition =   327683
         Size            =   "2514;635"
         Picture         =   "NewContEntry.frx":0676
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   0
      Left            =   390
      TabIndex        =   8
      Top             =   690
      Width           =   6420
      BackColor       =   11517387
      Caption         =   "  Department                                Contractor                               Date"
      Size            =   "11324;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "NewContEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Editing As Boolean
Public Sub enterLateHours()
    Load LateHours
    LateHours.Tag = LV.SelectedItem
    LateHours.Show
End Sub

Private Sub cmbShift_Change()
    Call FillList
End Sub

Private Sub cmdCheckAll_Click()
    With LV
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = True
        Next i
    End With
End Sub

Private Sub cmdClearAll_Click()
    With LV
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = False
        Next i
    End With
End Sub



Private Sub cmbCont_matched()
    
    cmbCont.Tag = cmbCont.Text
    cmbItem.ClearVals
    Call cmbItem.AddVals(con, "ItemName", "VAssignedItems", "ItemID", " Where EmpID='" & cmbCont.ID & "'")
    
End Sub

Private Sub cmbCont_Validate(cancel As Boolean)
    If cmbCont.Text <> cmbCont.Tag Then
        'cmbCont.Tag = cmbCont.Text
        LV.ListItems.Clear
    End If
End Sub

Private Sub cmbItem_matched()
    On Error GoTo ERR
    Dim rs As New ADODB.Recordset
    With rs
        .Open "select * from VAssignedItems Where ItemID='" & cmbItem.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        txtUnit = ![Unit] & ""
        txtRate = ![Rate] & ""
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
ERR:
    MsgBox ERR.Description
End Sub

Private Sub cmbItem_Validate(cancel As Boolean)
    txtQty.SetFocus
End Sub

Private Sub cmdAdd_Click()
    On Error Resume Next
    If Val(txtQty) = 0 Then
        MsgBox "Quantity Can't be Null.", vbInformation
        Exit Sub
    End If
    Dim i As Integer
    i = LV.ListItems.Count + 1
    LV.ListItems.Add , cmbItem.ID & "'", cmbItem.ID
    If ERR.Number = 35602 Then
        MsgBox "Invalid Entry", vbInformation
        Exit Sub
    End If
    On Error GoTo ERR
    LV.ListItems(i).SubItems(1) = cmbItem.Text
    LV.ListItems(i).SubItems(2) = txtUnit
    LV.ListItems(i).SubItems(3) = txtQty
    LV.ListItems(i).SubItems(4) = txtRate
    LV.ListItems(i).SubItems(5) = txtAmt
    txtUnit = ""
    txtQty = ""
    txtRate = ""
    txtAmt = ""
    'cmbItem.RemoveItem cmbItem.Index
    Exit Sub
ERR:
    MsgBox ERR.Description & vbNewLine & ERR.Number
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
On Error GoTo ERR
Call ShowWait

With cr1
    .ReportFileName = RptPRPath & "\CustVendLedger.rpt"
    .SubreportToChange = ""
    '.DataFiles(0) = App.Path & "\OnlyStock.mdb"
    .Connect = ConnectStr
    
    .SelectionFormula = "{VOrdItems.VendID}='" & cmbCustID & _
      "' AND ({VOrdItems.Dated}=date(" & _
      Format(DtFrom.Tag, "yyyy,MM,dd") & _
      ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
      
    .Formulas(0) = "CompanyName='" & getCompanyName & "'"
    .Formulas(1) = "FromTo='From " & Format(DtFrom.Tag, "dd-MMM-yyyy") & " To " & Format(DtTo.Tag, "dd-MMM-yyyy") & "'"
    .Formulas(2) = "Ledger='Ledger of " & cmbCustName & "'"
    '.Formulas(2) = "SubReport=" & LVItems.Visible
    'if Cheques Are Present then Show Cheque Detail Also
    'If LVItems.Visible Then
    '    .SubreportToChange = "ChqLedger"
    '    .DataFiles(0) = DBName
    '    .SelectionFormula = "{VChqLedger.AccNo}='" & cmbCustID & "' AND ({VChqLedger.CDate}=date(" & Format(DtFrom.Tag, "yyyy,MM,dd") & ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
    'End If
    .Action = 1
    .PageZoomNext
End With

Call HideWait
Exit Sub
ERR:
    MsgBox ERR.Description
End Sub
Public Sub HideWait()

    Unload frmWait
    Screen.MousePointer = vbDefault
End Sub
Public Sub cmdRefresh_Click()
    'Call RefreshLedger
    
    'Call Form_Resize
End Sub

Private Sub cmdSave_Click()

    On Error GoTo ERR
    
    con.BeginTrans
    With LV
        For i = 1 To .ListItems.Count
            
            Sql = "Insert into contLedger(EmpID,DT,ItemID,Rate,Qty) Values('" & _
             cmbCont.ID & "','" & DTPicker1 & "','" & LV.ListItems(i).Text & _
             "'," & LV.ListItems(i).SubItems(3) & "," & LV.ListItems(i).SubItems(4) & ")"
            
            con.Execute Sql
        Next i
    End With
    con.CommitTrans
    Unload Me
    Exit Sub
ERR:
    MsgBox ERR.Description
    con.RollbackTrans
    
End Sub



Private Sub DTPicker1_Change()
    Call FillList
End Sub

Private Sub Form_Load()

    DTPicker1 = ServerDate
    
    cmbDept.ListHeight = 1600
    cmbCont.ListHeight = 1250
    cmbItem.ListHeight = 1600
    'cmbShift.AddItem "<Both Shifts>"
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "ContDepartments", "DeptID"
    cmbDept.Text = "<All Departments>"
    'Set LV.SmallIcons = ImageList1
    
End Sub
Private Sub cmbdept_matched()
    Dim Where As String
    cmbCont.ClearVals
    If cmbDept.ID = "0" Then
        Where = ""
    Else
        Where = " Where DeptID='" & cmbDept.ID & "'"
    End If
    Call cmbCont.AddVals(con, "Name", "Contractors", "EmpID", Where)
    Where = ""
End Sub





Private Sub FillList()

    If cmbDept.MatchFound = False Then Exit Sub
    Dim rsList As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim ITM As ListItem
    
    Dim CanEdit As Boolean
    
    With rsList
        .Open "Select DT from AttendanceSheet Where DT>'" & Format(DTPicker1, "dd-MMM-yyyy") & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            'Exit Sub Coz Can't Edit...
            Exit Sub
        End If
        Set rsList = Nothing
    End With
    
        With cmd
            .CommandText = "Select * from VTakeAtt "
            .ActiveConnection = con
            .Parameters.Refresh
            .Parameters(0).Type = adDate
            .Parameters(0).value = Format(DTPicker1, "dd/MMM/yyyy")
            .Parameters(1).Type = adVarChar
            .Parameters(1).value = cmbShift
            rsList.CursorLocation = adUseClient
            rsList.CursorType = adOpenForwardOnly
            rsList.LockType = adLockReadOnly
            Set rsList = .Execute
        End With
    With rsList
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.Add(, ![EmpId], ![EmpId] & "", , ICO)
            ITM.SubItems(1) = ![Name]
            ITM.SubItems(2) = ![FName]
            ITM.SubItems(3) = cmbShift
            If Not IsNull(![LateHours]) Then ITM.SubItems(4) = ![LateHours]
            If IsNull(![Attendance]) Then
                ITM.Checked = False
                
            Else
                If ![Attendance] Then
                    ITM.Checked = True
                Else
                    ITM.Checked = False
                End If
                ITM.Tag = "Edit"
            End If
            'End If
            .MoveNext
        Next
        .Close
    End With
    
    Set rsLedger = Nothing
    
    'If cmbDept.ID = "0" Then
    '    If CanEdit = False Then
    '        Cond = " Where  empid not in" & _
    '              "(Select EmpID from AttendanceSheet Where " & _
    '            "format(DT,'dd-MMM-yyyy')='" & _
    '            Format(DTPicker1, "dd-MMM-yyyy") & "' and " & Shift & ")"
    '    End If
    'Else
    '    Cond = " Where DeptID='" & cmbDept.ID & "'"
    '    If CanEdit = False Then
    '        Cond = Cond & " and empid not in" & _
    '        "(Select EmpID from AttendanceSheet Where " & _
    '        "format(DT,'dd-MMM-yyyy')='" & _
    '        Format(DTPicker1, "dd-MMM-yyyy") & "' and " & Shift & ")"
    '    End If
    'End If
    'Sql = Sql & Cond
           
    
    
End Sub
Private Sub AddEmps(ITM As ListItem, rs As ADODB.Recordset, Shift As String)
    With rs
        Set ITM = LV.ListItems.Add(, ![EmpId] & Shift, ![EmpId] & "", , ICO)
        ITM.SubItems(1) = ![Name]
        ITM.SubItems(2) = ![FName]
        ITM.SubItems(3) = Shift
        ITM.SubItems(4) = 0
    End With
    'Itm.SubItems(4) = Format(![DT], "dddd dd MMMM ,yyyy")
        
End Sub




Private Sub LV_DblClick()
    If LV.ListItems.Count = 0 Then Exit Sub
    With txtQtyEdit
        .Move LV.Left + LV.ColumnHeaders(4).Left + 35, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(4).Width
        .Text = LV.SelectedItem.SubItems(3)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If LV.ListItems.Count = 0 Then Exit Sub
    If KeyCode = 46 Then 'If Del Key Code...
        If MsgBox("Delete This Entry?", vbQuestion + vbYesNo) = vbYes Then
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
End Sub


Private Sub txtQtyEdit_KeyPress(KeyAscii As Integer)

On Error GoTo ERR
    If KeyAscii = 27 Then
        txtQtyEdit = ""
        txtQtyEdit.Visible = False
    ElseIf KeyAscii = 13 Then
        LV.SelectedItem.SubItems(3) = Val(txtQtyEdit)
        LV.SelectedItem.SubItems(5) = Val(LV.SelectedItem.SubItems(4)) * Val(txtQtyEdit)
        'LV.SelectedItem.SubItems(5) = Val(txtqtyedit) * Val(LV.SelectedItem.SubItems(2))
        txtQtyEdit.Visible = False
        'lblTotal = GetTotal
    Else
        OnlyNumber (KeyAscii)
    End If
    Exit Sub
ERR:
    MsgBox ERR.Description
    
End Sub

Private Sub txtQty_Change()
    txtAmt = Val(txtQty) * Val(txtRate)
End Sub

Private Sub txtQty_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii.value = OnlyNumber(KeyAscii.value)
End Sub
