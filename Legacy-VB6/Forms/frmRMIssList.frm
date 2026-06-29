VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmRMIssList 
   ClientHeight    =   7695
   ClientLeft      =   75
   ClientTop       =   -30
   ClientWidth     =   12780
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleMode       =   0  'User
   ScaleWidth      =   11447.22
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   1080
      TabIndex        =   6
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Issued Entries"
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
         Index           =   8
         Left            =   75
         TabIndex        =   7
         Top             =   165
         Width           =   10425
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Issued Entries"
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
         Height          =   465
         Index           =   9
         Left            =   90
         TabIndex        =   8
         Top             =   180
         Width           =   10410
      End
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   7050
      Left            =   15
      TabIndex        =   5
      Top             =   615
      Width           =   12720
      Begin ComboList.Usercontrol1 cmbDept 
         Height          =   285
         Left            =   75
         TabIndex        =   15
         Top             =   990
         Width           =   3705
         _ExtentX        =   6535
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
         Appearance      =   0
      End
      Begin ComboList.Usercontrol1 cmbRMGroups 
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Top             =   405
         Width           =   2655
         _ExtentX        =   4683
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
         Appearance      =   0
      End
      Begin ComboList.Usercontrol1 cmbMaterial 
         Height          =   285
         Left            =   2730
         TabIndex        =   1
         Top             =   405
         Width           =   4350
         _ExtentX        =   7673
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
         Appearance      =   0
      End
      Begin ComboList.Usercontrol1 cmbEmp 
         Height          =   285
         Left            =   3780
         TabIndex        =   13
         Top             =   990
         Width           =   8880
         _ExtentX        =   15663
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
         Appearance      =   0
      End
      Begin VB.ComboBox cmbRange 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmRMIssList.frx":0000
         Left            =   7080
         List            =   "frmRMIssList.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   390
         Width           =   1590
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5250
         Left            =   75
         TabIndex        =   9
         Top             =   1290
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   9260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Iss No."
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Maker"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Employee"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Ist RM"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Total Qty"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   10215
         TabIndex        =   4
         Top             =   390
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   129630211
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   8670
         TabIndex        =   3
         Top             =   390
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   129630211
         CurrentDate     =   37055
      End
      Begin MSForms.CommandButton cmdPrintList 
         Height          =   360
         Left            =   9360
         TabIndex        =   16
         Top             =   6600
         Width           =   1620
         ForeColor       =   0
         Caption         =   "Print List"
         PicturePosition =   327683
         Size            =   "2857;635"
         Accelerator     =   76
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   " Department                                                           Maker / Employee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   1
         Left            =   75
         TabIndex        =   14
         Top             =   735
         Width           =   12585
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   $"frmRMIssList.frx":0066
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   12
         Top             =   150
         Width           =   11670
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   540
         Left            =   11745
         TabIndex        =   11
         Top             =   150
         Width           =   915
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1614;952"
         Accelerator     =   82
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
         Left            =   11040
         TabIndex        =   10
         Top             =   6600
         Width           =   1620
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2857;635"
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
End
Attribute VB_Name = "frmRMIssList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim bSample As Boolean
Private Sub cmbdept_matched()
     
    Dim strCondition As String
    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.ID = "0" Then
        strCondition = ""
        cmbEmp.AddVals con, "VendID1 + ' ' + VenderName", "Makers", "CAST(VendID AS VARCHAR(50))+''''"
    Else
        strCondition = " WHERE DeptID='" & cmbDept.ID & "'"
    End If
    cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name", "Employees", "EmpID", strCondition & " ORDER BY EmpID"

    
    cmbEmp.ID = "0"
    
End Sub

Private Sub cmbMaterial_matched()
    Call RefreshLV
End Sub

Private Sub cmbRange_Change()
    Call cmbRange_Click
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = GetServerDate
        Case 1
            DtFrom = DateAdd("d", -15, GetServerDate)
        Case 2
            DtFrom = DateAdd("d", -30, GetServerDate)
        Case 3
            DtFrom = DateAdd("d", -60, GetServerDate)
        Case 4
            DtFrom = DateAdd("d", -90, GetServerDate)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub


Private Sub cmbRMGroups_matched()

    Dim strCondition As String
    strCondition = "WHERE Inactive=0"
    If cmbRMGroups.ID <> "0" Then
        strCondition = strCondition & " AND GroupID=" & cmbRMGroups.ID
    End If
    cmbMaterial.ClearVals
    cmbMaterial.AddItem "<All Materials>", "0"
    cmbMaterial.AddVals con, "'{' + RMID1 + '} ' + RMName", "RM", "RMID1", strCondition
    cmbMaterial.ID = "0"
    
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdPrintList_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RawMaterialIssuanceList.rpt")
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy") & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, SQL, False
    
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub

Private Sub Form_Load()
    
'    cmbRange.ListIndex = 0
    bSample = False
    Call RefreshLV
    cmbRange.ListIndex = 0
    
    cmbRMGroups.ListHeight = 4000
    cmbRMGroups.AddItem "<All Groups>", "0"
    cmbRMGroups.AddVals con, "Description", "RMGroups", "ID"
    cmbRMGroups.ID = "0"
    
    cmbMaterial.ListHeight = 4000
    
    cmbEmp.ListHeight = 4000
    
    cmbDept.ListHeight = 4000
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.ID = "0"
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim DTRange As String
    
    Screen.MousePointer = vbHourglass
    
'    If cmbRange.ListIndex = 0 Or cmbRange.ListIndex = -1 Then
'        DTRange = " DT Between '" & DtFrom & "' and '" & DtTo & "'"
'    ElseIf cmbRange.ListIndex = 1 Then
'        DTRange = " DT Between '" & DateAdd("d", -30, GetServerDate) & "' and '" & GetServerDate & "'"
'    ElseIf cmbRange.ListIndex = 2 Then
'        DTRange = " DT Between '" & DateAdd("d", -60, GetServerDate) & "' and '" & GetServerDate & "'"
'    ElseIf cmbRange.ListIndex = 3 Then
'        DTRange = " DT Between '" & DateAdd("d", -90, GetServerDate) & "' and '" & GetServerDate & "'"
'    ElseIf cmbRange.ListIndex = 4 Then
'        DTRange = " DT Between '" & DtFrom & "' and '" & DtTo & "'"
'    End If
    
    DTRange = " WHERE DT BETWEEN '" & DtFrom & "' and '" & DtTo & "'"
    
    
    If cmbMaterial.ID <> "0" Then
        DTRange = DTRange & " AND IssNo IN(SELECT IssNo FROM RawMaterialIssuanceDetail WHERE RMID1='" & cmbMaterial.ID & "')"
    ElseIf cmbRMGroups.ID <> "0" Then
        DTRange = DTRange & " AND IssNo IN(SELECT IssNo FROM RawMaterialIssuanceDetail INNER JOIN RM ON RM.RMID1=RawMaterialIssuanceDetail.RMID1 WHERE GroupID=" & cmbRMGroups.ID & ")"
    End If
    
    
    If cmbEmp.ID <> "0" Then
        If Right(cmbEmp.ID, 1) = "'" Then
            DTRange = DTRange & " AND VendID=" & Val(cmbEmp.ID)
        Else
            DTRange = DTRange & " AND IssdToEmpIDOnly='" & cmbEmp.ID & "'"
        End If
    Else
        If cmbDept.ID <> "0" Then
            DTRange = DTRange & " AND DeptID='" & cmbDept.ID & "'"
        End If
    End If
    
    With rs
        SQL = "SELECT * FROM VRawMaterialIssuanceList " & DTRange & " ORDER BY DT,IssNo"
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        'LV.Sorted = False
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, ![IssNo] & "'", .AbsolutePosition)
            
            ITM.ListSubItems.add , , !IssNo
            ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
            ITM.ListSubItems.add , , !VenderName & ""
            ITM.ListSubItems.add , , !IssdToEmpID & ""
            ITM.ListSubItems.add , , !RM & ""
            ITM.ListSubItems.add , , Val(!Total_Issued_Qty & "")
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
 
    Dim strRet As String
'    If LV.ListItems.count = 0 Then
'        strRet = ShowPopUpMenu1("&Issue New Material")
'    Else
    'If cmbMaterial.MatchFound And cmbMaterial.ID <> "0" Then
        'strRet = ShowPopUpMenu1("&Print Issuance Report")
    'Else
    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        strRet = ShowPopUpMenu1("&Print This Entry", "&Print Issuance Report", "Delete Selected Issuance")
    Else
        strRet = ShowPopUpMenu1("&Print This Entry", "&Print Issuance Report")
    End If
    'End If
    
'    End If
    'strRet = ShowPopUpMenu1("&Issue New Material")
    Dim rpt As CRAXDDRT.Report
    Dim strSelection As String
    Dim f As Form
    Select Case strRet
    
            
        Case "&Print This Entry"
        
            Set f = New frmPrevRpt
            
            Set rpt = rptApp.OpenReport(RptPath & "\IssuanceNote.rpt")
            f.ShowReport "{IssItemsSimple.IssNo}='" & LV.SelectedItem.SubItems(1) & "' ", rpt
            
        Case "&Print Issuance Report"
            Set f = New frmPrevRpt
            If strCompany = "Dr-Frgz" Then
                Set rpt = rptApp.OpenReport(RptPath & "\RMIssuanceList_Grouped.rpt")
            Else
                Set rpt = rptApp.OpenReport(RptPath & "\RMIssuanceList.rpt")
            End If
            rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo, "dd-MMM-yyyy") & "'"
            strSelection = "{IssItemsSimple.DT}=#" & DtFrom & "# TO #" & DtTo & "#"
            If cmbMaterial.ID <> "0" Then
                strSelection = strSelection & " AND {IssItemsSimpleDetail.RMID1}='" & cmbMaterial.ID & "'"
            End If
            If cmbRMGroups.ID <> "0" Then
                strSelection = strSelection & " AND {RMGroups.ID}=" & cmbRMGroups.ID
            End If
            If cmbEmp.ID <> "0" Then
                If Right(cmbEmp.ID, 1) = "'" Then
                    strSelection = strSelection & " AND {VendIssued.VendID}=" & Val(cmbEmp.ID)
                Else
                    strSelection = strSelection & " AND {IssItemsSimple.IssdToEmpID}='" & cmbEmp.ID & "'"
                End If
            Else
                If cmbDept.ID <> "0" Then
                    strSelection = strSelection & " AND {VEmp.DeptID}='" & cmbDept.ID & "'"
                End If
            End If
            f.ShowReport strSelection, rpt
        Case "&Return Stock"
        
'            Dim frm As New frmReturnIssdItems
'            frm.ShowDetail (LV.SelectedItem.SubItems(1))
        
        Case "Print &List"
        
'            Set rpt = rptApp.OpenReport(rptProductionPath & "\rptMaterialIssList.rpt")
'            rpt.FormulaFields(1).Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo, "dd-MMM-yyyy") & "'"
            
'            f.ShowReport "{IssItemsSimple.DT}=#" & Format(DtFrom.value, "dd-MMM-yyyy") & "# To #" & Format(DtTo.value, "dd-MMM-yyyy") & "#", rpt
        Case "Delete Selected Issuance"
            If DeleteIssuance(LV.SelectedItem.SubItems(1)) Then
                Call RefreshLV
            End If
    End Select
        
End Sub


Public Sub ShowMe(Optional p_bSample As Boolean = False)
    bSample = p_bSample
    Me.Show
End Sub
