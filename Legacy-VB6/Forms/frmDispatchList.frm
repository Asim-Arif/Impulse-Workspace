VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmDispatchList 
   ClientHeight    =   7695
   ClientLeft      =   75
   ClientTop       =   270
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
      TabIndex        =   1
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dispatch List"
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
         TabIndex        =   2
         Top             =   165
         Width           =   10425
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Dispatch List"
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
         TabIndex        =   3
         Top             =   180
         Width           =   10410
      End
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   7050
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   12720
      Begin VB.CheckBox chkNotFinal 
         Caption         =   "Not Finalized"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9540
         TabIndex        =   12
         Top             =   450
         Width           =   1545
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   60
         TabIndex        =   11
         Top             =   405
         Width           =   3555
         _ExtentX        =   6271
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
         ItemData        =   "frmDispatchList.frx":0000
         Left            =   3630
         List            =   "frmDispatchList.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   405
         Width           =   2745
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5790
         Left            =   60
         TabIndex        =   4
         Top             =   735
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   10213
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
            Text            =   "Dispatch List No."
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Total Cartons"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Finalized"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Customer"
            Object.Width           =   2822
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
         Left            =   7920
         TabIndex        =   7
         Top             =   405
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
         Format          =   157155331
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6375
         TabIndex        =   8
         Top             =   405
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
         Format          =   157155331
         CurrentDate     =   37055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   $"frmDispatchList.frx":0097
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
         TabIndex        =   10
         Top             =   150
         Width           =   12585
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   345
         Left            =   11130
         TabIndex        =   9
         Top             =   345
         Width           =   1515
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2672;609"
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
         TabIndex        =   5
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
   Begin VB.Menu mnuPop 
      Caption         =   "Pop Up"
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuDash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewDispatchList 
         Caption         =   "Add New Dispatch List"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDispatchList 
         Caption         =   "Edit Dispatch List"
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintDispatchList 
         Caption         =   "Print Dispatch List"
      End
      Begin VB.Menu mnuPrintDespatchFinal 
         Caption         =   "Print Dispatch Final"
      End
      Begin VB.Menu mnuPrintDispatchListWithValue 
         Caption         =   "Print Dispatch List (With Value)"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFinalyzeDP 
         Caption         =   "Finalize Dispatch List"
      End
   End
End
Attribute VB_Name = "frmDispatchList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String

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
            DtFrom = DateAdd("d", -7, GetServerDate)
        Case 2
            DtFrom = DateAdd("d", -15, GetServerDate)
        Case 3
            DtFrom = DateAdd("d", -30, GetServerDate)
        Case 4
            DtFrom = DateAdd("d", -60, GetServerDate)
        Case 5
            DtFrom = DateAdd("d", -90, GetServerDate)
        Case 6
            DtFrom = DateAdd("yyyy", -1, GetServerDate)
        Case 7
            DtFrom = DateAdd("yyyy", -5, GetServerDate)
        Case 8
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub


Private Sub cmdClose_Click()
   Unload Me
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub

Private Sub Form_Load()
    
'    cmbRange.ListIndex = 0
    mnuPop.Visible = False
    Call RefreshLV
    cmbRange.ListIndex = 0
    
'    cmbMaterial.ListHeight = 4000
'    cmbMaterial.AddItem "<All Materials>", "0"
'    cmbMaterial.AddVals con, "'{' + RMID1 + '} ' + RMName", "RM", "RMID1"
'    cmbMaterial.ID = "0"

    cmbCust.ListHeight = 4000
    cmbCust.AddItem "<All Customers>", "0"
    cmbCust.AddVals con, "DISTINCT CustCode", "ForeignCustomers", "CustCode", " WHERE CustCode<>'Stock'"
    cmbCust.ID = "0"
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
    
'    DTRange = " WHERE DT BETWEEN '" & DtFrom & "' and '" & DtTo & "'"
'
'    If cmbMaterial.ID <> "0" Then
'        DTRange = DTRange & " AND IssNo IN(SELECT IssNo FROM RawMaterialIssuanceDetail WHERE RMID1='" & cmbMaterial.ID & "')"
'    End If
    DTRange = " WHERE DT BETWEEN '" & DtFrom & "' AND '" & DtTo & "'"
    If cmbCust.ID <> "0" Then
        DTRange = DTRange & " AND CustCode='" & cmbCust.ID & "'"
    End If
        
    If chkNotFinal.Value = vbChecked Then
        DTRange = DTRange & " AND Finalyzed=0"
    End If
    Dim strFinalizedDT As String
    With rs
        SQL = "SELECT * FROM VDispatchList " & DTRange & " ORDER BY DT"
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        'LV.Sorted = False
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, !EntryID & "'", .AbsolutePosition)
            
            ITM.ListSubItems.add , , !DispatchListNo & ""
            ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
            ITM.ListSubItems.add(, , Val(!TotalCartons & "")).Tag = !NewFormat
            
            If !Finalyzed Then
                strFinalizedDT = Format(!FinalyzedDT, "dd-MMM-yyyy")
            Else
                strFinalizedDT = "-"
            End If
            
            ITM.ListSubItems.add(, , strFinalizedDT).Tag = !Finalyzed
            ITM.ListSubItems.add , , !CustCode & ""
            ITM.ListSubItems.add , , Val(!TotalQty & "")
            
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
    
    If LV.ListItems.count = 0 Then
        mnuEditDispatchList.Enabled = False
    ElseIf CBool(LV.SelectedItem.ListSubItems(3).Tag) = False Then  'If Older Format, Disable Editing.
        mnuEditDispatchList.Enabled = False
    Else
        mnuEditDispatchList.Enabled = True
    End If
    Me.PopUpMenu mnuPop
        
End Sub


Private Sub mnuEditDispatchList_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
'    If strCompany = "Dr-Frgz" Then
'        MsgBox "Not Available.", vbInformation
'        Exit Sub
'    End If
        
    If CBool(LV.SelectedItem.ListSubItems(4).Tag) Then
        If getDBPassword("EditFinalizedDispatchList") = False Then Exit Sub
    End If
    
    Load frmCreateDispatchList_Adv
    frmCreateDispatchList_Adv.EditDispatchList (Val(LV.SelectedItem.key))
    
End Sub

Private Sub mnuFinalyzeDP_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If CBool(LV.SelectedItem.ListSubItems(4).Tag) Then
        MsgBox "Already Finalized.", vbInformation
        Exit Sub
    End If
       
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "DispatchList INNER JOIN DispatchListDetails_Adv ON " & _
     "DispatchList.EntryID=DispatchListDetails_Adv.RefID INNER JOIN DispatchListDetail_Inners ON " & _
     "DispatchListDetails_Adv.EntryID=DispatchListDetail_Inners.RefID", " WHERE DispatchList.EntryID=" & _
     Val(LV.SelectedItem.key) & " AND  ISNULL(CartonNo,0)=0")
    
    If lCount > 0 Then
        MsgBox "Can't Finalize, Carton No. not Assigned.", vbInformation
        Exit Sub
    End If
    
    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
        con.Execute "UPDATE DispatchList SET Finalyzed=1,FinalyzedDT=getDate(),FinalyzedUserName='" & _
         CurrentUserName & "',FinalyzedMachineName='" & strComputerName & _
         "' WHERE EntryID=" & Val(LV.SelectedItem.key)
    End If
    
    Call RefreshLV
    
End Sub

Private Sub mnuNewDispatchList_Click()
'    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" OR strCompany = "Kami" Then
'        frmCreateDispatchList_Adv.Show
'    Else
'        frmCreateDispatchList.Show
'    End If
    frmCreateDispatchList_Adv.Show
'    frmCreateDispatchList.Show
End Sub

Private Sub mnuPrintDespatchFinal_Click()
    Call PrintDispatchList(False, 1)
End Sub

Private Sub mnuPrintDispatchList_Click()
    Call PrintDispatchList
End Sub

Private Sub PrintDispatchList(Optional bWithValue As Boolean = False, Optional iReportType As Integer = 0)
    
    If LV.ListItems.count = 0 Then Exit Sub
        
    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    
    Dim strSelection As String
    If bWithValue Then
        Set rpt = rptApp.OpenReport(RptPath & "\DispatchListInnerwiseWithValue.rpt")
        strSelection = "{DispatchList.DispatchListEntryID}=" & Val(LV.SelectedItem.key)
    Else
        If CBool(LV.SelectedItem.ListSubItems(3).Tag) = False Then   'If Older Format, Disable Editing.
            Set rpt = rptApp.OpenReport(RptPath & "\DispatchList.rpt")
        Else
            If iReportType = 0 Then
                Set rpt = rptApp.OpenReport(RptPath & "\DispatchListInnerwise.rpt")
            Else
                Set rpt = rptApp.OpenReport(RptPath & "\DispatchListFinal.rpt")
            End If
        End If
        strSelection = "{DispatchList.EntryID}=" & Val(LV.SelectedItem.key)
    End If
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintDispatchListWithValue_Click()

    If getDBPassword("ViewPValueReport") = False Then
        Exit Sub
    End If
    
    Call PrintDispatchList(True)
    
End Sub

Private Sub mnuRefresh_Click()
    Call RefreshLV
End Sub
