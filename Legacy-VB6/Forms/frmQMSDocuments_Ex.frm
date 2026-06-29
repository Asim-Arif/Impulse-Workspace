VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmQMSDocuments_Ex 
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8955
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8955
      Left            =   60
      TabIndex        =   3
      Top             =   -45
      Width           =   11865
      Begin VB.CommandButton cmdPrintDistListNewlyUpdated 
         Caption         =   "Dist. List (Newly Updated)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4830
         TabIndex        =   16
         Top             =   8475
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.CommandButton cmdPrintDistributionListSelected 
         Caption         =   "Dist. List (Selected)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2580
         TabIndex        =   15
         Top             =   8475
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.CommandButton cmdDeselectAll 
         Caption         =   "Deselect All"
         Height          =   405
         Left            =   1335
         TabIndex        =   14
         Top             =   8475
         Width           =   1215
      End
      Begin VB.CommandButton chkSelectAll 
         Caption         =   "Select All"
         Height          =   405
         Left            =   60
         TabIndex        =   13
         Top             =   8475
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   11850
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Document Control"
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
            TabIndex        =   11
            Top             =   165
            Width           =   11715
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Document Control"
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
            TabIndex        =   12
            Top             =   180
            Width           =   11700
         End
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
         ItemData        =   "frmQMSDocuments_Ex.frx":0000
         Left            =   4440
         List            =   "frmQMSDocuments_Ex.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   915
         Visible         =   0   'False
         Width           =   2010
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   8190
         TabIndex        =   1
         Top             =   915
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   122290179
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6465
         TabIndex        =   0
         Top             =   915
         Visible         =   0   'False
         Width           =   1710
         _ExtentX        =   3016
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
         Format          =   122290179
         CurrentDate     =   37055
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   195
         TabIndex        =   6
         Top             =   915
         Visible         =   0   'False
         Width           =   2310
         _ExtentX        =   4075
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
      Begin MSComctlLib.ListView LV 
         Height          =   7740
         Left            =   60
         TabIndex        =   8
         Top             =   705
         Width           =   11730
         _ExtentX        =   20690
         _ExtentY        =   13653
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Revision"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Type"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "IssueDate"
            Object.Width           =   2328
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "PDF"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   0
         Top             =   585
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Please Select Picture"
         Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10455
         TabIndex        =   9
         Top             =   8505
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2550
         TabIndex        =   7
         Top             =   915
         Visible         =   0   'False
         Width           =   1860
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3281;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   60
         TabIndex        =   2
         Top             =   885
         Width           =   1500
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2646;635"
         Accelerator     =   82
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
         BackStyle       =   0  'Transparent
         Caption         =   $"frmQMSDocuments_Ex.frx":0078
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
         Left            =   165
         TabIndex        =   4
         Top             =   690
         Visible         =   0   'False
         Width           =   11130
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQMSDocuments_Ex.frx":0112
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQMSDocuments_Ex.frx":02CE
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add New"
      End
      Begin VB.Menu dd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit This Entry"
      End
      Begin VB.Menu mnuDashNC 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeRequest 
         Caption         =   "Change Request"
      End
      Begin VB.Menu mnuViewHistory 
         Caption         =   "View History"
      End
      Begin VB.Menu mnuDashPrint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Entry"
      End
   End
End
Attribute VB_Name = "frmQMSDocuments_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iColNo As Integer
Dim strMaster_DocType As String

Private Sub chkSelectAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    
    Sql = "SELECT DISTINCT Country FROM ForeignCustomers "
    
    If cmbCust.ID <> "0" Then
        Sql = Sql & " WHERE CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        
        cmbCountry.Clear
        cmbCountry.AddItem "<All Countries>"
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = Date
        Case 1
            DtFrom = DateAdd("d", -15, Date)
        Case 2
            DtFrom = DateAdd("d", -30, Date)
        Case 3
            DtFrom = DateAdd("d", -60, Date)
        Case 4
            DtFrom = DateAdd("d", -90, Date)
        Case 5
            DtFrom = DateAdd("d", -180, Date)
        Case 6
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDeselectAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdPrintDistListNewlyUpdated_Click()
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\QMSDocument_DistributionList.rpt")
    '--------------------------------------------------------------------------
    Dim myDTFrom As Date, myDTTo As Date
    myDTFrom = DateAdd("d", -7, Date)
    myDTTo = Date
    
    Dim f As New frmPrevRpt
    f.ShowReport "{QMS_Documents_Ex.LastApprovedEntryDT}=#" & myDTFrom & "# TO #" & myDTTo & "#", rpt
    
End Sub

Private Sub cmdPrintDistributionListSelected_Click()

    Dim strEntryIDs As String
    Dim i As Integer
    strEntryIDs = ""
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            strEntryIDs = strEntryIDs & Val(LV.ListItems(i).Tag) & ","
        End If
    Next
    If strEntryIDs = "" Then
        MsgBox "No Documents Selected.", vbInformation
        Exit Sub
    End If
    
    strEntryIDs = Left(strEntryIDs, Len(strEntryIDs) - 1)
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\QMSDocument_DistributionList.rpt")
    '--------------------------------------------------------------------------
    Dim f As New frmPrevRpt
    f.ShowReport "{QMS_Documents_Ex.EntryID} IN[" & strEntryIDs & "]", rpt
    
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub

Private Sub Form_Load()
    
    mnuPop.Visible = False
    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "Distinct CustCode", "ForeignCustomers", "CustCode")
    cmbCust.ID = "0"
    DtFrom = Date
    DtTo = Date
    cmbRange.ListIndex = 0
    
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim strCondition As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
'    strCondition = " WHERE DT BETWEEN '" & DtFrom & "' AND '" & DtTo & "'"
'
'    If cmbCust.ID <> "0" Then
'        strCondition = strCondition & " AND CustCode='" & cmbCust.ID & "' "
'    End If
'
'    If cmbCountry.ListIndex <> 0 Then
'        strCondition = strCondition & " and Country='" & cmbCountry & "'"
'    End If
    
'    Dim i As Integer, strExtra1  As String, strExtra2 As String
'    strExtra1 = " AND Type NOT IN("
'    strExtra2 = ""
'    For i = 0 To chkType.UBound
'        If chkType(i).Value = vbUnchecked Then
'            strExtra2 = strExtra2 & i & ","
'        End If
'    Next
'    If strExtra2 <> "" Then
'        strExtra2 = Left(strExtra2, Len(strExtra2) - 1)
'        strExtra1 = strExtra1 & strExtra2 & ")"
'        strCondition = strCondition & strExtra1
'    End If
    
    strCondition = " WHERE Master_DocumentType='" & strMaster_DocType & "'"
    With rs
        
        .Open "SELECT * FROM VQMS_Documents_Ex " & strCondition & " ORDER BY IssueDT", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , .AbsolutePosition)
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , !DocumentNumber & ""
            ITM.ListSubItems.add , , !DocumentName & ""
            ITM.ListSubItems.add , , !Revision & ""
            ITM.ListSubItems.add , , !DocumentType & ""
            ITM.ListSubItems.add , , Format(!IssueDT & "", "dd-MMM-yyyy")
            
            If Not IsNull(!AttachmentPDF) Then
                ITM.ListSubItems.add , , "Attached"
            Else
                ITM.ListSubItems.add , , "-"
            End If
             
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


Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If LV.SortKey + 1 = ColumnHeader.Index Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Down"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Up"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If
    'Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub LV_DblClick()

    If LV.SelectedItem.SubItems(6) = "Attached" Then
        Call ShowTechDrawing("QMS_Documents_Ex", "AttachmentPDF", " WHERE EntryID=" & Val(LV.SelectedItem.Tag), LV.SelectedItem.Text, CD1)
    Else
        MsgBox "No PDF attached.", vbInformation
        Exit Sub
    End If
        
    Exit Sub
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        With LV.ColumnHeaders(i)
            If x > Val(.Left) And x < (Val(.Left) + Val(.Width)) Then
                iColNo = i
                Exit For
            End If
        End With
    Next
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub mnuAdd_Click()
'    If UserHasAccess("AddEditCustomerComplaint") = False Then
'        MsgBox "Sorry, No Access.", vbInformation
'        Exit Sub
'    End If
    Load frmQMSDocument_Ex_AE
    frmQMSDocument_Ex_AE.AddNew strMaster_DocType
End Sub

Private Sub mnuAddAudit_Click()
    Load frmAuditAE
    frmAuditAE.AddNew
End Sub

Private Sub mnuAddCustomerFeedback_Click()

    If UserHasAccess("AddEditCustomerComplaint") = False Then
        MsgBox "Sorry, No Access.", vbInformation
        Exit Sub
    End If
    Load frmCustomerFeedbackAE
    frmCustomerFeedbackAE.AddNew

End Sub

Private Sub mnuChangeRequest_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim f As New frmQMSDocument_Ex_ChangeRequest
    Load f
    
    Call f.ChangeRequest(Val(LV.SelectedItem.Tag))
    
End Sub

Private Sub mnuedit_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
'    If UserHasAccess("AddEditCustomerComplaint") = False Then
'        MsgBox "Sorry, No Access.", vbInformation
'        Exit Sub
'    End If
     
    Dim f As New frmQMSDocument_Ex_AE
    
    Load f
    Call f.EditItem(Val(LV.SelectedItem.Tag))
    
End Sub

Private Sub mnuGenerateCAPA_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim lEntryID As Long
    lEntryID = Val(LV.SelectedItem.Tag)
    
    Dim strNCNo As String, strCAPANo As String
    strNCNo = GetSingleStringValue("NCNo", "CustomerComplaints", " WHERE EntryID=" & lEntryID)
    
    If strNCNo = "" Then
        MsgBox "CAPA can only be generated against NC.", vbInformation
        Exit Sub
    End If
    
    strCAPANo = GetSingleStringValue("CAPANo", "CustomerComplaints", " WHERE EntryID=" & lEntryID)
'    If strCAPANo <> "" Then
'        MsgBox "CAPA already generated.", vbInformation
'        Exit Sub
'    End If
     
    Dim lNCEntryID As Long
    lNCEntryID = GetSingleLongValue("EntryID", "NonCompliance", " WHERE NCNo='" & strNCNo & "'")
    Dim f As New frmCAPA_AE
    Load f
    
    If strCAPANo <> "" Then
        Dim lCAPAEntryID As Long
        lCAPAEntryID = GetSingleLongValue("EntryID", "CAPA", " WHERE CAPANo='" & strCAPANo & "'")
        Call f.EditItem(lCAPAEntryID)
    Else
        Call f.AddNew(lEntryID, lNCEntryID)
    End If
    
End Sub

Private Sub mnuGenerateNC_Click()
    
    Dim lEntryID As Long
    lEntryID = Val(LV.SelectedItem.Tag)
    
    Dim strNCNo As String
    strNCNo = GetSingleStringValue("NCNo", "CustomerComplaints", " WHERE EntryID=" & lEntryID)
     
    If strNCNo <> "" Then
        MsgBox "NC Already Generated.", vbInformation
        Exit Sub
    End If
    
    Dim f As New frmNC_AE
    Load f
    
    f.AddNew (lEntryID)
    
End Sub

Private Sub mnuPrint_Click()

'    Dim rpt As CRAXDRT.Report
'    Set rpt = rptApp.OpenReport(RptPath & "\CustomerComplaints.rpt")
'    Dim f As New frmPrevRpt
'    f.ShowReport "{CustomerComplaints.EntryID}=" & Val(LV.SelectedItem.Tag), rpt
    
End Sub

Private Sub mnuPrintCapa_Click()

    Dim rpt As CRAXDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CAPA.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{CAPA.CAPANo}='" & LV.SelectedItem.SubItems(7) & "'", rpt

End Sub

Private Sub mnuPrintCapaLogoSheet_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CAPAReport.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt
    
End Sub

Private Sub mnuPrintNC_Click()
    Dim rpt As CRAXDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\NC.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{NonCompliance.NCNo}='" & LV.SelectedItem.SubItems(5) & "'", rpt
End Sub

Private Sub mnuRefresh_Click()
    Call RefreshLV
End Sub

Private Sub mnuViewHistory_Click()
    Dim f As New frmQMSDocument_Ex_History
    Load f
    f.ShowHistory (Val(LV.SelectedItem.Tag))
End Sub

Public Sub ShowMe(p_strMaster_DocType As String)
    strMaster_DocType = p_strMaster_DocType
    label1(8).Caption = p_strMaster_DocType
    label1(9).Caption = p_strMaster_DocType
    
    Call RefreshLV
    Me.Show
End Sub
