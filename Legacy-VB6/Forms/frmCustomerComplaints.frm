VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmCustomerComplaints 
   ClientHeight    =   8790
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
   ScaleHeight     =   8790
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8790
      Left            =   60
      TabIndex        =   3
      Top             =   -45
      Width           =   11865
      Begin VB.TextBox txtComplaintNo 
         Height          =   285
         Left            =   8190
         TabIndex        =   15
         Top             =   675
         Width           =   1560
      End
      Begin VB.CheckBox chkSelection 
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   885
         Width           =   330
      End
      Begin VB.CommandButton cmdAddNewComplaint 
         Caption         =   "Add New Complaint"
         Height          =   375
         Left            =   60
         TabIndex        =   13
         Top             =   8295
         Width           =   1710
      End
      Begin VB.CheckBox chkType 
         Caption         =   "Internal / External Audit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   4950
         TabIndex        =   12
         Top             =   7935
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkType 
         Caption         =   "Feedback / Observation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2505
         TabIndex        =   11
         Top             =   7935
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkType 
         Caption         =   "Customer Complaints"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   10
         Top             =   7935
         Value           =   1  'Checked
         Width           =   2325
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
         ItemData        =   "frmCustomerComplaints.frx":0000
         Left            =   4440
         List            =   "frmCustomerComplaints.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   345
         Width           =   2010
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   8190
         TabIndex        =   1
         Top             =   345
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
         Format          =   157548547
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6465
         TabIndex        =   0
         Top             =   345
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
         Format          =   138477571
         CurrentDate     =   37055
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   195
         TabIndex        =   6
         Top             =   345
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
         Height          =   6765
         Left            =   60
         TabIndex        =   8
         Top             =   1110
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   11933
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Country"
            Object.Width           =   1722
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Complaint #"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "NC No."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "PDF"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "CAPA No."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Cr Note #"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Please Select Picture"
         Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
      End
      Begin VB.Label Label2 
         Caption         =   "Complaint No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6810
         TabIndex        =   16
         Top             =   705
         Width           =   1290
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10455
         TabIndex        =   9
         Top             =   8295
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
         Top             =   345
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
         Height          =   645
         Left            =   9810
         TabIndex        =   2
         Top             =   315
         Width           =   1500
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2646;1138"
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
         Caption         =   $"frmCustomerComplaints.frx":0078
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
         Top             =   120
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
            Picture         =   "frmCustomerComplaints.frx":0112
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomerComplaints.frx":02CE
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add New Customer Complaint"
      End
      Begin VB.Menu mnuAddCustomerFeedback 
         Caption         =   "Add Customer &Feedback / Observation"
      End
      Begin VB.Menu mnuAddAudit 
         Caption         =   "Add Internal / External Audit"
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
      Begin VB.Menu mnuGenerateNC 
         Caption         =   "Generate NC"
      End
      Begin VB.Menu mnuGenCreditNoteDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenCreditNote 
         Caption         =   "Generate Credit Note"
      End
      Begin VB.Menu mnuDashCAPA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenerateCAPA 
         Caption         =   "Generate/Edit CAPA"
      End
      Begin VB.Menu mnuDashPrint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Entry"
      End
      Begin VB.Menu mnuPrintNC 
         Caption         =   "Print NC"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintCreditNote 
         Caption         =   "Print Credit Note"
      End
      Begin VB.Menu mnuPrintCapa 
         Caption         =   "Print CAPA"
      End
      Begin VB.Menu mnuPrintList 
         Caption         =   "Print List"
      End
      Begin VB.Menu mnuPrintListWOPrice 
         Caption         =   "Print List (Without Price)"
      End
      Begin VB.Menu mnuDashMore1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintCapaLogoSheet 
         Caption         =   "Print Capa Logo Sheet/Report"
      End
   End
End
Attribute VB_Name = "frmCustomerComplaints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iColNo As Integer

Private Sub chkSelection_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = Abs(chkSelection.Value)
    Next
End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT DISTINCT Country FROM ForeignCustomers "
    
    If cmbCust.ID <> "0" Then
        SQL = SQL & " WHERE CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        
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

Private Sub cmdAddNewComplaint_Click()

    If UserHasAccess("AddEditCustomerComplaint") = False Then
        MsgBox "Sorry, No Access.", vbInformation
        Exit Sub
    End If
    Load frmCustomerComplaintAE
    frmCustomerComplaintAE.AddNew

End Sub

Private Sub cmdClose_Click()
   Unload Me
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
    
    strCondition = " WHERE DT BETWEEN '" & DtFrom & "' AND '" & DtTo & "'"
    
    If cmbCust.ID <> "0" Then
        strCondition = strCondition & " AND CustCode='" & cmbCust.ID & "' "
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        strCondition = strCondition & " and Country='" & cmbCountry & "'"
    End If
    
    Dim i As Integer, strExtra1  As String, strExtra2 As String
    strExtra1 = " AND Type NOT IN("
    strExtra2 = ""
    For i = 0 To chkType.UBound
        If chkType(i).Value = vbUnchecked Then
            strExtra2 = strExtra2 & i & ","
        End If
    Next
    If strExtra2 <> "" Then
        strExtra2 = Left(strExtra2, Len(strExtra2) - 1)
        strExtra1 = strExtra1 & strExtra2 & ")"
        strCondition = strCondition & strExtra1
    End If
        
    If Trim(txtComplaintNo) <> "" Then
        strCondition = " WHERE ComplaintNo LIKE '%" & Trim(txtComplaintNo) & "%'"
    End If
    With rs
        
        .Open "SELECT EntryID,CustCode,Country,ComplaintNo,Type,DT,NCNo,CapaNo,CASE WHEN AttachmentPDF IS NOT NULL THEN 1 ELSE 0 END AS PDFAttached,CreditNoteNo FROM CustomerComplaints " & strCondition & " ORDER BY DT", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , .AbsolutePosition)
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , !CustCode & ""
            ITM.ListSubItems.add , , !Country & ""
            ITM.ListSubItems.add(, , !ComplaintNo & "").Tag = Val(!Type & "")
            ITM.ListSubItems.add , , Format(!DT & "", "dd-MMM-yyyy")
            ITM.ListSubItems.add , , IIf(!NCNo & "" = "", "-", !NCNo & "")
            
            If !PDFAttached = 1 Then
                ITM.ListSubItems.add , , "Attached"
            Else
                ITM.ListSubItems.add , , "-"
            End If
            
            ITM.ListSubItems.add , , IIf(!CapaNo & "" = "", "-", !CapaNo & "")
            ITM.ListSubItems.add , , IIf(!CreditNoteNo & "" = "", "-", !CreditNoteNo & "")
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
        Call ShowTechDrawing("CustomerComplaints", "AttachmentPDF", " WHERE EntryID=" & Val(LV.SelectedItem.Tag), LV.SelectedItem.Text, CD1)
    Else
        MsgBox "No PDF attached.", vbInformation
        Exit Sub
    End If
        
    Exit Sub
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        With LV.ColumnHeaders(i)
            If X > Val(.Left) And X < (Val(.Left) + Val(.Width)) Then
                iColNo = i
                Exit For
            End If
        End With
    Next
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 2 Then Exit Sub
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub mnuAdd_Click()
    If UserHasAccess("AddEditCustomerComplaint") = False Then
        MsgBox "Sorry, No Access.", vbInformation
        Exit Sub
    End If
    Load frmCustomerComplaintAE
    frmCustomerComplaintAE.AddNew
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

Private Sub mnuEdit_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If UserHasAccess("AddEditCustomerComplaint") = False Then
        MsgBox "Sorry, No Access.", vbInformation
        Exit Sub
    End If
     
    Dim f As Form
    
    If Val(LV.SelectedItem.ListSubItems(3).Tag) = 0 Then
        Set f = New frmCustomerComplaintAE
    ElseIf Val(LV.SelectedItem.ListSubItems(3).Tag) = 1 Then
        Set f = New frmCustomerFeedbackAE
    ElseIf Val(LV.SelectedItem.ListSubItems(3).Tag) = 2 Then
        Set f = New frmAuditAE
    End If
    
    Load f
    Call f.EditItem(Val(LV.SelectedItem.Tag))
    
End Sub

Private Sub mnuGenCreditNote_Click()

    Load frmCustomerCreditNoteAE
    frmCustomerCreditNoteAE.EditItem (Val(LV.SelectedItem.Tag))
    
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

    Dim rpt As CRAXDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CustomerComplaints.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{CustomerComplaints.EntryID}=" & Val(LV.SelectedItem.Tag), rpt
    
End Sub

Private Sub mnuPrintCapa_Click()

    Dim rpt As CRAXDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CAPA.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{CAPA.CAPANo}='" & LV.SelectedItem.SubItems(7) & "'", rpt

End Sub

Private Sub mnuPrintCapaLogoSheet_Click()
    
    Dim myDTFrom As Date, myDTTo As Date
    Dim f As New frmGetDateDouble
    If f.getDate(myDTFrom, myDTTo) = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CAPAReport.rpt")
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CAPA.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#", rpt
    
End Sub

Private Sub mnuPrintCreditNote_Click()
    Dim rpt As CRAXDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\QMS_CreditNote.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{CustomerComplaints.EntryID}=" & Val(LV.SelectedItem.Tag) & " AND {CustomerComplaints_Detail.Price}<>0", rpt
End Sub

Private Sub mnuPrintList_Click()

    Dim i As Integer, strEntryIDs As String
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strEntryIDs = strEntryIDs & Val(.Tag) & ","
            End If
        End With
    Next
    If strEntryIDs <> "" Then
        strEntryIDs = Left(strEntryIDs, Len(strEntryIDs) - 1)
    Else
        MsgBox "Please Select Entries.", vbInformation
        Exit Sub
    End If
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CustomerComplaints_List.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{CustomerComplaints.EntryID} IN[" & strEntryIDs & "]", rpt
    
End Sub

Private Sub mnuPrintListWOPrice_Click()

    Dim i As Integer, strEntryIDs As String
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strEntryIDs = strEntryIDs & Val(.Tag) & ","
            End If
        End With
    Next
    If strEntryIDs <> "" Then
        strEntryIDs = Left(strEntryIDs, Len(strEntryIDs) - 1)
    Else
        MsgBox "Please Select Entries.", vbInformation
        Exit Sub
    End If
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CustomerComplaints_List_WP.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{CustomerComplaints.EntryID} IN[" & strEntryIDs & "]", rpt

End Sub

Private Sub mnuPrintNC_Click()
    Dim rpt As CRAXDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\NC.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{NonCompliance.NCNo}='" & LV.SelectedItem.SubItems(5) & "'", rpt
End Sub
