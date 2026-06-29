VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmRejection_Store_Alerts 
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
      TabIndex        =   1
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rejection Alerts"
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
         Caption         =   "Rejection Alerts"
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
      Begin ComboList.Usercontrol1 cmbStore 
         Height          =   285
         Left            =   6075
         TabIndex        =   15
         Top             =   390
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
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
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   1065
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.CheckBox chkRcving 
         Caption         =   "Receiving Entries"
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
         Left            =   8610
         TabIndex        =   12
         Top             =   435
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CheckBox chkIssuance 
         Caption         =   "Issuance Entries"
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
         Left            =   8610
         TabIndex        =   11
         Top             =   180
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1605
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
         ItemData        =   "frmRejection_Store_Alerts.frx":0000
         Left            =   75
         List            =   "frmRejection_Store_Alerts.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   375
         Width           =   2100
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5820
         Left            =   60
         TabIndex        =   4
         Top             =   705
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   10266
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   900
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   14111
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   4155
         TabIndex        =   7
         Top             =   375
         Width           =   1920
         _ExtentX        =   3387
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
         Format          =   157876227
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   2205
         TabIndex        =   8
         Top             =   375
         Width           =   1920
         _ExtentX        =   3387
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
         Format          =   157941763
         CurrentDate     =   37055
      End
      Begin MSForms.CommandButton cmdPrintLots 
         Height          =   360
         Left            =   7635
         TabIndex        =   16
         Top             =   6600
         Visible         =   0   'False
         Width           =   1620
         ForeColor       =   0
         Caption         =   "Print (With Lots)"
         PicturePosition =   327683
         Size            =   "2857;635"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   9330
         TabIndex        =   14
         Top             =   6600
         Width           =   1620
         ForeColor       =   0
         Caption         =   "Print"
         PicturePosition =   327683
         Size            =   "2857;635"
         Accelerator     =   80
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
         Caption         =   " Range                                   From                                 To"
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
         Width           =   6000
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   345
         Left            =   10230
         TabIndex        =   9
         Top             =   345
         Width           =   1470
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2593;609"
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
End
Attribute VB_Name = "frmRejection_Store_Alerts"
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


Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdPOSSI_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\FS_SampleIssuance_Pending.rpt")
    Dim f As New frmPrevRpt
    Load f
    f.ShowReport "{VIssItemsSimple.SampleIssuance}=TRUE AND ISNULL({VIssItemsSimple.IssNo_Rcv})", rpt
End Sub

Private Sub cmdPrint_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Rejection_Alerts.rpt")
    Dim f As New frmPrevRpt
    Dim strSelection As String
    strSelection = "{VendReceived.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    strSelection = strSelection & " AND {UsersToDoTasks.Done}=FALSE"
    strSelection = strSelection & " AND {UsersToDoTasks.Type}='Receive_Rejection'"
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub cmdPrintLots_Click()
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\StockTransactions_Detailed.rpt")
    Dim f As New frmPrevRpt
'    Dim strSelection As String
'    strSelection = "{VStockTransactions.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
'
'    If chkIssuance.Value = vbUnchecked Then
'        strSelection = strSelection & " AND {VStockTransactions.EntryType}<>0"
'    End If
'
'    If chkRcving.Value = vbUnchecked Then
'        strSelection = strSelection & " AND {VStockTransactions.EntryType}<>1"
'    End If
    Dim iEntryType As Integer
    If chkIssuance.Value = vbChecked And chkRcving.Value = vbChecked Then
        iEntryType = -1
    End If
    If chkIssuance.Value = vbChecked Then
        iEntryType = 0
    ElseIf chkRcving.Value = vbChecked Then
        iEntryType = 1
    End If
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    rpt.ParameterFields.GetItemByName("@EntryType").SetCurrentValue iEntryType
    
    rpt.EnableParameterPrompting = False
    
    f.ShowReport "", rpt, "", False
    
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub

Private Sub Form_Load()
    
'    cmbStore.ListHeight = 4000
'    cmbStore.AddItem "<All Stores>", "0"
'
'    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "PHF" Or strCompany = "Towne" Or strCompany = "Tecnimen" Then
'        cmbStore.AddVals con, "StoreName", "Stores", "EntryID"
'        cmbStore.ID = "0"
'    End If
    
    
    Call RefreshLV
    
    cmbRange.ListIndex = 0
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim DTRange As String
    
    Screen.MousePointer = vbHourglass
    
    DTRange = " WHERE VendReceived.DT BETWEEN '" & DtFrom & "' AND '" & DtTo & "'"
    
'    If chkIssuance.Value = vbUnchecked Then
'        DTRange = DTRange & " AND EntryType<>0"
'    End If
'    If chkRcving.Value = vbUnchecked Then
'        DTRange = DTRange & " AND EntryType<>1"
'    End If
    DTRange = DTRange & " AND UsersToDoTasks.Done=0 AND UsersToDoTasks.Type='Receive_Rejection'"
    
'    If cmbStore.ID <> "0" Then
'        DTRange = DTRange & " AND Store_RefID=" & cmbStore.ID
'    End If
    
    Dim strPrevEntryNo As String
    strPrevEntryNo = ""
    Dim LColor As Long
    With rs
        SQL = "SELECT UsersToDoTasks.* FROM UsersToDoTasks INNER JOIN VendRcvdDetail ON UsersToDoTasks.TaskValue=VendRcvdDetail.EntryID INNER JOIN VendReceived ON VendRcvdDetail.RefID=VendReceived.EntryID " & DTRange & " ORDER BY UsersToDoTasks.EntryID"
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , !Description & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
        
    If CurrentUserName = "Administrator" Then
        LV.Checkboxes = True
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub LV_DblClick()
    Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.ListSubItems(2).Tag = "" Then Exit Sub
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(10).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(10).Width
        .Text = LV.SelectedItem.SubItems(9)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If CurrentUserName <> "Administrator" Then Exit Sub
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim strRet As String
    strRet = ShowPopUpMenu1("Mark Done")

    If strRet = "Mark Done" Then
        Dim i As Long
        For i = 1 To LV.ListItems.count
            With LV.ListItems(i)
                If .Checked Then
                    con.Execute "UPDATE UsersToDoTasks SET Done=1,DoneUserName='" & CurrentUserName & "',DoneMachineName='" & strComputerName & "',DoneEntryDT=getDATE() WHERE EntryID=" & Val(.Tag)
                End If
            End With
        Next
    End If
    Call RefreshLV
    
End Sub


Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LV.SelectedItem.ListSubItems(2).Tag = 0 Then 'Issuance Entry
            con.Execute "UPDATE IssItemsSimpleDetail SET Location='" & txtEdit & "' WHERE EntryID=" & LV.SelectedItem.ListSubItems(1).Tag
            con.Execute "UPDATE ItemLedgerSimple SET ItemLedgerSimple.Location='" & txtEdit & "' FROM ItemLedgerSimple " & _
             "INNER JOIN IssItemsSimpleDetail ON ItemLedgerSimple.EntryID=IssItemsSimpleDetail.LedgerRefID WHERE IssItemsSimpleDetail.EntryID=" & LV.SelectedItem.ListSubItems(1).Tag
        ElseIf LV.SelectedItem.ListSubItems(2).Tag = 1 Then
            con.Execute "UPDATE RcvItemsSimpleDetail SET Location='" & txtEdit & "' WHERE EntryID=" & LV.SelectedItem.ListSubItems(1).Tag
            con.Execute "UPDATE ItemLedgerSimple SET ItemLedgerSimple.Location='" & txtEdit & "' FROM ItemLedgerSimple " & _
             "INNER JOIN RcvItemsSimpleDetail ON ItemLedgerSimple.EntryID=RcvItemsSimpleDetail.LedgerRefID WHERE RcvItemsSimpleDetail.EntryID=" & LV.SelectedItem.ListSubItems(1).Tag
        ElseIf LV.SelectedItem.ListSubItems(2).Tag = 2 Then
            'con.Execute "UPDATE StockOrderAdjustments SET Location='" & txtEdit & "' WHERE EntryID=" & LV.SelectedItem.ListSubItems(1).Tag
            con.Execute "UPDATE StockOrderAdjustmentsDetail SET Location='" & txtEdit & "' WHERE EntryID=" & LV.SelectedItem.ListSubItems(1).Tag
        End If
        LV.SelectedItem.ListSubItems(9) = txtEdit
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub PrintIssuanceNote(strIssNo As String)
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Rejection_Issuance_Note.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VRejection_Issuance.IssNo}='" & strIssNo & "'", rpt
End Sub
