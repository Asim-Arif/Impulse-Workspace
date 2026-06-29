VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmLettersList 
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12840
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleMode       =   0  'User
   ScaleWidth      =   11500.96
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   8385
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   12750
      Begin ComboList.Usercontrol1 cmbCompany 
         Height          =   300
         Left            =   48
         TabIndex        =   6
         Top             =   456
         Width           =   7188
         _ExtentX        =   12674
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
         Height          =   300
         ItemData        =   "frmLettersList.frx":0000
         Left            =   7272
         List            =   "frmLettersList.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   468
         Width           =   1530
      End
      Begin MSComctlLib.ListView LV 
         Height          =   7116
         Left            =   48
         TabIndex        =   1
         Top             =   780
         Width           =   12588
         _ExtentX        =   22199
         _ExtentY        =   12541
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   952
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "To"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Letter Text"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Company"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   312
         Left            =   10296
         TabIndex        =   4
         Top             =   468
         Width           =   1488
         _ExtentX        =   2646
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
         Format          =   108068867
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   312
         Left            =   8796
         TabIndex        =   5
         Top             =   468
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   108068867
         CurrentDate     =   37055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   $"frmLettersList.frx":0078
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
         Height          =   252
         Index           =   0
         Left            =   48
         TabIndex        =   8
         Top             =   204
         Width           =   12588
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   11796
         TabIndex        =   7
         Top             =   420
         Width           =   840
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1482;635"
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
         Height          =   375
         Left            =   11325
         TabIndex        =   2
         Top             =   7935
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
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuRefreshDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add New Letter"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Letter"
      End
      Begin VB.Menu mnuDashPrint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Letter"
         Begin VB.Menu mnuPrintLetterFor 
            Caption         =   "Print Letter for All"
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmLettersList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbRange_Change()
    Call cmbRange_Click
End Sub

Private Sub cmbRange_Click()
    
    DtTo = ServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = ServerDate
        Case 1
            DtFrom = DateAdd("d", -15, ServerDate)
        Case 2
            DtFrom = DateAdd("d", -30, ServerDate)
        Case 3
            DtFrom = DateAdd("d", -60, ServerDate)
        Case 4
            DtFrom = DateAdd("d", -90, ServerDate)
        Case 5
            DtFrom = DateAdd("d", -180, ServerDate)
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


Private Sub cmdRefresh_Click()
    
    Call RefreshLV
    
End Sub

Private Sub Form_Load()
    
    mnuPop.Visible = False
    cmbCompany.ClearVals
    cmbCompany.AddItem "<All Companies>", "0"
    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID"
    cmbCompany.ID = "0"
    cmbRange.ListIndex = 0
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim strStatus As String
    Dim iStatus As Integer, lColor As Long
    Dim strCondition As String
    Screen.MousePointer = vbHourglass
    
    strCondition = "WHERE DT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'"
    
    If cmbCompany.ID <> "0" Then
        strCondition = strCondition & " AND Company_RefID=" & cmbCompany.ID
    End If
    
    Dim i As Integer
    With rs
        
        .Open "SELECT * FROM VLetters " & strCondition & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, !EntryID & "'", .AbsolutePosition)
            ITM.Tag = Val(!EntryID)
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , !LetterTo & ""
            ITM.ListSubItems.add , , !LetterText & ""
            ITM.ListSubItems.add , , !CompanyName & ""
            
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

Private Sub LV_DblClick()
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    Load frmNewLetter
    
    Call frmNewLetter.EditItem(LV.SelectedItem.Tag)
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then
        Call ShowMenus(False)
    Else
        Call ShowMenus(True)
    End If
    
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
End Sub

Private Sub ShowMenus(vis As Boolean)

        
    If LV.ListItems.count > 0 Then
        mnuEdit.Visible = True
        mnuPrint.Visible = True
        
        Dim lEntryID As Long
        lEntryID = Val(LV.SelectedItem.Tag)
        
        Dim rs As New ADODB.Recordset, i As Integer
        With rs
            .Open "SELECT CustCode,Country FROM Letters_Customers WHERE RefID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
            For i = mnuPrintLetterFor.UBound To 1 Step -1
                Unload mnuPrintLetterFor(i)
            Next
            i = 1
            
            Do Until .EOF
                If i > 0 Then
                    Load mnuPrintLetterFor(i)
                End If
                mnuPrintLetterFor(i).Tag = !CustCode & "|" & !Country & ""
                mnuPrintLetterFor(i).Caption = !CustCode & " (" & !Country & ")"
                
                mnuPrintLetterFor(i).Visible = True
                i = i + 1
                .MoveNEXT
            Loop
            .Close
        End With
    Else
        mnuEdit.Visible = False
        mnuPrint.Visible = False
    End If
    
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub mnuAdd_Click()
    frmNewLetter.Show 1
End Sub

Private Sub mnuedit_Click()
    Call LV_DblClick
End Sub


Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Text = i
    Next i
End Sub

Private Sub mnuPrintLetterFor_Click(Index As Integer)

    On Error GoTo err
    
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Dim frm As New frmPrevRpt
    Set rpt = rptApp.OpenReport(RptPath & "\Letter.rpt")
    If Index = 0 Then
        frm.ShowReport "{VLetters.EntryID}=" & Val(LV.SelectedItem.Tag), rpt
    Else
        frm.ShowReport "{VLetters.EntryID}=" & Val(LV.SelectedItem.Tag) & " AND {Letters_Customers.CustCode}+'|'+{Letters_Customers.Country}='" & mnuPrintLetterFor(Index).Tag & "'", rpt
    End If
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuRefresh_Click()
    Call RefreshLV
End Sub
