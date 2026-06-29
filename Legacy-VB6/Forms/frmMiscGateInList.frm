VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMiscGateInList 
   ClientHeight    =   6795
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10365
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10365
   WindowState     =   2  'Maximized
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
      ItemData        =   "frmMiscGateInList.frx":0000
      Left            =   45
      List            =   "frmMiscGateInList.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   795
      Width           =   1710
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   180
      TabIndex        =   2
      Top             =   -75
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gate In"
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
         Index           =   0
         Left            =   4575
         TabIndex        =   3
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gate In"
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
         Index           =   1
         Left            =   4590
         TabIndex        =   4
         Top             =   135
         Width           =   1155
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4830
      Left            =   45
      TabIndex        =   0
      Top             =   1170
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   8520
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "GP #"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Detail"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Unit"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Packing"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Person Name"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Vehicle No."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Description"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComCtl2.DTPicker DtTo 
      Height          =   315
      Left            =   3225
      TabIndex        =   7
      Top             =   795
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
      Format          =   115015683
      CurrentDate     =   37081
   End
   Begin MSComCtl2.DTPicker DtFrom 
      Height          =   315
      Left            =   1740
      TabIndex        =   8
      Top             =   795
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
      Format          =   115015683
      CurrentDate     =   37055
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFBDCB&
      Caption         =   " Range                          From                       To"
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
      Index           =   2
      Left            =   45
      TabIndex        =   10
      Top             =   525
      Width           =   4665
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   570
      Left            =   4740
      TabIndex        =   9
      Top             =   525
      Width           =   1020
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "1799;1005"
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOpt 
      Height          =   390
      Left            =   7155
      TabIndex        =   1
      Top             =   6075
      Width           =   1560
      Caption         =   "Options         "
      PicturePosition =   327683
      Size            =   "2752;688"
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   390
      Left            =   8745
      TabIndex        =   5
      Top             =   6075
      Width           =   1560
      ForeColor       =   0
      Caption         =   "Close            "
      PicturePosition =   327683
      Size            =   "2752;688"
      Picture         =   "frmMiscGateInList.frx":006D
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&New Entry"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit Entry"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintDetail 
         Caption         =   "Print List"
      End
      Begin VB.Menu mnuPrintEntry 
         Caption         =   "Print Slip"
      End
   End
End
Attribute VB_Name = "frmMiscGateInList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbRange_Click()

    DtTo = ServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        
        Case 0
            DtFrom = DateAdd("d", -15, ServerDate)
        Case 1
            DtFrom = DateAdd("d", -30, ServerDate)
        Case 2
            DtFrom = DateAdd("d", -60, ServerDate)
        Case 3
            DtFrom = DateAdd("d", -90, ServerDate)
        Case 4
            DtFrom = DateAdd("d", -180, ServerDate)
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

Private Sub cmdOpt_Click()
   
   With cmdOpt
      Call CheckMenu
      Me.PopUpMenu mnupop, 255, .Left, .Top + .Height
   End With
   
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshList
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    LV.Move 50, cmbRange.Height + cmbRange.Top, Me.ScaleWidth - 100, Me.ScaleHeight - (cmdClose.Height + 100 + 1200)
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, LV.Top + LV.Height + 50
    
    cmdOpt.Move cmdClose.Left - cmdOpt.Width - 50, cmdClose.Top
        
    fTitle.Move (Me.Width - fTitle.Width) / 2
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    
    Call mnuEdit_Click
    
End Sub

Private Sub Form_Load()

    mnupop.Visible = False
    cmbRange.ListIndex = 0
     
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    Call CheckMenu
    Me.PopUpMenu mnupop
    
End Sub

Public Sub mnuAdd_Click()

   Load frmMiscGateInAE
   
   frmMiscGateInAE.ShowMe
   
End Sub


Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuEdit_Click()

    Load frmMiscGateInAE
    Call frmMiscGateInAE.ShowMe(Val(LV.SelectedItem.Tag))
    
End Sub

Private Sub mnuPrintDetail_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MiscGateInList.rpt")
    
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(DtFrom.Value, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy") & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport "{ItemIssuanceMisc.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#", rpt
    
End Sub

Private Sub mnuPrintEntry_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MiscGateIn.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "{MiscEntries_GateIn.EntryID}=" & Val(LV.SelectedItem.Tag), rpt
    
End Sub

Private Sub mnuRefresh_Click()
    Call RefreshList
End Sub

Private Sub RefreshList()

    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM VMiscEntries_GateIn WHERE OnlyDT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "' ORDER BY DT", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy HH:nn")
            ITM.ListSubItems.add , , !GatePassNo & ""
            ITM.ListSubItems.add , , !Detail & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , !Unit & ""
            ITM.ListSubItems.add , , !Packing & ""
            ITM.ListSubItems.add , , !PersonName & ""
            ITM.ListSubItems.add , , !VehicleNo & ""
            ITM.ListSubItems.add , , !Description & ""
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub mnuShowAll_Click()
    mnuShowAll.Checked = Not mnuShowAll.Checked
    Call RefreshList
End Sub
 
 
 
Private Sub CheckMenu()
 
    If LV.ListItems.count = 0 Then
        mnuEdit.Enabled = False
        mnuPrintDetail.Enabled = False
        
        mnuAdd.Enabled = True
        
    Else
        mnuEdit.Enabled = True
        
        mnuPrintDetail.Enabled = True
        mnuAdd.Enabled = True
        
        
    End If

    
End Sub
