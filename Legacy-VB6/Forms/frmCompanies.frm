VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompanies 
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
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   180
      TabIndex        =   2
      Top             =   -75
      Width           =   10365
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Companies"
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
         Left            =   30
         TabIndex        =   3
         Top             =   120
         Width           =   10095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Companies"
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
         Left            =   45
         TabIndex        =   4
         Top             =   135
         Width           =   10095
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5385
      Left            =   45
      TabIndex        =   0
      Top             =   585
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   9499
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NTN #"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Export Reg #"
         Object.Width           =   3528
      EndProperty
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
      Picture         =   "frmCompanies.frx":0000
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
         Caption         =   "Add &New Company"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit Company"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintDetail 
         Caption         =   "Print List"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintDesignation 
         Caption         =   "Print Airline"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDashDefault 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetDefault 
         Caption         =   "Set as Defualt Company"
      End
   End
End
Attribute VB_Name = "frmCompanies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpt_Click()
   
   With cmdOpt
      Call CheckMenu
      Me.PopUpMenu mnupop, 255, .Left, .Top + .Height
   End With
   
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    LV.Move 50, 600, Me.ScaleWidth - 100, Me.ScaleHeight - (cmdClose.Height + 100 + 600)
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, LV.Top + LV.Height + 50
    
    cmdOpt.Move cmdClose.Left - cmdOpt.Width - 50, cmdClose.Top
    
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    
    Call mnuEdit_Click
    
End Sub

Private Sub Form_Load()

    
    mnupop.Visible = False
    
    Call RefreshList
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    
    Call CheckMenu
    Me.PopUpMenu mnupop
End Sub

Public Sub mnuAdd_Click()

   Load frmCompanyAE
   
   frmCompanyAE.AddNew
   
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuEdit_Click()

    Dim f As New frmCompanyAE
    Load f
    Call f.EditCompany(Val(LV.SelectedItem.Tag))
    
End Sub

Private Sub mnuPrintDesignation_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\Airline.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "{Designations.EntryID}=" & Val(LV.SelectedItem.key), rpt
    
End Sub

Private Sub mnuPrintDetail_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\Airlines.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt
    
End Sub

Private Sub mnuRefresh_Click()
    Call RefreshList
End Sub

Private Sub RefreshList()

    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM Companies ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , !CompanyName & ""
            ITM.ListSubItems.add , , !NTNNo & ""
            ITM.ListSubItems.add , , !ExportRegNo & ""
            .MoveNEXT
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

Private Sub mnuSetDefault_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
        con.Execute "UPDATE Companies SET DefaultCompany=1 WHERE EntryID=" & Val(LV.SelectedItem.Tag)
        con.Execute "UPDATE Companies SET DefaultCompany=0 WHERE EntryID<>" & Val(LV.SelectedItem.Tag)
        Call RefreshList
    End If
    
End Sub
