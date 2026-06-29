VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmRptIssRcv 
   ClientHeight    =   10425
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   10425
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicBarCode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   3855
      ScaleHeight     =   465
      ScaleWidth      =   3330
      TabIndex        =   16
      Top             =   9930
      Visible         =   0   'False
      Width           =   3330
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Print Report"
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
         Left            =   105
         TabIndex        =   7
         Top             =   165
         Width           =   10605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Print Report"
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
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   10590
      End
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   9255
      Left            =   15
      TabIndex        =   3
      Top             =   615
      Width           =   11850
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   2820
         Top             =   2700
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
               Picture         =   "frmrptIssRcv.frx":0000
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmrptIssRcv.frx":01BC
               Key             =   "Up"
            EndProperty
         EndProperty
      End
      Begin ComboList.Usercontrol1 cmbProcess 
         Height          =   285
         Left            =   180
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   360
         Width           =   2355
         _ExtentX        =   4154
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
      Begin VB.ComboBox cmbIssRcv 
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
         ItemData        =   "frmrptIssRcv.frx":0378
         Left            =   2550
         List            =   "frmrptIssRcv.frx":037A
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   360
         Width           =   1740
      End
      Begin VB.CommandButton cmdUnchkAll 
         Caption         =   "Uncheck All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1695
         TabIndex        =   15
         Top             =   8790
         Width           =   1575
      End
      Begin VB.CommandButton cmdChkAll 
         Caption         =   "Check All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   14
         Top             =   8790
         Width           =   1575
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   8010
         Left            =   45
         TabIndex        =   11
         Top             =   720
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   14129
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Non-Printed"
         TabPicture(0)   =   "frmrptIssRcv.frx":037C
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "LV"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Printed"
         TabPicture(1)   =   "frmrptIssRcv.frx":0398
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "LVPrinted"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin MSComctlLib.ListView LV 
            Height          =   7560
            Left            =   -74910
            TabIndex        =   12
            Top             =   360
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   13335
            View            =   3
            LabelEdit       =   1
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
               Text            =   "Vender ID"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Vender"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Control #"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Waste"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Lost"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView LVPrinted 
            Height          =   7560
            Left            =   90
            TabIndex        =   13
            Top             =   360
            Width           =   11565
            _ExtentX        =   20399
            _ExtentY        =   13335
            View            =   3
            LabelEdit       =   1
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
               Text            =   "Vender ID"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Vender"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Control #"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Waste"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Lost"
               Object.Width           =   1764
            EndProperty
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
         ItemData        =   "frmrptIssRcv.frx":03B4
         Left            =   4290
         List            =   "frmrptIssRcv.frx":03CA
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   2010
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   8040
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
         Format          =   3866627
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6315
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
         Format          =   3866627
         CurrentDate     =   37055
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2010
         Top             =   2070
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin Crystal.CrystalReport Cr1 
         Left            =   270
         Top             =   2040
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
      Begin MSForms.CommandButton cmdPrintDuplicate 
         Height          =   375
         Left            =   6105
         TabIndex        =   19
         Top             =   8790
         Width           =   1860
         ForeColor       =   0
         Caption         =   "Print Duplicate"
         PicturePosition =   327683
         Size            =   "3281;661"
         Picture         =   "frmrptIssRcv.frx":041B
         Accelerator     =   68
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   375
         Left            =   7995
         TabIndex        =   10
         Top             =   8790
         Width           =   1860
         ForeColor       =   0
         Caption         =   "Print Orignal   "
         PicturePosition =   327683
         Size            =   "3281;661"
         Picture         =   "frmrptIssRcv.frx":052D
         Accelerator     =   80
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
         Left            =   9885
         TabIndex        =   9
         Top             =   8790
         Width           =   1860
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "3281;661"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   9615
         TabIndex        =   2
         Top             =   300
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
         Caption         =   $"frmrptIssRcv.frx":063F
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
End
Attribute VB_Name = "frmRptIssRcv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rptSQL As String
Dim TableNAme As String

Private Sub cmbIssRcv_Change()
    Call cmbIssRcv_Click
End Sub

Private Sub cmbIssRcv_Click()

    If cmbIssRcv.ListIndex = 0 Then
        TableNAme = "VVendIssItemsrpt"
    Else
        TableNAme = "VVendRcvItemsrpt"
    End If
    
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
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub



Private Sub cmdChkAll_Click()
    If SSTab1.Tab = 0 Then
        For i = 1 To LV.ListItems.Count
            LV.ListItems(i).Checked = True
        Next i
    Else
        For i = 1 To LVPrinted.ListItems.Count
            LVPrinted.ListItems(i).Checked = True
        Next i
    End If
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()

    Call PrintRpts("Original Report")
    
End Sub

Private Sub PrintRpts(strOriginal As String)

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Dim rs As New ADODB.Recordset
    Dim rsStream As New ADODB.Stream
    
    Call StartTrans(con)
    'Insert Into PrintIssRcvRptsTemp.
    con.Execute "Delete From PrintIssRcvRptsTemp"
    
    rs.Open "Select * From PrintIssRcvRptsTemp", con, adOpenForwardOnly, adLockOptimistic
    
    For i = 1 To LV.ListItems.Count
        If LV.ListItems(i).Checked Then
            'Call DoForBarCode(LV.ListItems(i).SubItems(3) & "", rs, rsStream)
            con.Execute "Insert Into PrintIssRcvRptsTemp(RecieptID) Values('" & LV.ListItems(i).SubItems(3) & "')"
            
            'Now Move The ID's From PrintIssRcvRptsTemp To VendIssRcvPrinted.
            con.Execute "Insert Into VendIssRcvPrinted Values('" & LV.ListItems(i).SubItems(3) & "" & "')"
        End If
    Next i
    
    For i = 1 To LVPrinted.ListItems.Count
        If LVPrinted.ListItems(i).Checked Then
            'Call DoForBarCode(LVPrinted.ListItems(i).SubItems(3) & "", rs, rsStream)
            con.Execute "Insert Into PrintIssRcvRptsTemp(RecieptID) Values('" & LVPrinted.ListItems(i).SubItems(3) & "')"
        End If
    Next i
    
    
    'con.Execute "Insert Into VendIssRcvPrinted Select RecieptID From PrintIssRcvRptsTemp"
    con.CommitTrans
    
    Dim rpt As craxddrt.Report  'New rptVendIssDetail
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendIssDetail.rpt") '
    Dim ForTxt As String
'    If TableNAme = "VVendIssItemsrpt" Then
'        ForTxt = "GATEPASS ISSUED"
'    Else
'        ForTxt = "GOODS RECEIVED"
'    End If
    
    Dim FFDefs As FormulaFieldDefinitions
    Dim FFDef As craxddrt.FormulaFieldDefinition
    Set FFDefs = rpt.FormulaFields
    
    For Each FFDef In FFDefs
        'If FFDef.Name = "{@Heading}" Then
        '    FFDef.Text = "'" & ForTxt & "'"
        If FFDef.Name = "{@ForOriginal}" Then
            FFDef.Text = "'" & strOriginal & "'"
        End If
    Next
        
    Load frmPrevRpt
    frmPrevRpt.ShowReport " ", rpt
    
    Call NewRefreshLV
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub cmdPrintDuplicate_Click()
    Call PrintRpts("Duplicate Report")
End Sub

Public Sub cmdRefresh_Click()
   
    Call NewRefreshLV
   
End Sub

Private Sub cmdUnchkAll_Click()
    If SSTab1.Tab = 0 Then
        For i = 1 To LV.ListItems.Count
            LV.ListItems(i).Checked = False
        Next i
    Else
        For i = 1 To LVPrinted.ListItems.Count
            LVPrinted.ListItems(i).Checked = False
        Next i
    End If
End Sub

Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()

    cmbRange.ListIndex = 0
    cmbProcess.ListHeight = 2800
    cmbProcess.AddItem "<All Processes>", "0"
    cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " Order By SNo"
    
    cmbIssRcv.AddItem "Issue Reports"
    cmbIssRcv.AddItem "Receive Reports"
    cmbIssRcv.ListIndex = 0
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub


Public Sub ShowForIss(Optional ForRcv As Boolean, Optional ProcessID As Integer)

    If ForRcv Then
        TableNAme = "VVendRcvItemsrpt"
    Else
        TableNAme = "VVendIssItemsrpt"
    End If
    
    
    cmbProcess.ID = ProcessID
    If ProcessID <> 0 Then
        cmbProcess.Enabled = False
    End If
    
    Me.Show
    
End Sub
Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
    If cmbRange.ListIndex = 0 Then
        DTRange = " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
    ElseIf cmbRange.ListIndex = 1 Then
        DTRange = " DT Between '" & DateAdd("d", -15, ServerDate) & "' and '" & ServerDate & "' or " & " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
    ElseIf cmbRange.ListIndex = 2 Then
        DTRange = " DT Between '" & DateAdd("d", -30, ServerDate) & "' and '" & ServerDate & "' or " & " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
    ElseIf cmbRange.ListIndex = 3 Then
        DTRange = " DT Between '" & DateAdd("d", -60, ServerDate) & "' and '" & ServerDate & "' or " & " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
    ElseIf cmbRange.ListIndex = 4 Then
        DTRange = " DT Between '" & DateAdd("d", -90, ServerDate) & "' and '" & ServerDate & "' or " & " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
    ElseIf cmbRange.ListIndex = 5 Then
        DTRange = " DT Between '" & DtFrom & "' and '" & DtTo & "'"
    End If
    
    If Cond = "" Then
        Cond = " Where "
    Else
        Cond = Cond & " and "
    End If
    Cond = Cond & "(" & DTRange & ") and (ReqAuth=0)"
    
    Dim ExtraCond As String
    
    If TableNAme = "VVendIssItemsrpt" Then
        ExtraCond = " and (FromPlan Is Null) "
    Else
        ExtraCond = " and (Wastage=0)"
    End If
    
    Sql = "Select * From " & TableNAme & Cond & ExtraCond & "  Order By DT"
    
    With rs
        rptSQL = Sql
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, , .AbsolutePosition)
            ITM.SubItems(1) = ![VendID1] & ""
            ITM.SubItems(2) = ![VenderName] & ""
            'itm.SubItems(3) = ![ItemCode] & ""
            'itm.SubItems(4) = ![CompItemDesc] & ""
            ITM.SubItems(3) = ![RecieptID] & ""
            If TableNAme = "VVendIssItemsrpt" Then
                ITM.SubItems(4) = ![IssQty] & ""
            Else
                ITM.SubItems(4) = ![RcvdQty] & ""
            End If
            
            .MoveNext
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
Private Sub NewRefreshLV()
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim FromDate As Date
    Dim ToDate As Date
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    FromDate = ServerDate
    ToDate = ServerDate
    If cmbRange.ListIndex = 0 Then
        'DTRange = " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
        FromDate = ServerDate
        ToDate = vbNull
    ElseIf cmbRange.ListIndex = 1 Then
        'DTRange = " DT Between '" & DateAdd("d", -15, ServerDate) & "' and '" & ServerDate & "' or " & " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
        FromDate = DateAdd("d", -15, ServerDate)
    ElseIf cmbRange.ListIndex = 2 Then
        'DTRange = " DT Between '" & DateAdd("d", -30, ServerDate) & "' and '" & ServerDate & "' or " & " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
        FromDate = DateAdd("d", -30, ServerDate)
    ElseIf cmbRange.ListIndex = 3 Then
        'DTRange = " DT Between '" & DateAdd("d", -60, ServerDate) & "' and '" & ServerDate & "' or " & " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
        FromDate = DateAdd("d", -60, ServerDate)
    ElseIf cmbRange.ListIndex = 4 Then
        'DTRange = " DT Between '" & DateAdd("d", -90, ServerDate) & "' and '" & ServerDate & "' or " & " Convert(Varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'"
        FromDate = DateAdd("d", -90, ServerDate)
    ElseIf cmbRange.ListIndex = 5 Then
        'DTRange = " DT Between '" & DtFrom & "' and '" & DtTo & "'"
        FromDate = DtFrom
        ToDate = DtTo
    End If
    
    'con.Execute "Delete From PrintIssRcvRptsTemp"
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "SP_PrintIssRcvRpts"
        Set .ActiveConnection = con
        .Parameters("@FromDate").value = FromDate
        If Not ToDate = "12/31/1899" Then
            .Parameters("@ToDate").value = ToDate
        End If
        
        If Not TableNAme = "VVendIssItemsrpt" Then
            .Parameters("@RcvRpt").value = 1
        End If
        If cmbProcess.ID = "0" Then
            .Parameters("@ProcessID").value = Null
        Else
            .Parameters("@ProcessID").value = cmbProcess.ID
        End If
        rs.LockType = adLockReadOnly
        rs.CursorType = adOpenForwardOnly
        Set rs = .Execute
        
    End With
    
    With rs
        'rptSQL = SQL
        'For Non-Printed RecieptID's...
        .Open "Select * From VPrintIssRcvRpts1 Where RecieptID Not In(Select RecieptID From VendIssRcvPrinted) Order By DT", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, , .AbsolutePosition)
            'itm.Tag = ![UserID] & ""
            If Not IsNull(![VendID1]) Then
                ITM.SubItems(1) = ![VendID1] & ""
                ITM.SubItems(2) = ![VenderName] & ""
            Else
                ITM.SubItems(1) = ![Process] & ""
                ITM.SubItems(2) = ![Supervisor] & ""
            End If
            
            'itm.SubItems(3) = ![ItemCode] & ""
            'itm.SubItems(4) = ![CompItemDesc] & ""
            ITM.SubItems(3) = ![RecieptID] & ""
            
            ITM.SubItems(4) = Val(![Qty] & "")
            ITM.SubItems(5) = Val(![Wastage] & "")
            ITM.SubItems(6) = Val(![LostQty] & "")
            
            .MoveNext
        Loop
        .Close
        
        'For Printed RecieptID's...
        .Open "Select * From VPrintIssRcvRpts1 Where RecieptID In(Select RecieptID From VendIssRcvPrinted)", con, adOpenForwardOnly, adLockReadOnly
        LVPrinted.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LVPrinted.ListItems.Add(, , .AbsolutePosition)
            'itm.Tag = ![UserID] & ""
            If Not IsNull(![VendID1]) Then
                ITM.SubItems(1) = ![VendID1] & ""
                ITM.SubItems(2) = ![VenderName] & ""
            Else
                ITM.SubItems(1) = ![Process] & ""
                ITM.SubItems(2) = ![Supervisor] & ""
            End If
            'itm.SubItems(3) = ![ItemCode] & ""
            'itm.SubItems(4) = ![CompItemDesc] & ""
            ITM.SubItems(3) = ![RecieptID] & ""
            
            ITM.SubItems(4) = Val(![Qty] & "")
            ITM.SubItems(5) = Val(![Wastage] & "")
            ITM.SubItems(6) = Val(![LostQty] & "")
            
            .MoveNext
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

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Exit Sub
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
    Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub


Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Exit Sub
    If LV.ListItems.Count = 0 Or Button <> 2 Then Exit Sub
    
    Dim strRet As String
    strRet = ShowPopUpMenu("&Edit This Order", "-", "&Add Production Planning", "-", "History Of The Order (Summary)", "History Of The Order (Detail)", "-", "Print Order (&Summary)", "Print Order (&Detail)", "-", "&Print Order")
    
    Select Case strRet
        Case "&Edit This Order"
            
        Case "&Add Production Planning"
            
        Case "History Of The Order (Summary)"
            
        Case "History Of The Order (Detail)"
            
        Case "Print Order (&Summary)"
            
        Case "Print Order (&Detail)"
            
        Case "&Print Order"
            
    End Select
    'Me.PopupMenu mnuPop
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
End Sub
Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.Count
        LV.ListItems(i).Text = i
    Next i
End Sub
Private Sub DoForBarCode(RecieptID As String, rs As ADODB.Recordset, rsStream As ADODB.Stream)
    
    With rs
        Call DrawBarCode(PicBarCode, RecieptID)
        SavePicture PicBarCode.Image, "C:\tmp.bmp"
        rsStream.Open
        rsStream.Type = adTypeBinary
        rsStream.LoadFromFile "C:\tmp.bmp"
        
        
        .AddNew
        .Fields![RecieptID] = RecieptID
        rs.Fields(1) = rsStream.Read
        rsStream.Close
        
        .Update
    End With
    
    
End Sub



