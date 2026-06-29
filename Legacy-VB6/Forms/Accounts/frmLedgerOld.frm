VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmLedgerOld 
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   12750
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7665
   ScaleMode       =   0  'User
   ScaleWidth      =   11420.35
   WindowState     =   2  'Maximized
   Begin VB.Frame FraFind 
      Height          =   480
      Left            =   45
      TabIndex        =   14
      Top             =   7125
      Width           =   5475
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1965
         TabIndex        =   17
         Top             =   135
         Width           =   2430
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Enabled         =   0   'False
         Height          =   300
         Left            =   4410
         TabIndex        =   16
         Top             =   135
         Width           =   1020
      End
      Begin VB.ComboBox cmbfield 
         Height          =   315
         ItemData        =   "frmLedgerOld.frx":0000
         Left            =   60
         List            =   "frmLedgerOld.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   135
         Width           =   1875
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   1935
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   4320
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   930
      Top             =   180
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
   Begin VB.Frame FDrag 
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   60
      MousePointer    =   7  'Size N S
      TabIndex        =   8
      Top             =   4155
      Width           =   11805
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   15
      TabIndex        =   10
      Top             =   30
      Width           =   11190
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   7965
         TabIndex        =   3
         Top             =   345
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   65273859
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6210
         TabIndex        =   2
         Top             =   345
         Width           =   1755
         _ExtentX        =   3096
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
         Format          =   65273859
         CurrentDate     =   36892
         MinDate         =   36892
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   720
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerOld.frx":004E
               Key             =   "PND"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerOld.frx":04A6
               Key             =   "CAN"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerOld.frx":08F8
               Key             =   "CLR"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerOld.frx":0D4C
               Key             =   "DUE"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerOld.frx":11A4
               Key             =   "BNC"
            EndProperty
         EndProperty
      End
      Begin MSForms.CommandButton cmdRefresh 
         Default         =   -1  'True
         Height          =   360
         Left            =   9615
         TabIndex        =   4
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
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmLedgerOld.frx":15F8
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
         Height          =   210
         Index           =   0
         Left            =   15
         TabIndex        =   9
         Top             =   120
         Width           =   11130
      End
      Begin MSForms.ComboBox cmbAccNo 
         Height          =   315
         Left            =   30
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   345
         Width           =   1890
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3334;556"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox CmbAccTitle 
         Height          =   315
         Left            =   1905
         TabIndex        =   1
         Top             =   345
         Width           =   4305
         VariousPropertyBits=   545277979
         DisplayStyle    =   3
         Size            =   "7594;556"
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3375
      Left            =   45
      TabIndex        =   5
      Top             =   765
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   5953
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dated"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Voucher #"
         Object.Width           =   2515
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   8440
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Debit"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Credit"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Balance"
         Object.Width           =   2757
      EndProperty
   End
   Begin MSComctlLib.ListView LVCheque 
      Height          =   2925
      Left            =   45
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   5159
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15330543
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dated"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cheque #"
         Object.Width           =   2068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bank & Branch"
         Object.Width           =   3939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   4529
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Type"
         Object.Width           =   1969
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Chq. Date"
         Object.Width           =   2068
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Status"
         Object.Width           =   1575
      EndProperty
   End
   Begin MSForms.CommandButton cmdPrintAll 
      Height          =   360
      Left            =   8265
      TabIndex        =   13
      Top             =   7200
      Visible         =   0   'False
      Width           =   1425
      ForeColor       =   64
      Caption         =   "Print All"
      PicturePosition =   327683
      Size            =   "2514;635"
      Accelerator     =   65
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
      Left            =   11175
      TabIndex        =   12
      Top             =   7200
      Width           =   1425
      ForeColor       =   64
      Caption         =   "Close     "
      PicturePosition =   327683
      Size            =   "2514;635"
      Accelerator     =   67
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
      Left            =   9720
      TabIndex        =   7
      Top             =   7200
      Width           =   1425
      ForeColor       =   64
      Caption         =   "Print         "
      PicturePosition =   327683
      Size            =   "2514;635"
      Accelerator     =   80
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Pop Menu"
      Begin VB.Menu mnuToCash 
         Caption         =   "Transfer to &Cash In Hand"
         Begin VB.Menu mnuOnChqDate 
            Caption         =   "On Cheque's Date"
         End
         Begin VB.Menu mnuOnToday 
            Caption         =   "On Today's Date"
         End
      End
      Begin VB.Menu mnuToBank 
         Caption         =   "Transfer to Bank &Account On Cheque's Date"
         Begin VB.Menu mnuBanks 
            Caption         =   "< No Bank Account Defined >"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuToBankToday 
         Caption         =   "Transfer to Bank &Account On Today's Date"
         Begin VB.Menu mnuBanksToday 
            Caption         =   "< No Bank Account Defined >"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChqBounced 
         Caption         =   "Cheque &Bounced"
      End
   End
   Begin VB.Menu mnuClear 
      Caption         =   "mnuClear"
      Visible         =   0   'False
      Begin VB.Menu mnuChqCashed 
         Caption         =   "Make This Cheque &Cleared On Cheque's Date"
      End
      Begin VB.Menu mnuChqToday 
         Caption         =   "Make This Cheque &Cleared On Today's Date"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBounced 
         Caption         =   "Chque &Bounced"
      End
      Begin VB.Menu dashprint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprintChq 
         Caption         =   "Print This Cheque"
      End
   End
   Begin VB.Menu mnuopts 
      Caption         =   "opts"
      Begin VB.Menu mnuComparison 
         Caption         =   "View Comparison Ledger For"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "View Transaction Report For"
      End
      Begin VB.Menu mnuMonthTrans 
         Caption         =   "View Monthly Transaction For"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDuePR 
         Caption         =   "View Due Payments"
      End
      Begin VB.Menu dashopt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBalanceTag 
         Caption         =   "Insert  Balance OK Tag Here"
      End
   End
   Begin VB.Menu mnutag 
      Caption         =   "mnutags"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveTag 
         Caption         =   "Remove This Tag"
      End
   End
End
Attribute VB_Name = "frmLedgerOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReadyToDrag As Boolean
Dim BankAccNo As String
Dim dontShowList As Boolean
Dim PermissionDenied As Boolean

Private Sub CmbAccTitle_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode = vbKeyDown Then
        If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex + 1
        KeyCode = 0
        dontShowList = True
    ElseIf KeyCode = vbKeyUp Then
        If List1.ListIndex > 0 Then List1.ListIndex = List1.ListIndex - 1
        KeyCode = 0
        dontShowList = True
    ElseIf KeyCode = 13 And List1.Visible Then
        CmbAccTitle = List1
        List1.Visible = False
        CmbAccTitle.SetFocus
    End If

End Sub

Private Sub CmbAccTitle_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If dontShowList Then dontShowList = False Else ShowList List1, CmbAccTitle, CmbAccTitle
End Sub

Private Sub CmbAccTitle_DropButtonClick()
    List1.Visible = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdFind_Click()

    Dim StartIdx As Integer
    
    If LV.SelectedItem Is Nothing Then
       StartIdx = 1
    Else
       If LV.SelectedItem.Index = LV.ListItems.count Then
          StartIdx = 1
       Else
          StartIdx = LV.SelectedItem.Index
       End If
    End If

    Dim fldValue As String
    Dim strtemp As String
    For i = StartIdx To LV.ListItems.count
       
        If cmbField.ListIndex = 0 Then
            fldValue = Trim(Replace(LV.ListItems(i), "»", ""))
        ElseIf cmbField.ListIndex < 3 Then
            fldValue = Trim(Replace(LV.ListItems(i).SubItems(cmbField.ListIndex), "»", ""))
        Else
            strtemp = LV.ListItems(i).SubItems(cmbField.ListIndex)
            If Len(strtemp) > 3 Then
                fldValue = Format(Left(strtemp, Len(strtemp) - 3), "#")
            Else
                fldValue = strtemp
            End If
            
        End If
    
    
        If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
            LV.SetFocus
            If i <> StartIdx Then Exit For
        End If
       
    Next

    If i = LV.ListItems.count + 1 Then
        MsgBox "Item Not Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If

End Sub

Private Sub cmdPrint_Click()

    Screen.MousePointer = vbHourglass

    With cr1
        .ReportFileName = RptPath & "\ledger.rpt"
        '    .use
        .SubreportToChange = ""

        .Connect = ConnectStr
        .SelectionFormula = "{VLedger.AccNo}='" & cmbAccNo & "' AND ({VLedger.VDate}=date(" & Format(DtFrom.Tag, "yyyy,MM,dd") & ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
        .Formulas(0) = "Company='" & company & "'"
    
        .Formulas(1) = "FromTo='From " & Format(DtFrom.Tag, "dd-MMM-yyyy") & " To " & Format(DtTo.Tag, "dd-MMM-yyyy") & "'"
        .Formulas(2) = "SubReport=" & LVCheque.Visible
    
        .Formulas(3) = "Address='" & COMPANYADDRESS & "'"
        'if Cheques Are Present then Show Cheque Detail Also
        If LVCheque.Visible Then
            .SubreportToChange = "ChqLedger"
            .Connect = ConnectStr
            .SelectionFormula = "{VChqLedger.AccNo}='" & cmbAccNo & "' AND ({VChqLedger.CDate}=date(" & Format(DtFrom.Tag, "yyyy,MM,dd") & ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
        End If
        .Action = 1
        .PageZoomNext
    End With
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdPrintAll_Click()

    Screen.MousePointer = vbHourglass
    
    With cr1
        .ReportFileName = RptPath & "\Ledger.rpt"
    '    .use
        .SubreportToChange = ""
        DtFrom.Tag = DtFrom.value
        DtTo.Tag = DtTo.value
        .Connect = ConnectStr
        .SelectionFormula = "{VLedger.VDate}=date(" & Format(DtFrom.value, "yyyy,MM,dd") & ") TO date(" & Format(DtTo.value, "yyyy,MM,dd") & ")"
        
        .Formulas(0) = "Company='" & company & "'"
        
        
        .Formulas(1) = "FromTo='From " & Format(DtFrom.Tag, "dd-MMM-yyyy") & " To " & Format(DtTo.Tag, "dd-MMM-yyyy") & "'"
        .Formulas(2) = "SubReport=" & True
        
        .Formulas(3) = "Address='" & COMPANYADDRESS & "'"
        .Formulas(4) = "ForDTFrom=#" & Format(DtFrom.Tag, "yyyy,MM,dd") & "#"
        .Formulas(5) = "ForDTTo=#" & Format(DtTo.Tag, "yyyy,MM,dd") & "#"
        '.SubreportToChange = "ChqLedger"
        '.Connect = ConnectStr
        '.SelectionFormula = "{VChqLedger.AccNo}='" & cmbAccNo & "' AND ({VChqLedger.CDate}=date(" & Format(DtFrom.Tag, "yyyy,MM,dd") & ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
        
        .Action = 1
        .PageZoomNext
    End With
    Screen.MousePointer = vbDefault

End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLedger
    cmbAccNo.Tag = cmbAccNo
    CmbAccTitle.Tag = CmbAccTitle
    Call Form_Resize
End Sub

Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub

Private Sub FDrag_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReadyToDrag = True
    FDrag.BackColor = &H808080
    FDrag.ZOrder
End Sub

Private Sub FDrag_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim TheTop As Long
    TheTop = FDrag.Top + y

    If ReadyToDrag Then
        If (TheTop > (FAcc.Top + FAcc.Height + 500)) And (TheTop < cmdClose.Top - 500) Then
            FDrag.Top = TheTop
        End If
    End If
    
End Sub

Private Sub FDrag_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If ReadyToDrag Then
    
        FDrag.BackColor = Me.BackColor
        LV.Height = FDrag.Top - LV.Top
        LVCheque.Move 50, FDrag.Top + FDrag.Height, LVCheque.Width, Me.ScaleHeight - (FAcc.Height + cmdClose.Height + LV.Height + FDrag.Height + 200)
        If LVCheque.Top > cmdClose.Top - 500 Then
            LVCheque.Top = cmdClose.Top - 500
        End If
        
        ReadyToDrag = False
        
        FDrag.ZOrder 1
    End If
    
End Sub

Private Sub Form_Activate()
    'Check For User Access
    If PermissionDenied Then
        Me.Hide
        MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()

    mnupop.Visible = False
    mnuopts.Visible = False
    
    'PermissionDenied = Not UserHasAccess("DetailedLedger")
   
    DtTo = DateAdd("yyyy", 1, Date)
    DtFrom = DateAdd("M", -1, Date)
    Call AddToCombo(cmbAccNo, "AccNo", "Accounts", " Where Parent=0 Order By AccNo")
    Call AddToCombo(CmbAccTitle, "AccTitle", "Accounts", " Where Parent=0 Order By AccNo")
   
    Call GetBankAccountsInMenu
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long

    FAcc.Move (Me.ScaleWidth - FAcc.Width) / 2
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 100
    cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
    cmdPrintAll.Move cmdPrint.Left - cmdPrintAll.Width - 50, cmdClose.Top
    FraFind.Top = cmdPrint.Top - 100
    FDrag.Move 50, Me.ScaleHeight * 3 / 5, Me.ScaleWidth - 100
 
    If LVCheque.Visible Then
        TheHeight = FDrag.Top
        FDrag.Visible = True
    Else
        TheHeight = Me.ScaleHeight - cmdClose.Height - 200
        FDrag.Visible = False
    End If
 
    LV.Move 50, FAcc.Top + FAcc.Height, Me.ScaleWidth - 100, TheHeight - (FAcc.Top + FAcc.Height)
    LVCheque.Move 50, FDrag.Top + FDrag.Height, Me.ScaleWidth - 100, Me.ScaleHeight - (FAcc.Height + cmdClose.Height + LV.Height + 230)
    FDrag.ZOrder 1
 
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = Len(Trim(txtFind)) > 0
End Sub

Private Sub txtFind_GotFocus()
   txtFind.SelStart = 0
   txtFind.SelLength = Len(txtFind)
   cmdFind.Default = True
End Sub


Private Sub cmbAccNo_Change()
    If cmbAccNo.MatchFound Then Call cmbAccNo_Click
End Sub

Private Sub cmbAccNo_Click()

    If CmbAccTitle.ListIndex <> cmbAccNo.ListIndex Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
    
End Sub

Private Sub CmbAccTitle_Change()
    If CmbAccTitle.MatchFound Then
      Call cmbAccTitle_Click
    End If
      'List1.Visible = True
      
End Sub

Private Sub cmbAccTitle_Click()
    If cmbAccNo.ListIndex <> CmbAccTitle.ListIndex Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
End Sub


Private Sub RefreshLedger()

    Dim SnoList As String      'Variable for Balance Tags
    SnoList = ""

    Dim TotalCr As Double, TotalDr As Double, PrevDate As Date, b As Boolean
    Dim LColor As Long
    If Not cmbAccNo.MatchFound Then Exit Sub

    Screen.MousePointer = vbHourglass
    Dim rsLedger As New ADODB.Recordset
    Sql = "Select * From VLedger Where AccNo='" & cmbAccNo & "' and (VDate between '" & DtFrom & "' and '" & DtTo & "') Order By VDate,SNo"

    With rsLedger
        .CursorLocation = adUseClient
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        Dim ITM As ListItem, Balance As String
 
        LV.ListItems.Clear
        
        If .RecordCount > 0 Then
            Balance = Val(![Balance] & "") + Val(![Credit] & "") - Val(![Debit] & "")
        Else
            '        Balance = Val(con.Execute("Select Balance From Accounts Where AccNo='" & cmbAccNo & "'")![Balance] & "")
            Balance = getBalance(cmbAccNo, DtFrom.value)
        End If
    
        Set ITM = LV.ListItems.add(, , "")
        ITM.SubItems(2) = Space(25) & "<<< Balance Forward >>>"
        If Val(Balance) > 0 Then
            Balance = Format(Abs(Val(Balance)), "###,##0.00") & "  DR"
        ElseIf Val(Balance) < 0 Then
            Balance = Format(Abs(Val(Balance)), "###,##0.00") & "  CR"
        Else
            Balance = Format(Val(Balance), "###,##0.00") & "       "
        End If
 
        ITM.SubItems(5) = Balance
 
        For i = 0 To .RecordCount - 1
            
            Set ITM = LV.ListItems.add(, ![Sno] & "'", Format(![VDate], "dd-MMM-yyyy"))
            SnoList = SnoList & ![Sno] & ","
            If ![VDate] <> PrevDate Then
               PrevDate = ![VDate]
            Else
               ITM.ForeColor = RGB(125, 125, 225)
            End If
      
            'Add The Entries
            ITM.ListSubItems.add , , ![VchrNo] & ""
            ITM.ListSubItems.add , , ![Description] & IIf(Len(![DpstSlip] & "") = 0, "", "[Dp.Slip #:" & ![DpstSlip] & "]")
            ITM.ListSubItems.add , , Format(Val(![Debit] & ""), "###,##0.00")
            ITM.ListSubItems.add , , Format(Val(![Credit] & ""), "###,##0.00")
            ITM.Tag = Val(![CSNo] & "")
        
            'Specify Balance as DBT OR CDT
            If Val(![Balance] & "") > 0 Then
                Balance = Format(Abs(Val(![Balance] & "")), "###,##0.00") & "  DR"
            ElseIf Val(![Balance] & "") < 0 Then
                Balance = Format(Abs(Val(![Balance] & "")), "###,##0.00") & "  CR"
            Else
                Balance = Format(Val(![Balance] & ""), "###,##0.00") & "       "
            End If
        
            ITM.ListSubItems.add , , Balance
            If Right(Balance, 2) = "CR" Then
                LColor = vbRed
            Else
                LColor = vbBlue
            End If
            '        itm.SubItems(6) = ![DpstSlip] & ""
      
            ' Calculate Total
            TotalCr = TotalCr + Val(![Credit] & "")
            TotalDr = TotalDr + Val(![Debit] & "")
            ITM.ForeColor = LColor
            For j = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(j).ForeColor = LColor
            Next
            .MoveNext
        
        Next
    
    End With


    If LV.ListItems.count > 1 Then
        Set ITM = LV.ListItems.add(, , "")
        ITM.SubItems(2) = Space(30) & "<<< Last Balance >>>"
        ITM.SubItems(5) = LV.ListItems(LV.ListItems.count - 1).SubItems(5)
        
        Set ITM = LV.ListItems.add(, , "")
        ITM.SubItems(2) = Space(60) & "Totals :"
        ITM.SubItems(3) = Format(TotalDr, "#,##0.00")
        ITM.SubItems(4) = Format(TotalCr, "#,##0.00")
    End If

    DtTo.Tag = DtTo
    DtFrom.Tag = DtFrom
     
     
    '"(VDate" & Format(DtFrom.value, "(yyyy,MM,dd)") & " to Date" & Format(DtTo.value, "(yyyy,MM,dd)") & ") and ({ledgerrpt.acno}='" & Trim(Text16.Text) & "')"
    'Now Fill The Other List With Cheques
    Call RefreshChequeLedger
    'Me.Caption = "Ledger of " & CmbAccTitle

    If LV.ListItems.count > 1 Then Call RefreshBalanceTags(SnoList)

    Screen.MousePointer = vbDefault

End Sub

Private Sub RefreshChequeLedger()

    Screen.MousePointer = vbHourglass
    Set LVCheque.SmallIcons = ImageList1
    Dim ListColor As Long
    Dim CrsLedger As New ADODB.Recordset
    
    With CrsLedger
        .CursorLocation = adUseClient
        .Open "Select * from VChqLedger Where (AccNo='" & cmbAccNo & "' OR BankAccNo='" & cmbAccNo & "') " & _
            "and (CDate between '" & DtFrom & "' and '" & DtTo & "') " & _
            "Order by CDate,Sno", con, 0, 1
            
        LVCheque.ListItems.Clear
        If .RecordCount > 0 Then
            Dim ITM As ListItem
            Do Until .EOF
                If ![Bounced] Or ![ChequeType] = "Canceled" Then
                  ListColor = vbRed
                ElseIf (![ChequeDate] <= Date And Not ![Posted]) Then
                  ListColor = vbBlue
                ElseIf ![Posted] Then
                  ListColor = RGB(0, 125, 0)
                Else
                  ListColor = vbBlack
                End If
                
                Set ITM = LVCheque.ListItems.add(, ![Sno] & "'", Format(![CDate], "dd-MMM-yyyy"))
                ITM.ForeColor = ListColor
                
                
                
                For i = 1 To 7
                    With ITM.ListSubItems.add(, , "")
                     .ForeColor = ListColor
                    End With
                    
                Next i
                
                ITM.ListSubItems(1) = ![ChequeNo] & ""
                ITM.ListSubItems(2) = ![Bank] & Space$(3) & ![Branch]
                ITM.ListSubItems(3) = ![Description]
                ITM.ListSubItems(4) = Format(Val(![Amount] & ""), "###,##0.00")
                ITM.ListSubItems(5) = ![ChequeType] & ""
                ITM.ListSubItems(6) = Format(![ChequeDate], "dd-MMM-yyyy")
                
                'Check For The Chq Status
                Dim Status As String, img As String
                If ![ChequeType] = "Canceled" Then
                  ListColor = vbRed
                  img = "CAN"
                ElseIf ![Posted] Then
                  Status = "Cleared"
                  img = "CLR"
                ElseIf ![Bounced] Then
                  Status = "Bounced"
                  img = "BNC"
                ElseIf ITM.ForeColor = vbBlue Then
                  Status = "Due"
                  img = "DUE"
                Else
                  Status = "Pending"
                  img = "PND"
                End If
                 
                 ITM.ListSubItems(7) = Status
                 ITM.SmallIcon = img
                 
                'Set Flag That it Was Payment Or Recept
                If ![Payment] Then ITM.Tag = "P" Else ITM.Tag = "R"
                
                
                
                .MoveNext
                LVCheque.Visible = True
                FDrag.Visible = True
            Loop
            
        Else
            LVCheque.Visible = False
            FDrag.Visible = False
        End If
        .Close
    End With
    
    Set CrsLedger = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub List1_Click()
If UCase(Screen.ActiveControl.Name) <> "CMBACCTITLE" Then
   CmbAccTitle = List1
   List1.Visible = False
   CmbAccTitle.SetFocus
End If
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next

If Not LVCheque.Visible Then Exit Sub

LVCheque.HideSelection = True
With LVCheque.SelectedItem
   .Bold = False
   For i = 1 To .ListSubItems.count
      .ListSubItems(i).Bold = False
   Next
   
End With



 If Val(Item.Tag) > 0 Then


    With LVCheque.ListItems(Val(Item.Tag) & "'")
     .Selected = True
     .EnsureVisible
     .Ghosted = True
     
      .Bold = True
      
      For i = 1 To .ListSubItems.count
         .ListSubItems(i).Bold = True
      Next
      
      If err.Number = 0 Then LVCheque.HideSelection = False

    End With
 End If
 
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Or LV.ListItems.count <= 1 Then Exit Sub
    If UCase(CurrentUserName) = "AUDIT" Then Exit Sub
    
    If LV.SelectedItem.key Like "TAG*" Then
        PopUpMenu mnutag
    Else
        mnuTrans.Caption = "View Transaction Report For " & LV.SelectedItem
        mnuComparison.Caption = "View Comparison Ledger Of " & CmbAccTitle
        mnuMonthTrans.Caption = "View Monthly Transaction Of " & CmbAccTitle
   
        If cmbAccNo Like "15-003*" Then
           mnuDuePR.Caption = "View Due Receipts From " & CmbAccTitle
           mnuDuePR.Tag = ""
           mnuDuePR.Visible = True
        ElseIf cmbAccNo Like "24-004*" Then
           mnuDuePR.Caption = "View Due Payments To " & CmbAccTitle
           mnuDuePR.Visible = True
           mnuDuePR.Tag = "P"
        Else
            mnuDuePR.Visible = False
        End If
   
        PopUpMenu mnuopts
   
    End If
    
End Sub

Private Sub LVCheque_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Checked = Not Item.Checked
End Sub

Private Sub LVCheque_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If UCase(CurrentUserName) = "AUDIT" Then Exit Sub
    With LVCheque
    
        If .ListItems.count = 0 Or Button <> 2 Then Exit Sub
        
        If .SelectedItem.ForeColor <> vbBlue Then Exit Sub
        
        If .SelectedItem.ListSubItems(5).Text = "Open" Then
            mnuToCash.Visible = True
        Else
            mnuToCash.Visible = False
        End If
        If .SelectedItem.Tag = "R" Then         ' Receipt Chq
         Me.PopUpMenu mnupop
        ElseIf .SelectedItem.Tag = "P" Then         ' Payment Chq
          Me.PopUpMenu mnuClear
        End If
    End With
    
End Sub

Private Function SaveReceiptTrans(ReceiveInAccNo As String, Dated As Date, Optional SlipNo As String) As Boolean

    Dim i As Integer
    Dim NextVchrNo As String
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim dAmount As Double
    Dim sScNo As String
    Dim sCIH As String
    Dim Bal As Double
    Dim Desc As String
    
    'Get Serial # and Balances
    NextVchrNo = getNextVchrNo(Dated, "RV")
    INextSno = getNextSno(Dated)
    SnoToUse = INextSno
    Bal = getBalance(cmbAccNo.Tag, Dated)
    dAmount = Format$(LVCheque.SelectedItem.ListSubItems(4).Text, "###")
    sScNo = Left$(LVCheque.SelectedItem.key, Len(LVCheque.SelectedItem.key) - 1)
    sCIH = ReceiveInAccNo
    Bal = Bal - dAmount
    'Start New Transaction
    con.BeginTrans
    'Credit Entry
    con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & cmbAccNo.Tag & _
        "','" & LVCheque.SelectedItem.ListSubItems(3).Text & _
        "',0," & dAmount & "," & Bal & ",'" & SlipNo & "'," & sScNo & ")"
    
    Call adjustLedger(cmbAccNo.Tag, Dated, INextSno, -Val(dAmount))
    
    Dim ChqNo As String
    INextSno = INextSno + 1
    Bal = getBalance(sCIH, Dated)
    Bal = Bal + dAmount
    'Description
    Dim ChqRec As New ADODB.Recordset
    With ChqRec
        .CursorLocation = adUseClient
        .Open "Select ChequeNo From Cheque Where (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            Desc = "Cheque Recieved from " & CmbAccTitle.Tag & " Chq# : " & _
                .Fields(0) & ""
             ChqNo = .Fields(0) & ""
        Else
            Desc = "Cheque Recieved from " & CmbAccTitle.Tag & " Chq# : "
        End If
        
        .Close
    End With
    Set ChqRec = Nothing
    'Debit Entry
    
    con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & sCIH & _
        "','" & Desc & _
        "'," & dAmount & ",0," & Bal & ",'" & SlipNo & "'," & sScNo & ")"
        
    Call adjustLedger(sCIH, Dated, INextSno, Val(dAmount))
    
    Dim rsCopy As New ADODB.Recordset
    With rsCopy
        .CursorLocation = adUseClient
        .Open "Select * from ChequeDetail Where (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            con.Execute "Insert into VoucherDetail Values('" & ![BalancedSno] & "','" & SnoToUse & "'," & ![Amount] & ")"
            .MoveNext
        Loop
        .Close
    End With
    Set rsCopy = Nothing
    con.Execute "Update Cheque Set Posted=1 Where Sno=" & sScNo
    'Commit Transaction
    con.CommitTrans
    
    SaveReceiptTrans = True
    Exit Function

err:
    MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
    con.RollbackTrans
        SaveReceiptTrans = False
End Function

Private Sub mnuBalanceTag_Click()
On Error GoTo err
   con.Execute "insert into BalanceTags  Values(" & Val(LV.SelectedItem.key) & ")"
    With LV.ListItems.add(LV.SelectedItem.Index + 1, "TAG" & Val(LV.SelectedItem.key))
            .ListSubItems.add
            .ListSubItems.add(, , "*-*-*-*-*-*-*-*-*<  Balance OK  >*-*-*-*-*-*-*-*-*").ForeColor = RGB(0, 125, 0)
         End With
   
   Exit Sub
err:
   MsgBox "There is Already An OK TAG Here Cannot Add Another.", vbInformation
End Sub

Private Sub mnuBanks_Click(Index As Integer)

If MsgBox("Are You Sure To Transfer This Cheque To " & mnuBanks(Index).Caption, vbInformation + vbYesNo) = vbNo Then Exit Sub
   Dim SlipNo As String
   SlipNo = InputBox("Enter Deposit Slip No.")
   
   If Trim(SlipNo) = "" Then
      MsgBox "Transaction Terminated Due To Invlalid Bank Deposit SlipNo", vbInformation
      Exit Sub
   End If
   
   
   Screen.MousePointer = vbHourglass
   Dim Dated As Date
   Dated = CDate(LVCheque.SelectedItem.SubItems(6))
   If SaveReceiptTrans(mnuBanks(Index).Tag, Dated, SlipNo) Then
      Call PlaySaveSound
      MsgBox "Transaction Saved Successfully.", vbInformation
      Call RefreshLedger
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuBanksToday_Click(Index As Integer)
If MsgBox("Are You Sure To Transfer This Cheque To " & mnuBanks(Index).Caption, vbInformation + vbYesNo) = vbNo Then Exit Sub
   Dim SlipNo As String
   SlipNo = InputBox("Enter Deposit Slip No.")
   
   If Trim(SlipNo) = "" Then
      MsgBox "Transaction Terminated Due To Invlalid Bank Deposit SlipNo", vbInformation
      Exit Sub
   End If
   
   
   Screen.MousePointer = vbHourglass
   
   If SaveReceiptTrans(mnuBanks(Index).Tag, Date, SlipNo) Then
      Call PlaySaveSound
      MsgBox "Transaction Saved Successfully.", vbInformation
      Call RefreshLedger
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuBounced_Click()
If MsgBox("Are You Sure To Make This Cheque Bounced.", vbInformation + vbYesNo) = vbNo Then Exit Sub
Screen.MousePointer = vbHourglass
   con.Execute "Update Cheque Set Bounced=1 Where Sno=" & Val(LVCheque.SelectedItem.key)
   ''sleep (1000)
   Call RefreshChequeLedger
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuChqBounced_Click()
   If MsgBox("Are You Sure To Make This Cheque Bounced.", vbInformation + vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
      con.Execute "Update Cheque Set Bounced=1 Where Sno=" & Val(LVCheque.SelectedItem.key)
      'sleep (1000)
      Call RefreshChequeLedger
      Screen.MousePointer = vbDefault
End Sub

Private Sub mnuChqCashed_Click()
If MsgBox("Are You Sure To Make This Cheque Cleared.", vbInformation + vbYesNo) = vbNo Then Exit Sub

   Screen.MousePointer = vbHourglass
   Dim Dated As Date
   Dated = CDate(LVCheque.SelectedItem.SubItems(6))
   If SavePaymentTrans(Dated) Then
      Call PlaySaveSound
      MsgBox "Transaction Saved Successfully.", vbInformation
      Call RefreshLedger
   End If
       
   Screen.MousePointer = vbDefault
   
   
End Sub

Private Sub mnuChqToday_Click()
If MsgBox("Are You Sure To Make This Cheque Cleared.", vbInformation + vbYesNo) = vbNo Then Exit Sub

   Screen.MousePointer = vbHourglass
   
   If SavePaymentTrans(Date) Then
      Call PlaySaveSound
      MsgBox "Transaction Saved Successfully.", vbInformation
      Call RefreshLedger
   End If
       
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuComparison_Click()
With frmMonthLedger
   .cmbAccNo = cmbAccNo
   .DtFrom = DtFrom
   .DtTo = DtTo
   .Show
   Call .ShowLedger
   .SetFocus
End With
End Sub

Private Sub mnuDuePR_Click()
    
    Dim f As New frmDueRcptsPymts
 
    Call f.FillCmbs(CBool(mnuDuePR.Tag = "P"))
    f.Show
    f.cmbAccNo = Me.cmbAccNo
    f.SetFocus
    
End Sub

Private Sub mnuMonthTrans_Click()

'    With frmComparison
'       .DtFrom = DtFrom
'       .DtTo = DtTo
'       .cmbAccNo = cmbAccNo
'       .cmdRefresh_Click
'       .SetFocus
'    End With

End Sub

Private Sub mnuOnChqDate_Click()
If MsgBox("Are You Sure To Transfer This Cheque To Cash In Hand.", vbInformation + vbYesNo) = vbNo Then Exit Sub
   Screen.MousePointer = vbHourglass
   Dim Dated As Date
   Dated = CDate(LVCheque.SelectedItem.SubItems(6))
   If SaveReceiptTrans(CashInHand, Dated) Then
      
      Call RefreshLedger
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuOnToday_Click()

    If MsgBox("Are You Sure To Transfer This Cheque To Cash In Hand.", vbInformation + vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    If SaveReceiptTrans(CashInHand, Date) Then
        Call RefreshLedger
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub GetBankAccountsInMenu()

  Dim rs As New ADODB.Recordset
  With rs
    .Open "Select * from VBankAccounts Order By AccNo", con, adOpenForwardOnly, adLockReadOnly
  
  For i = 0 To .RecordCount - 1
  
   Load mnuBanks(i + 1)
   mnuBanks(i + 1).Caption = "[" & ![AccNo] & "] " & ![AccTitle]
   mnuBanks(i + 1).Enabled = True
   mnuBanks(i + 1).Tag = ![AccNo] & ""
   
   Load mnuBanksToday(i + 1)
   mnuBanksToday(i + 1).Caption = "[" & ![AccNo] & "] " & ![AccTitle]
   mnuBanksToday(i + 1).Enabled = True
   mnuBanksToday(i + 1).Tag = ![AccNo] & ""
   
  .MoveNext
  
  Next
   mnuBanks(0).Visible = .RecordCount = 0
   mnuBanksToday(0).Visible = .RecordCount = 0
  .Close
 End With
 Set rs = Nothing
  
End Sub






Private Function SavePaymentTrans(Dated As Date) As Boolean

    Dim i As Integer
    Dim NextVchrNo As String
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim dAmount As Double
    Dim sScNo As String
    Dim sCIH As String
    Dim Bal As Double
    Dim Desc As String
    'Get Serial # and Balances
    NextVchrNo = getNextVchrNo(Dated, "BPV")
    INextSno = getNextSno(Dated)
    SnoToUse = INextSno
    Bal = getBalance(cmbAccNo.Tag, Dated)
    dAmount = Format$(LVCheque.SelectedItem.ListSubItems(4).Text, "###")
    sScNo = Left$(LVCheque.SelectedItem.key, Len(LVCheque.SelectedItem.key) - 1)
    
    
    
    sCIH = con.Execute("Select AccNo From VPayChqs Where Sno=" & Val(LVCheque.SelectedItem.key)).Fields(0)
    
       
    
    Bal = Bal + dAmount
    'Start New Transaction
    con.BeginTrans
    'Debit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & cmbAccNo.Tag & _
        "','" & LVCheque.SelectedItem.ListSubItems(3).Text & _
        "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
    
    
    
    Call adjustLedger(cmbAccNo.Tag, Dated, INextSno, Val(dAmount))
    
    Dim ChqNo As String
    INextSno = INextSno + 1
    Bal = getBalance(sCIH, Dated)
    Bal = Bal - dAmount
    'Description
    Dim ChqRec As New ADODB.Recordset
    With ChqRec
        .CursorLocation = adUseClient
        .Open "Select ChequeNo From Cheque Where (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            Desc = "Cheque Paid To " & CmbAccTitle.Tag & " Chq# : " & _
                .Fields(0) & ""
             ChqNo = .Fields(0) & ""
        Else
            Desc = "Cheque Paid To " & CmbAccTitle.Tag & " Chq# : "
        End If
        
        .Close
    End With
    Set ChqRec = Nothing
    'Credit Entry
    
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & sCIH & _
        "','" & Desc & _
        "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
        
    
    
    Call adjustLedger(sCIH, Dated, INextSno, -Val(dAmount))
    
    Dim rsCopy As New ADODB.Recordset
    With rsCopy
        .CursorLocation = adUseClient
        .Open "Select * from ChequeDetail Where (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            con.Execute "Insert into VoucherDetail Values('" & ![BalancedSno] & "','" & SnoToUse & "'," & ![Amount] & ")"
            .MoveNext
        Loop
        .Close
    End With
    Set rsCopy = Nothing
    con.Execute "Update Cheque Set Posted=1 Where Sno=" & sScNo
    con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    'Commit Transaction
    con.CommitTrans
   
    
    
    SavePaymentTrans = True
    Exit Function
err:
    MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
    con.RollbackTrans
        SavePaymentTrans = False
End Function



Private Sub PrintLedger()
With cr1
        

End With
End Sub

Private Sub mnuprintChq_Click()
Call PrintChq(MainForm.cr1, Val(LVCheque.SelectedItem.key))
End Sub

Private Sub mnuRemoveTag_Click()
   con.Execute "Delete From BalanceTags Where Sno=" & Right(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 3)
   LV.ListItems.Remove LV.SelectedItem.Index
End Sub

Private Sub mnuTrans_Click()

    With frmTransaction
        .opts(0) = True
        .Loaded = True
        .cmbRange.ListIndex = 4
        .DtFrom = CDate(LV.SelectedItem)
        .DtTo = .DtFrom
        Call .cmdRefresh_Click
        
        On Error Resume Next
        .LV.ListItems(LV.SelectedItem.key).Selected = True
        .LV.ListItems(LV.SelectedItem.key).EnsureVisible
        
        frmTransaction.Show
        .SetFocus
        .LV.SetFocus
    End With

End Sub



Private Sub RefreshBalanceTags(SnoList As String)
Dim rs As New ADODB.Recordset
Dim idx As Long
On Error GoTo err

If Len(SnoList) = 0 Then
   Exit Sub
Else
   SnoList = Left(SnoList, Len(SnoList) - 1)       'Remove Last ","
End If

With rs
   .CursorLocation = adUseClient
   .Open "Select * From BalanceTags Where Sno in(" & SnoList & ")", con, adOpenForwardOnly, adLockReadOnly
      For i = 0 To .RecordCount - 1
         idx = LV.ListItems(![Sno] & "'").Index + 1
         With LV.ListItems.add(idx, "TAG" & ![Sno])
            .ListSubItems.add
            .ListSubItems.add(, , "*-*-*-*-*-*-*-*-*<  Balance OK  >*-*-*-*-*-*-*-*-*").ForeColor = RGB(0, 125, 0)
         End With
         .MoveNext
      Next
   .Close
End With
Exit Sub

err:
MsgBox "Cannot Retreive Balance Tag InfoDue To Error:" & err.Number & vbNewLine & err.Description, vbInformation
Resume Next
End Sub
