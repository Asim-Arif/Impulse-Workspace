VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmDueRcptsPymts 
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   10905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   10905
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   0
      TabIndex        =   3
      Top             =   60
      Width           =   10860
      Begin VB.CheckBox chkShowPaid 
         Caption         =   "Show Paid Bills"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7815
         TabIndex        =   10
         Top             =   420
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   2
         Left            =   6150
         TabIndex        =   8
         Top             =   345
         Width           =   1590
      End
      Begin MSForms.CommandButton cmdRefresh 
         Default         =   -1  'True
         Height          =   360
         Left            =   9420
         TabIndex        =   4
         Top             =   315
         Width           =   1365
         Caption         =   "Refresh   "
         PicturePosition =   327683
         Size            =   "2408;635"
         Picture         =   "frmDueRcptsPymts.frx":0000
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox CmbAccTitle 
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Top             =   345
         Width           =   3870
         VariousPropertyBits=   545277979
         DisplayStyle    =   3
         Size            =   "6826;556"
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbAccNo 
         Height          =   315
         Left            =   30
         TabIndex        =   6
         Top             =   345
         Width           =   2220
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3916;556"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmDueRcptsPymts.frx":044E
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
         Index           =   1
         Left            =   15
         TabIndex        =   5
         Top             =   120
         Width           =   10560
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   2685
      TabIndex        =   2
      Top             =   810
      Visible         =   0   'False
      Width           =   4320
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   0
      Top             =   0
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   9
      ImageHeight     =   5
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDueRcptsPymts.frx":04F2
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDueRcptsPymts.frx":05D2
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lv 
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   930
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
      Appearance      =   1
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
         Text            =   "Voucher #"
         Object.Width           =   2337
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "Amt"
         Text            =   "Total Amout"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Amount Due"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Pending From"
         Object.Width           =   2293
      EndProperty
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   390
      Left            =   7410
      TabIndex        =   9
      Top             =   6885
      Width           =   1560
      ForeColor       =   64
      Caption         =   "Print         "
      PicturePosition =   327683
      Size            =   "2752;688"
      Picture         =   "frmDueRcptsPymts.frx":06B2
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
      Height          =   390
      Left            =   9030
      TabIndex        =   1
      Top             =   6885
      Width           =   1560
      ForeColor       =   64
      Caption         =   "Close            "
      PicturePosition =   327683
      Size            =   "2752;688"
      Picture         =   "frmDueRcptsPymts.frx":07C4
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
Attribute VB_Name = "frmDueRcptsPymts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReadyToDrag As Boolean
Dim PymtSQL As String, RcptSQL As String
Dim PermissionDenied As Boolean
Dim Pymt As Boolean, dontShowList As Boolean

Private Sub chkShowPaid_Click()
Call cmdRefresh_Click
End Sub

Private Sub cmbAccNo_Change()
    If cmbAccNo.MatchFound Then Call cmbAccNo_Click
End Sub

Private Sub cmbAccNo_Click()
    If CmbAccTitle.ListIndex <> cmbAccNo.ListIndex Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
    If cmbAccNo.MatchFound Then Call cmdRefresh_Click
End Sub

Private Sub CmbAccTitle_Change()
    If CmbAccTitle.MatchFound Then
      Call cmbAccTitle_Click
    End If
      'List1.Visible = True
      
End Sub
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
Private Sub cmbAccTitle_Click()
    If cmbAccNo.ListIndex <> CmbAccTitle.ListIndex Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
End Sub


Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdRefresh_Click()
If Not cmbAccNo.MatchFound Or cmbAccNo = "" Then
   MsgBox "Select An Account First.", vbInformation
   Exit Sub
End If

   Label1(2) = Format(getBalance(cmbAccNo, Date), "0.00 DR;0.00 CR;0.00")
   
   Call GetEntries


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
    
    'PermissionDenied = Not UserHasAccess("DueReceiptsPayments")
    Call FillCmbs

   'Call GetPymts
      'Call GetRcpts
End Sub


Private Sub Form_Resize()

On Error Resume Next



'Me.Move LEFTMARGIN, TOPMARGIN, FORMWIDTH, FORMHEIGHT

cmdClose.Move Me.ScaleWidth - cmdClose.Width - 100, Me.ScaleHeight - cmdClose.Height - 100
cmdPrint.Move cmdClose.Left - cmdPrint.Width - 100, cmdClose.Top

'Frame1.Move (Me.ScaleWidth - Frame1.Width) / 2
FAcc.Move (Me.ScaleWidth - FAcc.Width) / 2, 100 'Frame1.Top + Frame1.Height
LV.Move 50, FAcc.Top + FAcc.Height + 50, Me.ScaleWidth - 100, cmdClose.Top - (FAcc.Top + FAcc.Height + 100)

   
End Sub




Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


If LV.SortKey = ColumnHeader.Index - 1 Then
 Set LV.ColumnHeaderIcons = ImageList1
   If LV.SortOrder = lvwAscending Then
      LV.SortOrder = lvwDescending
      ColumnHeader.Icon = "Up"
   Else
      LV.SortOrder = lvwAscending
      ColumnHeader.Icon = "Down"
   End If
Else
 LV.ColumnHeaderIcons = Nothing
 Set LV.ColumnHeaderIcons = ImageList1
 LV.SortKey = ColumnHeader.Index - 1
 LV.SortOrder = lvwAscending
 ColumnHeader.Icon = "Down"
End If

On Error Resume Next
LV.SelectedItem.EnsureVisible
LV.Sorted = True



End Sub


Private Sub cmdPrint_Click()

    Dim PymtCond As String, RcptCond As String
    If chkShowPaid = vbUnchecked Then
        PymtCond = " and ({AccPay.AmountLeft}>0 or isnull({AccPay.AmountLeft}))"
        RcptCond = " and ({AccRec.AmountLeft}>0 or isnull({AccRec.AmountLeft}))"
    Else
        PymtCond = ""
        RcptCond = ""
    End If

    With cr1
        .ReportFileName = RptPath & "\" & IIf(Pymt, "duePymts.rpt", "DueRcpts.rpt")
        If Pymt Then
            .ReplaceSelectionFormula "{AccPay.Accno}='" & cmbAccNo & "'" & PymtCond
        Else
            .ReplaceSelectionFormula "{AccRec.Accno}='" & cmbAccNo & "'" & RcptCond
        End If
      
        .Formulas(0) = "Company='" & strCompany & "'"
        .Formulas(1) = "address='" & COMPANYADDRESS & "'"
        .Action = 1
    End With
    
End Sub

Private Sub GetEntries()

    Screen.MousePointer = vbHourglass

    Dim TableName As String
    Dim rsList As New ADODB.Recordset
    Dim Sql As String
    Dim total As Double
    rsList.CursorLocation = adUseClient

    If Pymt Then TableName = "VAccpay" Else TableName = "VAccRec"
    Dim Cond As String

    Cond = IIf(chkShowPaid = vbUnchecked, "and ((AmountLeft>0) or (AmountLeft is null))", "")

    Sql = "Select Sno,AccNo,AccTitle,Vdate,VchrNo,Description,TotalAmount,AmountLeft from " & TableName & " Where  Ltrim(AccNo)='" & cmbAccNo & "' AND  (Vdate<='" & Date & "') " & Cond & " Order By SNO"
    rsList.Open Sql, con, adOpenForwardOnly, adLockReadOnly
    ' rsList.Sort = "Sno"
    With LV.ListItems
        .Clear
        'cmd.Caption = "Auto Apply         "
        Dim ITM As ListItem
        Do Until rsList.EOF
            Set ITM = .add(, rsList![Sno] & "?", rsList![VchrNo])
            ITM.Tag = rsList![AccNo] & ""
            'itm.ListSubItems.Add(, , rsList![AccTitle] & "").ToolTipText = rsList![AccTitle] & ""
            ITM.ListSubItems.add(, , rsList![Description] & "").ToolTipText = rsList![Description] & ""
            ITM.SubItems(2) = Format(rsList![TotalAmount], "0.00")
            If IsNull(rsList![AmountLeft]) Then
                ITM.SubItems(3) = Format(rsList![TotalAmount], "0.00")
            Else
                ITM.SubItems(3) = Format(rsList![AmountLeft] & "", "0.00")
            End If
        
            total = total + Val(ITM.SubItems(3))
            If Val(ITM.SubItems(3)) = 0 Then
                ITM.SubItems(4) = "Paid"
            Else
                ITM.SubItems(4) = IIf((Date - rsList![VDate]) = 0, "Today.", Date - rsList![VDate] & " Days.")
            End If
        
            rsList.MoveNext
        Loop
        
        With .add
            .ListSubItems.add
            .ListSubItems.add
            .ListSubItems.add(, , Format(total, "0.00")).Bold = True
        End With
    
        'If .Count > 0 Then cmd.Enabled = True Else cmd.Enabled = False
        
    End With

    rsList.Close
    Set rsList = Nothing
    Screen.MousePointer = vbDefault
End Sub

Public Sub FillCmbs(Optional DuePymts As Boolean)

    Dim TableName As String
   
    Pymt = DuePymts
    TableName = IIf(DuePymts, "VVenders", "VFCustomers")
   
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select AccNo, AccTitle From " & TableName & " Where Acctitle Is Not Null Order By AccNo", con, adOpenForwardOnly, adLockReadOnly
        cmbAccNo.Clear
        CmbAccTitle.Clear
        Do Until .EOF
            cmbAccNo.AddItem !AccNo & ""
            CmbAccTitle.AddItem !AccTitle & ""
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    '   Call AddToCombo(cmbAccNo, "AccNo", TableName, "Order by AccNo")
    '   Call AddToCombo(CmbAccTitle, "AccTitle", TableName, "Order by AccNo")
    '
    If DuePymts Then
        Me.Caption = "Due Payments To Venders"
    Else
        Me.Caption = "Due Payments From Customers"
    End If
    'Lbls(0) = Me.Caption
    'Lbls(1) = Me.Caption
   
End Sub
