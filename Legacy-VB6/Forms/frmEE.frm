VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEE 
   ClientHeight    =   8715
   ClientLeft      =   1110
   ClientTop       =   60
   ClientWidth     =   12870
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   12870
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtEdit 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2355
      TabIndex        =   10
      Top             =   4125
      Visible         =   0   'False
      Width           =   650
   End
   Begin MSComCtl2.DTPicker DTRealisation 
      Height          =   300
      Left            =   3795
      TabIndex        =   9
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   100532227
      CurrentDate     =   41121
   End
   Begin VB.ComboBox cmbFinYear 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   45
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   330
      Width           =   2265
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   375
      Top             =   5325
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
            Picture         =   "frmEE.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEE.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   510
      Top             =   5715
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEE.frx":0378
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEE.frx":07CC
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   7485
      Left            =   45
      TabIndex        =   0
      Top             =   705
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   13203
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImageList2"
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sr No."
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Form E #"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Consignee Name"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Commodity"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Inv No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "PageNoForEE"
         Text            =   "Page #"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "ItemNoForEE"
         Text            =   "Item #"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "BankNoForEE"
         Text            =   "Bank #"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "AmountEE"
         Text            =   "Amt Foreign"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Amount Pak"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Rel. Date"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSForms.ComboBox cmbCompany 
      Height          =   315
      Left            =   2310
      TabIndex        =   8
      Top             =   330
      Width           =   4365
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "7699;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbBank 
      Height          =   315
      Left            =   6660
      TabIndex        =   7
      Top             =   330
      Width           =   4530
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "7990;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdPrintEF 
      Height          =   390
      Left            =   7980
      TabIndex        =   6
      Top             =   8235
      Width           =   1560
      Caption         =   "Print EF         "
      PicturePosition =   327683
      Size            =   "2752;688"
      Accelerator     =   70
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   390
      Left            =   9600
      TabIndex        =   5
      Top             =   8235
      Width           =   1560
      Caption         =   "Print EE          "
      PicturePosition =   327683
      Size            =   "2752;688"
      Accelerator     =   69
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
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Date Period                            Company                                                                               Bank"
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
      TabIndex        =   4
      Top             =   75
      Width           =   11145
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   11535
      TabIndex        =   2
      Top             =   315
      Width           =   1215
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2143;635"
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
      Height          =   390
      Left            =   11220
      TabIndex        =   1
      Top             =   8235
      Width           =   1560
      ForeColor       =   0
      Caption         =   "Close            "
      PicturePosition =   327683
      Size            =   "2752;688"
      Picture         =   "frmEE.frx":0C20
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
Attribute VB_Name = "frmEE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLString As String, strSQLStringrpt As String
Dim PermissionDenied As Boolean
Dim iColNo As Integer

Private Sub cmbCompany_Change()
    Call cmbCompany_Click
End Sub

Private Sub cmbCompany_Click()
    
    If cmbCompany.MatchFound = False Then
        cmbBank.Clear
    Else
        AddToCombo cmbBank, "Bank + ' {' + Branch + '}'", "BankList", " Where CompanyRefID=" & cmbCompany.List(cmbCompany.ListIndex, 1), , "Bank,Branch,BankID"
    End If
    
End Sub

Private Sub cmbFinYear_Change()
    Call cmbFinYear_Click
End Sub

Private Sub cmbFinYear_Click()
    Call RefreshList
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Call PrintReport(True)
End Sub

Private Sub cmdPrintEF_Click()
    Call PrintReport(False)
End Sub

Private Sub PrintReport(EE As Boolean)

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Dim f As New frmPrevRpt
    Set rpt = rptApp.OpenReport(RptPath & "\rptEE.rpt")
    
'    On Error Resume Next
'    SrcReport.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
'    rpt.SQLQueryString = SQLString
    rpt.FormulaFields(3).Text = "'" & cmbFinYear.Text & "'" '(@FinYear)
    rpt.FormulaFields(5).Text = EE '(@EE)
    rpt.FormulaFields(10).Text = "'" & cmbBank.List(cmbBank.ListIndex, 1) & " " & cmbBank.List(cmbBank.ListIndex, 2) & "'"
    'rpt.RecordSortFields.add rpt.Database.Tables(1).Fields(9), crDescendingOrder
    Dim rptSUB As CRAXDDRT.Report
    If EE Then
        Set rptSUB = rpt.OpenSubreport("EESub")
    Else
        Set rptSUB = rpt.OpenSubreport("EFSub")
    End If
    rptSUB.SQLQueryString = strSQLStringrpt
    
    f.ShowReport "", rpt, strSQLStringrpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshList
End Sub

Private Sub DTRealisation_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        con.Execute "UPDATE FCustPayments SET DT='" & DTRealisation.Value & "' WHERE EntryID=" & Val(LV.SelectedItem.Tag)
        LV.SelectedItem.SubItems(11) = Format(DTRealisation, "dd-MMM-yyyy")
        DTRealisation.Visible = False
    ElseIf KeyCode = 27 Then
        DTRealisation.Visible = False
    End If
    
End Sub

Private Sub DTRealisation_LostFocus()
    DTRealisation.Visible = False
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

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    LV.Sorted = True
    If LV.SortKey = ColumnHeader.Index - 1 Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Up"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Down"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If

End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If Not (iColNo = 8 Or iColNo = 9 Or iColNo = 10 Or iColNo = 11) Then iColNo = 13
        
    If iColNo = 13 Then
        With DTRealisation
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top
            .Value = LV.SelectedItem.SubItems(iColNo - 1)
            .Visible = True
            .SetFocus
        End With
    Else
        With txtEdit
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
            .Text = LV.SelectedItem.SubItems(iColNo - 1)
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    End If
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    ''PermissionDenied = Not UserHasAccess("CustomerList")

    FillCombo
    
    'Call RefreshList
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub FillCombo()
    
    cmbFinYear.Clear
    For i = 2007 To year(Date)
        cmbFinYear.AddItem i & " x " & i + 1
    Next i
   
    AddToCombo cmbCompany, "CompanyName", "Companies", " Order By CompanyName", , "EntryID"
    
    
End Sub

Private Sub RefreshList()

    
    On Error GoTo err
    If cmbFinYear.ListIndex = -1 Then Exit Sub
    If cmbBank.ListIndex = -1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Cond As String
    Dim StartDate As Date
    Dim EndDate As Date
    StartDate = DateSerial(Left(cmbFinYear.Text, 4), 7, 1)
    EndDate = DateSerial(Right(cmbFinYear.Text, 4), 6, 30)
    Cond = " WHERE RealisationDate BETWEEN '" & StartDate & "' AND '" & EndDate & "' AND OurBankID=" & Val(cmbBank.List(cmbBank.ListIndex, 3) & "") & " AND FDBCNo<>'' AND CompanyRefID=" & Val(cmbCompany.List(cmbCompany.ListIndex, 1) & "")
    
    SQLString = "SELECT * FROM VEE" & Cond
    strSQLStringrpt = SQLString & " AND HideEntry=0"
    
    Dim lForeColor As Long
    With rs
    
        .CursorLocation = adUseClient
        .Open SQLString, con, adOpenForwardOnly, adLockReadOnly
        SQLString = SQLString & " Order By RealisationDate"
        LV.ListItems.Clear
        
        For i = 0 To .RecordCount - 1
        
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , ![FormE] & ""
            ITM.ListSubItems.add , , ![CustCode] & " " & ![Country]
            ITM.ListSubItems.add , , !InvoiceOf & ""
            ITM.ListSubItems.add , , ![CustomInvoice] & ""
            ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
            ITM.ListSubItems.add , , ![Curr] & " " & Val(![TotalCommercialAmt] & "")
            
            ITM.ListSubItems.add , , !PageNoForEE & ""
            ITM.ListSubItems.add , , !ItemNoForEE & ""
            ITM.ListSubItems.add , , !BankNoForEE & ""
            
            ITM.ListSubItems.add , , Val(![FAmount] & "")
            ITM.ListSubItems.add(, , Val(![PAmount] & "")).Tag = Val(!ExchRate & "")
            ITM.ListSubItems.add(, , Format(![RealisationDate] & "", "dd-MMM-yyyy")).Tag = !HideEntry
            
            If !HideEntry Then
                lForeColor = vbRed
            Else
                lForeColor = LV.ForeColor
            End If
            
            ITM.ForeColor = lForeColor
            For j = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(j).ForeColor = lForeColor
            Next
            
            .MoveNext
        Next
        .Close
    End With
    Set rs = Nothing
        
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    MsgBox err.Description
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
    If LV.ListItems.count = 0 Then Exit Sub
     
    Dim strRet As String, iValue As Integer, lForeColor As Long, bTagValue As Boolean
    strRet = ""
    If CBool(LV.SelectedItem.ListSubItems(12).Tag) Then
        strRet = ShowPopUpMenu1("Show This Entry")
        iValue = 0
        lForeColor = LV.ForeColor
        
    Else
        strRet = ShowPopUpMenu1("Hide This Entry")
        iValue = 1
        lForeColor = vbRed
    End If
         
    If strRet <> "" Then
    
        con.Execute "UPDATE FCustPayments SET HideEntry=" & iValue & " WHERE EntryID=" & Val(LV.SelectedItem.Tag)
        Dim i As Integer
        LV.SelectedItem.ForeColor = lForeColor
        LV.SelectedItem.ListSubItems(12).Tag = CBool(iValue)
        For i = 1 To LV.SelectedItem.ListSubItems.count
            LV.SelectedItem.ListSubItems(i).ForeColor = lForeColor
        Next
        
    End If
     
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        
        
        
        If iColNo = 11 Then
            con.Execute "UPDATE FCustPayments SET AmountEE=" & Val(txtEdit.Text) & " WHERE EntryID=" & Val(LV.SelectedItem.Tag)
        Else
            con.Execute "UPDATE FCustPayments SET " & LV.ColumnHeaders(iColNo).Tag & "='" & txtEdit.Text & "' WHERE EntryID=" & Val(LV.SelectedItem.Tag)
        End If
        
        LV.SelectedItem.SubItems(iColNo - 1) = txtEdit
        If iColNo < 10 Then
            iColNo = iColNo + 1
            Call LV_DblClick
            DoEvents
            Exit Sub
        Else
            If LV.SelectedItem.Index < LV.ListItems.count Then
                LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                LV.SelectedItem.EnsureVisible
                iColNo = 8
                Call LV_DblClick
                DoEvents
                Exit Sub
            Else
                txtEdit.Visible = False
            End If
        End If
        
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub
