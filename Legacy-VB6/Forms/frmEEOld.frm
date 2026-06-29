VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmEE 
   ClientHeight    =   6795
   ClientLeft      =   1110
   ClientTop       =   60
   ClientWidth     =   10365
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10365
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   855
      TabIndex        =   10
      Top             =   2100
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ComboBox cmbFinYear 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   825
      Width           =   2265
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   6735
      Top             =   615
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   375
      Top             =   5835
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
      Top             =   6225
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
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   180
      TabIndex        =   1
      Top             =   -75
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Form EE"
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
         Left            =   4755
         TabIndex        =   2
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Form EE"
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
         Left            =   4770
         TabIndex        =   3
         Top             =   135
         Width           =   1425
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4845
      Left            =   45
      TabIndex        =   0
      Top             =   1215
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   8546
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
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Form E #"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Consignee Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Commodity"
         Object.Width           =   2540
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
         SubItemIndex    =   7
         Text            =   "Shpment"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Neg. Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Bank No."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
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
   Begin MSForms.CommandButton cmdPrintEF 
      Height          =   390
      Left            =   5490
      TabIndex        =   9
      Top             =   6060
      Width           =   1560
      Caption         =   "Print EF         "
      PicturePosition =   327683
      Size            =   "2752;688"
      Picture         =   "frmEE.frx":0C20
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
      Left            =   7080
      TabIndex        =   8
      Top             =   6075
      Width           =   1560
      Caption         =   "Print EE          "
      PicturePosition =   327683
      Size            =   "2752;688"
      Picture         =   "frmEE.frx":0D32
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
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Date Period"
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
      Left            =   840
      TabIndex        =   7
      Top             =   555
      Width           =   2280
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   4440
      TabIndex        =   5
      Top             =   675
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
   Begin MSForms.CommandButton cmdClose 
      Height          =   390
      Left            =   8745
      TabIndex        =   4
      Top             =   6075
      Width           =   1560
      ForeColor       =   0
      Caption         =   "Close            "
      PicturePosition =   327683
      Size            =   "2752;688"
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
Dim SQLString As String
Dim PermissionDenied As Boolean

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
    
    Dim rpt As craxddrt.Report 'New rptEE
    Set rpt = rptApp.OpenReport(RptPath & "\rptEE.rpt") '
    Dim F As New frmPrevRpt
    
    rpt.SQLQueryString = SQLString
    rpt.FormulaFields(3).Text = "'" & cmbFinYear.Text & "'" '(@FinYear)
    rpt.FormulaFields(5).Text = EE '(@EE)
    F.ShowReport " ", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub cmdRefresh_Click()
    
    Call RefreshList
    
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

    Exit Sub
    If LV.ListItems.Count = 0 Then Exit Sub
    Call ResetLV(LV)
    
    With txtEdit
        If LV.SelectedItem.SubItems(9) = "" Then
            .Tag = "New"
        Else
            .Tag = "Edit"
        End If
        .Text = LV.SelectedItem.SubItems(9)
        .Move LV.Left + LV.ColumnHeaders(10).Left, LV.SelectedItem.Top + LV.Top
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    'PermissionDenied = Not UserHasAccess("CustomerList")
    Call SaveLV(LV)
    FillCombo
    
    'Call RefreshList
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub FillCombo()
    
    cmbFinYear.Clear
    For i = 2002 To Year(ServerDate)
        cmbFinYear.AddItem i & " x " & i + 1
    Next i
    
End Sub
Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
    Dim TitleBottom As Long
    Dim OtherHeight As Long
    TitleBottom = fTitle.Top + fTitle.Height
    OtherHeight = cmbFinYear.Height + Label1(2).Height + 150
    
    LV.Move 50, TitleBottom + OtherHeight, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100 + OtherHeight)
 
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
    cmdPrint.Move cmdClose.Left - cmdPrint.Width - 75, cmdClose.Top
    cmdPrintEF.Move cmdPrint.Left - cmdPrintEF.Width - 75, cmdPrint.Top
 
End Sub

Private Sub RefreshList()

    
    On Error GoTo err
    If cmbFinYear.ListIndex = -1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    Dim Cond As String
    Dim StartDate As Date
    Dim EndDate As Date
    StartDate = DateSerial(Left(cmbFinYear.Text, 4), 7, 1)
    EndDate = DateSerial(Right(cmbFinYear.Text, 4), 6, 30)
    Cond = " Where RealisationDate Between '" & StartDate & "' And '" & EndDate & "' Order By RealisationDate"
    SQLString = "Select * From VEE" & Cond
    With rs
    
        .CursorLocation = adUseClient
        .Open SQLString, con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
        
        For i = 0 To .RecordCount - 1
        
            Set itm = LV.ListItems.Add(, ![EntryID] & "'" & ![ManualEntry], LV.ListItems.Count + 1)
            itm.Tag = ![ManualEntry]
            
            itm.ListSubItems.Add , , ![FormE] & ""
            itm.ListSubItems.Add , , ![CustCode] & " " & ![Country]
            itm.ListSubItems.Add , , "SURGICAL INSTRUMENTS"
            itm.ListSubItems.Add , , ![PInvoice] & ""
            itm.ListSubItems.Add , , Format(![DT] & "", "dd-MMM-yyyy")
            itm.ListSubItems.Add , , ![Curr] & " " & Val(![TotalCustomAmt] & "")
            itm.ListSubItems.Add , , ""
            itm.ListSubItems.Add , , ""
            itm.ListSubItems.Add , , ![BankNo] & ""
            
            itm.ListSubItems.Add , , ![Curr] & " " & Val(![FAmount] & "")
            itm.ListSubItems.Add , , "Rs. " & Val(![PAmount] & "")
            itm.ListSubItems.Add , , Format(![RealisationDate] & "", "dd-MMM-yyyy")
            
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
 
 
 
Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtEdit.Visible = False
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
'        If txtEdit.Text = "" Then
'            txtEdit.Visible = False
'            Exit Sub
'        End If
        
        Call StartTrans(con)
        If txtEdit.Tag = "New" Then
            con.Execute "Insert Into FormEEBankNos(Manual,RefID" & _
            ",BankNo) Values(" & LV.SelectedItem.Tag & "," & _
            Left(LV.SelectedItem.Key, Len(LV.SelectedItem.Key) - 2) & _
            ",'" & txtEdit.Text & "')"
        Else
            con.Execute "Update FormEEBankNos Set BankNo='" & _
            txtEdit.Text & "' Where Manual=" & LV.SelectedItem.Tag & _
            " And RefID=" & Left(LV.SelectedItem.Key, Len(LV.SelectedItem.Key) - 2)
        End If
        con.CommitTrans
        
        LV.SelectedItem.SubItems(9) = txtEdit.Text
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        
        txtEdit.Visible = False
    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
