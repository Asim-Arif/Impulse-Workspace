VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmChqBooks 
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   360
   ClientWidth     =   10620
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   10620
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSearch 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   420
      Left            =   5550
      TabIndex        =   13
      Tag             =   "Type here to Search"
      Text            =   "Type here to Search"
      Top             =   30
      Width           =   5085
   End
   Begin MSComCtl2.DTPicker DTClearance 
      Height          =   285
      Left            =   7830
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yy"
      Format          =   126550019
      CurrentDate     =   39974
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   1635
      Top             =   1080
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
   Begin VB.Frame FVDrag 
      BorderStyle     =   0  'None
      Height          =   3390
      Left            =   5520
      MousePointer    =   9  'Size W E
      TabIndex        =   10
      Top             =   50
      Width           =   50
   End
   Begin VB.Frame FHDrag 
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   -225
      MousePointer    =   7  'Size N S
      TabIndex        =   9
      Top             =   3405
      Width           =   11805
   End
   Begin VB.PictureBox PicEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5610
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   8
      Top             =   885
      Visible         =   0   'False
      Width           =   4845
      Begin VB.TextBox txtChqBookDetail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "RptTime"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   3420
         MaxLength       =   10
         TabIndex        =   4
         Top             =   -15
         Width           =   1410
      End
      Begin VB.TextBox txtManualNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "RptTime"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   2025
         MaxLength       =   10
         TabIndex        =   3
         Top             =   -15
         Width           =   1410
      End
      Begin VB.ComboBox cmbChqs 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmChqBooks.frx":0000
         Left            =   1155
         List            =   "frmChqBooks.frx":000D
         TabIndex        =   2
         Text            =   "25"
         Top             =   -30
         Width           =   885
      End
      Begin VB.TextBox txtStartingfrom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "RptTime"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   0
         MaxLength       =   10
         TabIndex        =   1
         Top             =   -15
         Width           =   1305
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2970
      Left            =   5550
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   465
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   5239
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
      BorderStyle     =   1
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
         Text            =   "Starting From"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Chqs"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Manual #"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Bank CHq"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.ListView LVChqs 
      Height          =   3000
      Left            =   15
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3540
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   5292
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
      BorderStyle     =   1
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Chq No"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cheque Type"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Issue Date"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Chq Date"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Account Title"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Description"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Amount"
         Object.Width           =   2073
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Payee"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Clearance"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Handed Over To"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   2580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChqBooks.frx":001E
            Key             =   "BNK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChqBooks.frx":0472
            Key             =   "CB"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChqBooks.frx":08C6
            Key             =   "ISS"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChqBooks.frx":0D1A
            Key             =   "BNC"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChqBooks.frx":116E
            Key             =   "CHQ"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChqBooks.frx":15C2
            Key             =   "ACC"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChqBooks.frx":1A16
            Key             =   "ACCOD"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChqBooks.frx":1E68
            Key             =   "BCH"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChqBooks.frx":22BC
            Key             =   "CAN"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tv 
      Height          =   3390
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5980
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   7815
      TabIndex        =   11
      Top             =   6645
      Width           =   1320
      ForeColor       =   64
      VariousPropertyBits=   25
      Caption         =   "Print         "
      PicturePosition =   327683
      Size            =   "2328;635"
      Picture         =   "frmChqBooks.frx":2714
      Accelerator     =   80
      FontName        =   "Tahoma"
      FontEffects     =   1073750017
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   9255
      TabIndex        =   7
      Top             =   6645
      Width           =   1320
      ForeColor       =   64
      Caption         =   "Close     "
      PicturePosition =   327683
      Size            =   "2328;635"
      Picture         =   "frmChqBooks.frx":2826
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnuPop 
      Caption         =   "mnuPop"
      Begin VB.Menu mnuNewChqBook 
         Caption         =   "New Cheque &Book"
      End
      Begin VB.Menu mnuPrintChqBook 
         Caption         =   "Print This Chqbook's Detail"
      End
      Begin VB.Menu mnuDeleteChqBook 
         Caption         =   "Delete this Cheque Book"
      End
   End
   Begin VB.Menu mnuCancel 
      Caption         =   "mnucancel"
      Begin VB.Menu mnuCancelChq 
         Caption         =   "Cancel This Cheque"
      End
      Begin VB.Menu mnuPrintChq 
         Caption         =   "Print This Cheque"
      End
      Begin VB.Menu mnuViewInTransactionRegister 
         Caption         =   "View in Transaction Register"
      End
   End
   Begin VB.Menu mnuBank 
      Caption         =   "mnuBank"
      Begin VB.Menu mnuNewBankAcc 
         Caption         =   "&New Bank Account"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDefault 
         Caption         =   "Set As Default Bank Account"
      End
      Begin VB.Menu mnuOD 
         Caption         =   "Allow Over Draft For This Account"
      End
      Begin VB.Menu mnuPrintChqBooks 
         Caption         =   "Print Chq Books"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLedger 
         Caption         =   "View Ledger For Last 30 Days"
      End
   End
End
Attribute VB_Name = "frmChqBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReadyToDrag As Boolean
Dim SecondTime As Boolean

Dim PermissionDenied As Boolean


Private Sub cmbChqs_KeyPress(KeyAscii As Integer)
 
    If KeyAscii = 13 Then
        If Val(cmbChqs) > 0 Then
            txtManualNo.SetFocus
        Else
            MsgBox "Number of Cheques Cannot Be Zero.", vbInformation
            cmbChqs.SetFocus
        End If
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
        LV.SetFocus
    Else
        KeyAscii = OnlyNumber(KeyAscii)
    End If
    
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub cmdPrint_Click()

On Error GoTo err

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ChqBook.rpt")
    rpt.FormulaFields.GetItemByName("Company").Text = "'" & strCompany & "'"
    rpt.FormulaFields.GetItemByName("Address").Text = "'" & COMPANYADDRESS & "'"
    Dim f As New frmPrevRpt
    f.ShowReport "{VChqBookLedger.ChqbookNo}=" & Val(LV.SelectedItem.key), rpt
   
Exit Sub

err:
    MsgBox err.Description
End Sub

Private Sub DTClearance_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        con.Execute "UPDATE Cheque SET ClearanceDT='" & DTClearance.Value & "' WHERE SNo=" & Val(LVChqs.SelectedItem.key)
        LVChqs.SelectedItem.SubItems(8) = Format(DTClearance, "dd-MMM-yy")
        DTClearance.Visible = False
    ElseIf KeyCode = 27 Then
        DTClearance.Visible = False
    End If
    
End Sub

Private Sub DTClearance_LostFocus()
    DTClearance.Visible = False
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

    'PermissionDenied = Not UserHasAccess("BanksChqbooks")
        
    If strCompany = "Banzai" Then
        LV.ColumnHeaders(3).Text = "BNZ Chq Book #"
        LV.ColumnHeaders(4).Text = "UBL Chq Book #"
    Else
        LV.ColumnHeaders(3).Text = "Manual Chq Book #"
        LV.ColumnHeaders(4).Text = "Bank Chq Book #"
    End If
    
    Call RefreshTV
    Set Tv.ImageList = ImageList1
    Set LV.SmallIcons = ImageList1
    Set LVChqs.SmallIcons = ImageList1
    
    mnuPop.Visible = False
    mnuCancel.Visible = False
    mnuBank.Visible = False

End Sub
Private Sub RefreshTV()
Screen.MousePointer = vbHourglass
Dim DefaultIndex As Integer

Set Tv.ImageList = ImageList1
Tv.Nodes.Clear

Dim rsBnk As New ADODB.Recordset
Dim rsBch As New ADODB.Recordset
Dim rsAcc As New ADODB.Recordset

rsBnk.Open "Select Distinct Bank From VBankAccounts Order By Bank", con, adOpenForwardOnly

For i = 0 To rsBnk.RecordCount - 1
   Tv.Nodes.add , , rsBnk![Bank] & "", rsBnk![Bank] & "", "BNK"
      
      If rsBch.State = 1 Then rsBch.Close
      rsBch.Open "Select Distinct BankID,Branch From VBankAccounts Where Bank='" & rsBnk![Bank] & "'", con, adOpenForwardOnly
      
      For k = 0 To rsBch.RecordCount - 1
         Tv.Nodes.add rsBnk![Bank] & "", tvwChild, rsBch![BankID] & "BCH" & k, rsBch![Branch] & "", "BCH"
                  
                  
               If rsAcc.State = 1 Then rsAcc.Close
                rsAcc.Open "Select * From VBankAccounts Where BankID=" & rsBch![BankID], con, adOpenForwardOnly
                  For l = 0 To rsAcc.RecordCount - 1
                  
                     Tv.Nodes.add rsBch![BankID] & "BCH" & k, tvwChild, rsAcc![AccNo], "[" & rsAcc![AccNo] & "] " & rsAcc![AccTitle], IIf(rsAcc![AllowOD], "ACCOD", "ACC")
                     
                     If rsAcc![DefaultBank] Then DefaultIndex = Tv.Nodes.count
                      
                     rsAcc.MoveNEXT
                  Next
                  
         rsBch.MoveNEXT
      Next
      
      

   rsBnk.MoveNEXT
Next

If DefaultIndex > 0 Then
   Tv.Nodes(DefaultIndex).Bold = True
   Tv.Nodes(DefaultIndex).Selected = True
   Call Tv_NodeClick(Tv.Nodes(DefaultIndex))
End If

Set rsBnk = Nothing
Set rsCty = Nothing
Set rsAcc = Nothing
Screen.MousePointer = vbDefault
End Sub







Private Sub Form_Resize()
   
   On Error Resume Next
   Tv.Move 50, Me.ScaleTop + 50, Me.ScaleWidth / 2, Me.ScaleHeight / 2
   LV.Move Tv.Left + Tv.Width + 50, LV.Top, Me.ScaleWidth / 2, Tv.Height - txtSearch.Height
   txtSearch.Move LV.Left, Tv.Top, LV.Width
   
   FHDrag.Move 0, LV.Top + LV.Height, Me.ScaleWidth
   cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 100
   cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
   
   FVDrag.Move LV.Left - 50, Tv.Top, 50, Tv.Height
   LVChqs.Move Tv.Left, Tv.Height + 100, Me.ScaleWidth - 100, cmdClose.Top - FHDrag.Top - 150
   
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Screen.MousePointer = vbHourglass

    If SecondTime And LV.ListItems.count < 2 Then
        SecondTime = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Dim strCondition As String
    strCondition = " WHERE ChqBookNo=" & Val(Item.key) & ""
    Call RefreshChqs(strCondition)
    
    SecondTime = True
    cmdPrint.Enabled = (LVChqs.ListItems.count > 0)
    Screen.MousePointer = vbDefault

End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If OutOfFinancialYear Then Exit Sub

mnuPrintChqBook.Visible = LV.ListItems.count > 0
If Button = 2 Then Me.PopUpMenu mnuPop

End Sub



Private Sub LVChqs_DblClick()
    If LVChqs.ListItems.count = 0 Then Exit Sub
    DTClearance.Move LVChqs.Left + LVChqs.ColumnHeaders(9).Left, LVChqs.Top + LVChqs.SelectedItem.Top
    
    If LVChqs.SelectedItem.SubItems(8) = "-" Then
        DTClearance.Value = Date
    Else
        DTClearance.Value = LVChqs.SelectedItem.SubItems(8)
    End If
    DTClearance.Visible = True
    DTClearance.SetFocus
End Sub

Private Sub LVChqs_ItemClick(ByVal Item As MSComctlLib.ListItem)
   LVChqs.ToolTipText = Item.SubItems(6)
End Sub

Private Sub LVChqs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Or LVChqs.ListItems.count = 0 Or OutOfFinancialYear Then Exit Sub
    
    If UCase(LVChqs.SelectedItem.SubItems(5)) = "< NOT ISSUED >" Then
        mnuCancelChq.Tag = "CANCEL"
        mnuCancelChq.Visible = True
        mnuPrintChq.Visible = False
        mnuCancelChq.Caption = "Cancel This Cheque"
        Me.PopUpMenu mnuCancel
    ElseIf Trim(UCase(LVChqs.SelectedItem.SubItems(1))) = "CANCELED" Then
        mnuCancelChq.Visible = True
        mnuCancelChq.Caption = "Make This Cheque Uncanceled"
        mnuCancelChq.Tag = "UNDOCANCEL"
        mnuPrintChq.Visible = False
        Me.PopUpMenu mnuCancel
    ElseIf Val(LVChqs.SelectedItem.SubItems(6)) > 0 Then
        mnuPrintChq.Visible = True
        mnuCancelChq.Visible = False
        Me.PopUpMenu mnuCancel
    End If
    
End Sub

Private Sub mnuCancelChq_Click()

    Select Case mnuCancelChq.Tag
        Case "CANCEL"
            If MsgBox("Are You Sure To Cancel Cheque # :" & LVChqs.SelectedItem, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            Dim Desc As String, ChqNo As String
            Desc = InputBox("Enter Description For Cheque To Be Canceled", "Cancel Cheque")
            If Trim(Desc) = "" Then
                MsgBox "Invalid Description, Cheque is Not Canceled", vbInformation
                Exit Sub
            End If
    
          
            ChqNo = LVChqs.SelectedItem
    
            INextSno = getNextSno(Date, "Cheque")
          
            con.Execute "INSERT INTO Cheque(SNo,CDate,AccNo,Description,Amount,BankID,chequeno,ChqBookNo,chequeType,chequeDate,Posted,Payment,Bounced) Values(" & INextSno & ",'" & Date & "','','" & Desc & "',0,0,'" & _
            ChqNo & "'," & Val(LV.SelectedItem.key) & ",'Canceled','" & Date & "',1,1,0)"
        
            'Remove the Cheque From Cheque book
            Call IssueChq(Val(LV.SelectedItem.key), ChqNo)
        
            'Refresh The List
            Call LV_ItemClick(LV.SelectedItem)

        
        
Case "UNDOCANCEL"
        If MsgBox("Are You Sure To Make Cheque # :" & LVChqs.SelectedItem & " Uncanceled ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        con.Execute "Delete From Cheque Where Sno=" & Val(LVChqs.SelectedItem.key)
        con.Execute "Update ChqList Set issued=0 Where ChqNo='" & LVChqs.SelectedItem & "'"
        'Refresh The List
        SecondTime = False
        Call LV_ItemClick(LV.SelectedItem)
End Select

        
End Sub

Private Sub mnuDefault_Click()
    con.Execute "Update BankAccounts Set [DefaultBank]=0"
    con.Execute "Update BankAccounts Set [DefaultBank]=1 Where AccNo='" & Tv.SelectedItem.key & "'"
    Call RefreshTV
End Sub

Private Sub mnuDeleteChqBook_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If Val(LV.SelectedItem.ListSubItems(2).Tag) > 0 Then
        MsgBox "Can't Delete. Cheque(s) issued.", vbInformation
        Exit Sub
    End If
    If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion) = vbYes Then
        con.Execute "DELETE FROM ChqBooks WHERE ChqBookNo=" & Val(LV.SelectedItem.key)
                
        LV.ListItems.Remove LV.SelectedItem.Index
        LVChqs.ListItems.Clear
    End If
    
        
End Sub

Private Sub mnuLedger_Click()
    With frmLedger
        .Show
        .cmbAccNo = GetID(Tv.SelectedItem)
        Call .cmdRefresh_Click
    End With
End Sub

Private Sub mnuNewBankAcc_Click()
 Dim f As New frmNewAccount
 With f
 .chkBankAcc.Value = vbChecked
 Call FindInCombo(.cmbBank, Tv.SelectedItem.Parent.Parent)
 Call FindInCombo(.cmbBankBranch, Tv.SelectedItem.Parent)
 .Show 1
 End With
 Call RefreshTV
 
End Sub

Private Sub mnuNewChqBook_Click()

    If HasRunnigChqBook(GetID(Tv.SelectedItem)) Then
        If MsgBox("This Account Has Already an Incomplete Cheque Book." & vbNewLine & "Do You Want To Enter Another.", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If
    
    If LV.ListItems.count > 0 Then
        With LV.ListItems(LV.ListItems.count)
            PicEdit.Move LV.Left + 30 + .Left, .Top + LV.Top + 250
        End With
    Else
        PicEdit.Move LV.Left + 30, LV.Top + 250
    End If
    PicEdit.Visible = True
    txtStartingfrom.SetFocus

End Sub

Private Sub RefreshLV(AccNo As String)

    Screen.MousePointer = vbHourglass
    Set LV.SmallIcons = ImageList1
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM VChqBooks WHERE AccNo='" & AccNo & "'", con, adOpenForwardOnly
    Dim ITM As ListItem
    LV.ListItems.Clear
    For i = 0 To rs.RecordCount - 1
        Set ITM = LV.ListItems.add(, rs![ChqBookNo] & "'" & i, rs![StartingFrom], , "CB")
        ITM.ListSubItems.add , , rs![Chqs]
        ITM.ListSubItems.add , , rs!ManualNo & ""
        ITM.ListSubItems.add , , rs!ChqBookDetail & ""
        If Val(rs![ChqsLeft] & "") = 0 Then
            ITM.ListSubItems.add , , "Used Completely"
        Else
            ITM.ListSubItems.add , , Val(rs![ChqsLeft] & "") & " Chqs Left."
        End If
        ITM.ListSubItems(2).Tag = Val(rs!Chqs & "") - Val(rs!ChqsLeft & "")         'Used Chqs
        rs.MoveNEXT
    Next
    
    rs.Close
    Set rs = Nothing
    
    If LV.ListItems.count > 0 Then Call LV_ItemClick(LV.SelectedItem)
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuOD_Click()
    Dim f As New frmBankOD
    f.ShowMe Tv.SelectedItem.key
    Call RefreshTV
End Sub

Private Sub mnuprintChq_Click()

    Dim dChqSNo As Double
    dChqSNo = Val(LVChqs.SelectedItem.key)
    
    Dim f As New frmPrintCheque
    If dChqSNo <> 0 Then
        Call f.ShowMe(dChqSNo)
    Else
        MsgBox "Cheque not Issued yet", vbInformation
    End If
    
End Sub

Private Sub mnuPrintChqbook_Click()

    On Error GoTo err

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ChqBook.rpt")
    rpt.FormulaFields.GetItemByName("Company").Text = "'" & strCompany & "'"
    rpt.FormulaFields.GetItemByName("Address").Text = "'" & COMPANYADDRESS & "'"
    Dim f As New frmPrevRpt
    f.ShowReport "{VChqBookLedger.ChqbookNo}=" & Val(LV.SelectedItem.key), rpt
   
Exit Sub

err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintChqBooks_Click()

    On Error GoTo err

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ChqBooksSummary.rpt")
    rpt.FormulaFields.GetItemByName("Company").Text = "'" & strCompany & "'"
    rpt.FormulaFields.GetItemByName("Address").Text = "'" & COMPANYADDRESS & "'"
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt
   
Exit Sub

err:
    MsgBox err.Description

End Sub

Private Sub mnuViewInTransactionRegister_Click()

    If LVChqs.ListItems.count = 0 Then Exit Sub
    If LVChqs.SelectedItem.ListSubItems(4).Tag = "" Then Exit Sub
    
    Dim myDT As Date, dSNo As Double
    myDT = LVChqs.SelectedItem.ListSubItems(4).Tag
    dSNo = Val(LVChqs.SelectedItem.ListSubItems(5).Tag)
    
    Dim f As New frmTransaction
    Load f
    Call f.ShowMe(myDT, dSNo)
    
End Sub

Private Sub Tv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Or Tv.Nodes.count = 0 Then Exit Sub
      
    mnuNewBankAcc.Enabled = Not OutOfFinancialYear
      
    mnuDefault.Enabled = Tv.SelectedItem.Image Like "ACC*" And Not OutOfFinancialYear
    mnuOD.Enabled = mnuDefault.Enabled And Not OutOfFinancialYear
    
    If Tv.SelectedItem.Image = "ACCOD" Then
        mnuOD.Caption = "Disallow Over Draft For This Account"
        mnuOD.Tag = "OD"
        Me.PopUpMenu mnuBank
    Else
        Tv.SelectedItem.Image = "ACC"
        mnuOD.Caption = "Allow Over Draft For This Account"
        mnuOD.Tag = ""
        Me.PopUpMenu mnuBank
    End If
    
End Sub

Private Sub Tv_NodeClick(ByVal Node As MSComctlLib.Node)

    With Node
        If .Image Like "ACC*" Then
            Call RefreshLV(.key)
        Else
            Call RefreshLV("")
        End If
    End With

    If LV.ListItems.count > 0 Then
        Call LV_ItemClick(LV.ListItems(1))
    Else
        LVChqs.ListItems.Clear
    End If

End Sub

Private Sub txtChqBookDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call SaveChqBook(Tv.SelectedItem.key, Val(txtStartingfrom), Val(cmbChqs), Len(txtStartingfrom))
        PicEdit.Visible = False
        Call RefreshLV(Tv.SelectedItem.key)
        LV.SetFocus
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
End Sub

Private Sub txtManualNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtChqBookDetail.SetFocus
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtSearch_Change()

    Dim strCondition As String
    
    
    strCondition = " WHERE BankAccNo='" & Tv.SelectedItem.key & "'"
    
    Dim strFind As String
    
    If txtSearch <> txtSearch.Tag And txtSearch <> "" Then
        strFind = "%" & txtSearch & "%"
        strCondition = strCondition & " AND (AccTitle LIKE '" & strFind & "' OR Description LIKE '" & strFind & "' OR Payee LIKE '" & strFind & "' OR ChqNo LIKE '" & strFind & "' OR Amount LIKE '" & strFind & "')"
    End If
    Call RefreshChqs(strCondition)
    
End Sub

Private Sub txtStartingfrom_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
      cmbChqs.SetFocus
   ElseIf KeyAscii = 27 Then
      PicEdit.Visible = False
      LV.SetFocus
   Else
      KeyAscii = OnlyNumber(KeyAscii)
   End If
End Sub



Private Sub SaveChqBook(AccNo, StartingFrom As Long, Chqs As Long, Optional ChqNoLen As Long)

    Dim rs As New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    With rs
        .Open "SELECT * FROM ChqBooks WHERE AccNo='fake'", con, adOpenDynamic, adLockOptimistic
        .AddNew
        ![AccNo] = AccNo
        ![StartingFrom] = Format(StartingFrom, String(ChqNoLen, "0"))
        ![Chqs] = Chqs
        !ManualNo = txtManualNo.Text
        !ChqBookDetail = txtChqBookDetail.Text
        .Update
   
        Do Until Val(![ChqBookNo] & "") > 0
            .Requery
        Loop
 
        For i = StartingFrom To (StartingFrom + Chqs) - 1
            con.Execute "INSERT INTO ChqList VALUES(" & ![ChqBookNo] & ",'" & Format(i, String(ChqNoLen, "0")) & "',0)"
        Next
   
    End With

    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault

End Sub


Private Sub FHDrag_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReadyToDrag = True
    FHDrag.BackColor = &H808080
    FHDrag.ZOrder
End Sub

Private Sub FHDrag_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim TheTop As Long
TheTop = FHDrag.Top + y

    If ReadyToDrag Then
        If (TheTop > 500) And (TheTop < cmdClose.Top - 500) Then
            FHDrag.Top = TheTop
        End If
    End If

    
End Sub

Private Sub FHDrag_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If ReadyToDrag Then
    
        FHDrag.BackColor = Me.BackColor
        Tv.Height = FHDrag.Top - Tv.Top
        LV.Height = Tv.Height
        FVDrag.Height = LV.Height
        
        LVChqs.Move 50, FHDrag.Top + FHDrag.Height, LVChqs.Width, Me.ScaleHeight - (cmdClose.Height + LV.Height + FHDrag.Height + 200)
        If LVChqs.Top > cmdClose.Top - 500 Then
            LVChqs.Top = cmdClose.Top - 500
        End If
        
        ReadyToDrag = False
        
        FHDrag.ZOrder 1
    End If
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''


Private Sub fvdrag_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReadyToDrag = True
    FVDrag.BackColor = &H808080
    FVDrag.ZOrder
End Sub

Private Sub fvdrag_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim TheLeft As Long
TheLeft = FVDrag.Left + x

    If ReadyToDrag Then
        If (TheLeft > 500) And (TheLeft < Me.ScaleWidth - 500) Then
            FVDrag.Left = TheLeft
        End If
    End If

    
End Sub

Private Sub fvdrag_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If ReadyToDrag Then
    
        FVDrag.BackColor = Me.BackColor
        Tv.Move Tv.Left, Tv.Top, FVDrag.Left - Tv.Left
        LV.Move FVDrag.Left + FVDrag.Width, LV.Top, Me.ScaleWidth - 50
             
        
        
        ReadyToDrag = False
        
        FVDrag.ZOrder 1
    End If
    
End Sub

Private Sub RefreshChqs(strCondition As String)

    Dim lForeColor As Long
    Dim rsLedger As New ADODB.Recordset, ITM As ListItem
    With rsLedger
    
       .Open "SELECT * FROM VChqBookLedger " & strCondition & " ORDER BY CAST(ChqNo AS NUMERIC)", con_ServerSide, adOpenForwardOnly
       
       Set LVChqs.SmallIcons = ImageList1
       
       LVChqs.ListItems.Clear
       
       Dim img As String
       Dim i As Long
       i = 0
       While .EOF = False
            i = i + 1
         Set ITM = LVChqs.ListItems.add(, ![SNo] & "'" & i, ![ChqNo])
           ITM.SubItems(1) = ![ChequeType] & " "
           ITM.SubItems(2) = Format(![CDate], "dd-MMM-yyyy")
           ITM.SubItems(3) = Format(![ChequeDate], "dd-MMM-yyyy")
           ITM.SubItems(4) = ![AccTitle] & " "
           
           If ![Bounced] = True Then
             ITM.SubItems(5) = "< Bounced >"
             img = "BNC"
           ElseIf ![ChequeType] & "" = "Canceled" Then
             ITM.SubItems(5) = ![Description] & "" & !DeleteDescription & ""
             img = "CAN"
           ElseIf ![Issued] Then
             ITM.SubItems(5) = ![Description] & "" & !DeleteDescription & ""
             img = "ISS"
           Else
             ITM.SubItems(5) = "< Not Issued >"
             img = "CHQ"
           End If
           
           ITM.SmallIcon = img
           img = ""
            ITM.SubItems(6) = IIf(Val(![Amount] & "") = 0, "", Format(Abs(Val(![Amount] & "")), "0.00"))
            
            ITM.SubItems(7) = ![Payee] & " "
             
            ITM.SubItems(8) = Format(IIf(IsNull(!ClearanceDT), "-", !ClearanceDT), "dd-MMM-yy")
            ITM.SubItems(9) = !Handed_Over_To & ""
            If !ChequeType & "" = "Canceled" Then
                lForeColor = vbRed
            Else
                lForeColor = LVChqs.ForeColor
            End If
            ITM.ForeColor = lForeColor
            For j = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(j).ForeColor = lForeColor
            Next
            
            ITM.ListSubItems(4).Tag = Format(!VDate, "dd-MMM-yyyy")
            ITM.ListSubItems(5).Tag = Val(!Vouchers_SNo & "")
            
           .MoveNEXT
           
       Wend
    
       .Close
       Set rsLedger = Nothing
    End With

End Sub

Private Sub txtSearch_GotFocus()
    If txtSearch.Text = txtSearch.Tag Then
        txtSearch.Text = ""
        txtSearch.FontItalic = False
        txtSearch.ForeColor = vbBlack
    End If
End Sub



Private Sub txtSearch_LostFocus()

    If txtSearch = "" Then
        txtSearch.Text = txtSearch.Tag
        txtSearch.FontItalic = True
        txtSearch.ForeColor = &HC0C0C0
    Else
        
    End If
    
End Sub

