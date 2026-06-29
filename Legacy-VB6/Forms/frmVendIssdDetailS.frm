VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmVendIssdDetailS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Status"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleMode       =   0  'User
   ScaleWidth      =   10869.48
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12105
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item Status"
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
         TabIndex        =   8
         Top             =   165
         Width           =   8955
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Status"
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
         TabIndex        =   9
         Top             =   180
         Width           =   8940
      End
   End
   Begin VB.Frame FLV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6675
      Left            =   30
      TabIndex        =   0
      Top             =   645
      Width           =   12030
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   105
         ScaleHeight     =   255
         ScaleWidth      =   4140
         TabIndex        =   4
         Top             =   945
         Visible         =   0   'False
         Width           =   4170
         Begin MSForms.TextBox txtDescription 
            Height          =   300
            Left            =   600
            TabIndex        =   6
            Top             =   -15
            Width           =   645
            VariousPropertyBits=   679495707
            Size            =   "1138;529"
            SpecialEffect   =   3
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtName 
            Height          =   300
            Left            =   -15
            TabIndex        =   5
            Top             =   -15
            Width           =   645
            VariousPropertyBits=   679495707
            Size            =   "1138;529"
            SpecialEffect   =   3
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5580
         Left            =   90
         TabIndex        =   3
         Top             =   510
         Width           =   11820
         _ExtentX        =   20849
         _ExtentY        =   9843
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ColHdrIcons     =   "ImageList1"
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S.#"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vend ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Vender"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ID #"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Issd Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Rcvd Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Wastage"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Balance"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lblItemCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   555
         TabIndex        =   18
         Top             =   180
         Width           =   3510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Item :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   210
         Width           =   435
      End
      Begin VB.Label lblBal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7770
         TabIndex        =   16
         Top             =   4920
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Net Balance :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6780
         TabIndex        =   15
         Top             =   4950
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblRcvd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5580
         TabIndex        =   14
         Top             =   4920
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Net Rcvd :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4770
         TabIndex        =   13
         Top             =   4950
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lblIssd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3570
         TabIndex        =   12
         Top             =   4920
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Net Issued :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2640
         TabIndex        =   11
         Top             =   4950
         Visible         =   0   'False
         Width           =   885
      End
      Begin MSForms.CheckBox chkZero 
         Height          =   285
         Left            =   90
         TabIndex        =   10
         Top             =   4890
         Visible         =   0   'False
         Width           =   2025
         BackColor       =   -2147483633
         ForeColor       =   8388608
         DisplayStyle    =   4
         Size            =   "3572;503"
         Value           =   "1"
         Caption         =   "Show Zero Balanced"
         Accelerator     =   90
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   8955
         TabIndex        =   2
         Top             =   6210
         Visible         =   0   'False
         Width           =   1425
         ForeColor       =   0
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
      Begin MSForms.CommandButton cmdClose 
         Height          =   360
         Left            =   10470
         TabIndex        =   1
         Top             =   6210
         Width           =   1425
         ForeColor       =   0
         Caption         =   " Close     "
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
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   1005
      Top             =   5580
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
      Left            =   360
      Top             =   5340
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
            Picture         =   "frmVendIssdDetailS.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendIssdDetailS.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmVendIssdDetailS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OrderNo As String, ItemCode As String, ProcessID As Long
Dim VendID As Long
Private Sub cmdAdd_Click()

    On Error GoTo err
    
    LV.Sorted = False
    
    Dim iAffect As Integer
    
    con.Execute "Insert into Compositions (Symbol,Name)" & _
     " Values('" & txtUnit & "','" & txtUnitDesc & "')", iAffect
     
    If iAffect = 1 Then
        Call FillList
        txtUnit = ""
        txtUnitDesc = ""
        txtUnit.SetFocus
    End If
    
    LV.Sorted = True
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub chkZero_Click()
    Call ShowDetail(ItemCode, ProcessID)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    Screen.MousePointer = vbHourglass
    Dim DateFrom As Date
    Dim DateTo As Date
    DateFrom = DateSerial(cmbYear, 1, 1)
    DateTo = DateSerial(cmbYear, 12, 31)
    With cr1
        .ReportFileName = RptPRPath & "\Holidays.rpt"
        .DataFiles(0) = DatabasePath
        
        .SelectionFormula = "Year({Holidays.DT})=" & cmbYear
          
        .Formulas(0) = "ForYear='Year-" & cmbYear & "'"
        
        .Action = 1
        .PageZoomNext
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillList()

    On Error GoTo err
    Dim rsList As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Sql As String
    
    Sql = "select * from Compositions Order By Symbol"

    With rsList
        .CursorLocation = adUseClient
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.Add(, ![ID] & "'", ![Symbol] & "")
            ITM.SubItems(1) = ![Name] & ""
            .MoveNext
        Next
        .Close
    End With
    
    Set rsList = Nothing
    Sql = ""
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    'Call FillList
    
End Sub

Private Sub Form_Resize()
    Me.Top = 1700 'MainForm.Toolbar1.Top + MainForm.Toolbar1.Height + 100
End Sub



Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Exit Sub
    With LV
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
                ColumnHeader.Icon = 2
            Else
                .SortOrder = lvwAscending
                ColumnHeader.Icon = 1
            End If
        Else
            LV.ColumnHeaderIcons = Nothing
            Set LV.ColumnHeaderIcons = ImageList1
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
            ColumnHeader.Icon = 1
        End If
        .Sorted = True
    End With
End Sub

Private Sub LV_DblClick()
    Exit Sub
    'Edit This Record...
    If LV.ListItems.Count = 0 Then Exit Sub
    '-------------------------------------------------------
    PicEdit.Move LV.Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(1).Width + LV.ColumnHeaders(2).Width
    
    
    txtName.Move 0, 0, LV.ColumnHeaders(1).Width
    txtName.Text = LV.SelectedItem
    
    txtDescription.Move txtName.Width - 10, 0, LV.ColumnHeaders(2).Width
    txtDescription.Text = LV.SelectedItem.SubItems(1)
    
    PicEdit.Visible = True
    
    With txtName
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Exit Sub
    PicEdit.Visible = False
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'For Now
    Exit Sub
    If LV.ListItems.Count = 0 Then Exit Sub
    On Error Resume Next
    If KeyCode = 46 Then
        If MsgBox("Do You Want To Delete This Component?", vbQuestion + vbYesNo) = vbYes Then
            Dim Affect As Integer
            con.Execute "Delete from Compositions Where ID=" & Val(LV.SelectedItem.Key), Affect
            If Affect = 1 Then
                LV.ListItems.Remove (LV.SelectedItem.Index)
            End If
        End If
    End If
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub txtDescription_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    On Error GoTo err
    Dim iAffect As Integer
    If KeyCode.value = 13 Then
        con.Execute "Update Compositions Set Symbol='" & txtName & _
         "',Name='" & txtDescription & "' Where ID=" & _
         Val(LV.SelectedItem.Key), iAffect
        If iAffect = 1 Then
            'LV.SelectedItem.Text = txtName
            'LV.SelectedItem.SubItems(1) = txtDescription
            PicEdit.Visible = False
            Call FillList
        End If
    ElseIf KeyCode.value = 27 Then
        PicEdit.Visible = False
    End If
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    If KeyCode.value = 13 Then
        With txtDescription
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyCode.value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub
Public Sub ShowDetailOld(sItemCode As String, lProcessID As Long, lVendID As Long, Optional ShowForm As Boolean = True)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim ITM As ListItem
    Dim ShowZero As Boolean
    ShowZero = chkZero.value
    
    OrderNo = sOrderNo
    ItemCode = sItemCode
    ProcessID = lProcessID
    VendID = lVendID
    lblOrdNo = OrderNo
    lblItemCode = ItemCode
    
    Sql = "Select * From VVendLedger1 Where ItemCode='" & ItemCode & "' and ProcessID=" & ProcessID & " and VendID=" & VendID
    If Not ShowZero Then
        Sql = Sql & " and RemQty>0 "
    End If
    Sql = Sql & " Order By DT ,IssQty Desc"
    
    With rs
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        lblIssd = ""
        lblRcvd = ""
        lblBal = ""
        Dim Bal As Long
        Do Until .EOF
            lblVender = ![VenderName] & ""
            Set ITM = LV.ListItems.Add(, , LV.ListItems.Count + 1)
            ITM.ListSubItems.Add , , Format(![DT], "dd-MMM-yyyy")
            ITM.ListSubItems.Add , , ![RecieptID] & ""
            ITM.ListSubItems.Add(, , IIf(Val(![IssQty] & "") > 0, Val(![IssQty] & ""), "0")).ForeColor = RGB(0, 100, 255)
            
            ITM.ListSubItems.Add(, , IIf(Val(![IssQty] & "") <= 0, Val(![RcvdQty] & ""), "0")).ForeColor = RGB(255, 100, 0)
            ITM.ListSubItems.Add(, , Val(![Wastage] & "")).ForeColor = RGB(255, 0, 0)
            Bal = Bal + Val(ITM.SubItems(3)) - Val(ITM.SubItems(4)) - Val(ITM.SubItems(5))
            ITM.ListSubItems.Add(, , Bal).ForeColor = RGB(50, 200, 100)
            ITM.ListSubItems(6).Bold = True
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    If ShowForm Then Me.Show 1
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub ShowDetail(sItemCode As String, lProcessID As Long)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim ITM As ListItem
    Dim ShowZero As Boolean
    ShowZero = chkZero.value
    
    ItemCode = sItemCode
    ProcessID = lProcessID
    VendID = lVendID
    lblOrdNo = OrderNo
    lblItemCode = ItemCode
    
    Sql = "Select * From VVendLedger1 Where ItemCode='" & ItemCode & "' and ProcessID=" & ProcessID & " AND issqty>0 and issqty-rcvdqty>0"
    If Not ShowZero Then
        Sql = Sql & " and RemQty>0 "
    End If
    Sql = Sql & " Order By DT ,IssQty Desc"
    
    With rs
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        lblIssd = ""
        lblRcvd = ""
        lblBal = ""
        Dim Bal As Long
        Do Until .EOF
            lblVender = ![VenderName] & ""
            Set ITM = LV.ListItems.Add(, , LV.ListItems.Count + 1)
            ITM.ListSubItems.Add , , ![VendID1] & ""
            ITM.ListSubItems.Add , , ![VenderName] & ""
            ITM.ListSubItems.Add , , Format(![DT], "dd-MMM-yyyy")
            ITM.ListSubItems.Add , , ![RecieptID] & ""
            'ITM.ListSubItems.Add(, , IIf(Val(![IssQty] & "") > 0, Val(![IssQty] & ""), "0")).ForeColor = RGB(0, 100, 255)
            ITM.ListSubItems.Add(, , Val(![IssQty] & "")).ForeColor = RGB(0, 100, 255)
            'ITM.ListSubItems.Add(, , IIf(Val(![IssQty] & "") <= 0, Val(![RcvdQty] & ""), "0")).ForeColor = RGB(255, 100, 0)
            ITM.ListSubItems.Add(, , Val(![RcvdQty] & "")).ForeColor = RGB(255, 100, 0)
            ITM.ListSubItems.Add(, , Val(![Wastage] & "")).ForeColor = RGB(255, 0, 0)
            'Bal = Bal + Val(ITM.SubItems(3)) - Val(ITM.SubItems(4)) - Val(ITM.SubItems(5))
            'ITM.ListSubItems.Add(, , Bal).ForeColor = RGB(50, 200, 100)
            Bal = Val(ITM.SubItems(5)) - Val(ITM.SubItems(6)) - Val(ITM.SubItems(7))
            ITM.ListSubItems.Add(, , Bal).ForeColor = RGB(50, 200, 100)
            ITM.ListSubItems(6).Bold = True
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Me.Show 1
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

