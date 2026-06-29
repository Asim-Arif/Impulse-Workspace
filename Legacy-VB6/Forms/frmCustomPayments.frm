VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmCustomPayments 
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12225
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12225
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   45
      TabIndex        =   1
      Top             =   30
      Width           =   12120
      Begin VB.CommandButton cmdShowAllCustomers 
         Caption         =   "Show All Customers Balance"
         Height          =   390
         Left            =   285
         TabIndex        =   17
         Top             =   8325
         Width           =   2970
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E7EBEF&
         Caption         =   "All Customers Balance :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3465
         Left            =   3165
         TabIndex        =   15
         Top             =   2400
         Visible         =   0   'False
         Width           =   6195
         Begin MSComctlLib.ListView LVAll 
            Height          =   3120
            Left            =   90
            TabIndex        =   16
            Top             =   225
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   5503
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
            BackColor       =   16777215
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
               Text            =   "Currency"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Balance"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Exch. Rate"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Balance PK"
               Object.Width           =   2469
            EndProperty
         End
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         ItemData        =   "frmCustomPayments.frx":0000
         Left            =   2445
         List            =   "frmCustomPayments.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1050
         Width           =   2160
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   1050
         Width           =   2130
         _ExtentX        =   3757
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
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   3
         Top             =   15
         Width           =   12120
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " Custom Payment Status"
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
            Left            =   60
            TabIndex        =   4
            Top             =   150
            Width           =   12000
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   " Custom Payment Status"
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
            TabIndex        =   5
            Top             =   165
            Width           =   11985
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6585
         Left            =   285
         TabIndex        =   10
         Top             =   1350
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   11615
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         TextBackground  =   -1  'True
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
            Text            =   "S No."
            Object.Width           =   1014
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Invoice No"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Order No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Form E"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "FDBC No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "FDBC Date"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Inv. Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Inv. Amt"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Balance"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Status"
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
         Left            =   2445
         TabIndex        =   14
         Top             =   795
         Width           =   2130
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   4635
         TabIndex        =   13
         Top             =   990
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
      Begin MSForms.CommandButton cmdCancel 
         Height          =   360
         Left            =   10350
         TabIndex        =   11
         Top             =   8535
         Width           =   1665
         Caption         =   "      Close    "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.TextBox lblTotWords 
         Height          =   270
         Left            =   2100
         TabIndex        =   9
         Top             =   7980
         Visible         =   0   'False
         Width           =   7545
         VariousPropertyBits=   746604575
         BackColor       =   15199215
         BorderStyle     =   1
         Size            =   "13309;476"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount :"
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
         Height          =   270
         Index           =   1
         Left            =   9660
         TabIndex        =   8
         Top             =   7980
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount (In Words) :"
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
         Height          =   270
         Index           =   2
         Left            =   285
         TabIndex        =   7
         Top             =   7980
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   10530
         TabIndex        =   6
         Top             =   7980
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Customer"
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
         Index           =   7
         Left            =   285
         TabIndex        =   2
         Top             =   795
         Width           =   2130
      End
   End
End
Attribute VB_Name = "frmCustomPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PInvoice As String
Dim InvsForPrint As String

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    If cmbCust.id <> "0" Then
        SQL = " Where CustCode='" & cmbCust.id & "'"
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        If SQL = "" Then
            SQL = " Where "
        Else
            SQL = SQL & " And "
        End If
        SQL = SQL & " Country='" & cmbCountry.Text & "'"
    End If
    
    SQL = "Select PInvoice From VFCustPayments " & SQL
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        cmbInvNo.AddItem "<All Invoices>"
        cmbInvNo.Clear
        Do Until .EOF
            cmbInvNo.AddItem .Fields(0) & ""
            
            .MoveNext
        Loop
        .Close
        
        
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_PaymentStatus"
        If cmbCust.id <> "0" Then
            .Parameters("@CustCode").Value = cmbCust.id
        End If
        
        .Parameters("@Status").Value = cmbStatus.ListIndex
        .Parameters("@CustomPayments").Value = 1
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
    End With
        
    Dim ITM As ListItem
    Dim lColor As Long
    Dim FDBCNo As String
    Dim BalAmt As Double
    
    LV.ListItems.Clear
    
    With rs
        InvsForPrint = "In["
        Do Until .EOF
            If ![EntryType] = 0 Then
                
                FDBCNo = ![FDBCNo] & ""
                BalAmt = Val(![CustomAmtToRcv] & "")
                
                Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
                ITM.ListSubItems.add , , !CustCode & ""
                ITM.ListSubItems.add , , ![CustomInvoice] & ""
                
                ITM.ListSubItems.add , , ![Orders] & ""
                ITM.ListSubItems.add , , ![FormE] & ""
                ITM.ListSubItems.add , , ![FDBCNo] & ""
                
                ITM.ListSubItems.add , , Format(![FDBCDT] & "", "dd-MMM-yyyy")
                ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
                
                'ITM.ListSubItems.add , , ![TotalCustomAmtRcvd] & "" 'AmtRcvd=TotalCustomAmt
                ITM.ListSubItems.add , , Round(Val(![TotalCustomAmt] & ""))
                ITM.ListSubItems.add , , Round(Val(![CustomAmtToRcv] & ""))
                
                
                lColor = LV.ForeColor
                
                InvsForPrint = InvsForPrint & "'" & ![CustomInvoice] & "',"
                
            Else
                Set ITM = LV.ListItems.add(, , "")
                
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems(2).Tag = ![EntryID] & ""
                ITM.ListSubItems.add(, , "").Tag = !PRCVchrNo & ""
                ITM.ListSubItems.add , , ![FDBCNo] & ""
                
                ITM.ListSubItems.add(, , Format(![DT] & "", "dd-MMM-yyyy")).Tag = Val(!ExchRate & "")
                ITM.ListSubItems.add(, , Format(![DT] & "", "dd-MMM-yyyy")).Tag = Val(!ExchRate & "")
                
                ITM.ListSubItems.add , , ![TotalCommercialAmt] & ""
                ITM.ListSubItems.add , , ![Description] & ""
                
                lColor = vbBlue
            End If
            ITM.Tag = ![CustCode] & ""
            ITM.ListSubItems(1).Tag = ![CustomInvoice] & ""
            ITM.ListSubItems(4).Tag = FDBCNo
            ITM.ListSubItems(8).Tag = BalAmt
            
            ITM.ForeColor = lColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ToolTipText = ITM.ListSubItems(i).Text
                ITM.ListSubItems(i).ForeColor = lColor
                
                If ITM.Text <> "" Then
                    ITM.ListSubItems(8).ForeColor = vbRed
                End If
            Next i
            
            .MoveNext
        Loop
        .Close
    End With
    CalculateTotals
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbinvno_Change()
    Call cmbinvno_click
End Sub

Private Sub cmbinvno_click()

    On Error GoTo err
    If cmbInvNo.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    LV.ListItems.Clear
    With rs
        .Open "Select * From VFCustPayments", con, adOpenForwardOnly, adLockReadOnly
        Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
        ITM.ListSubItems.add , , ![PInvoice] & ""
        ITM.ListSubItems.add , , ![Orders] & ""
        ITM.ListSubItems.add , , ![TotalCommercialAmt] & ""
        ITM.ListSubItems.add , , Val(![TotalCommercialAmt] & "") - Val(![TotalCommercialAmtToRcv] & "")
        .MoveNext
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    If cmbInvNo.MatchFound = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report 'New rptCustomInvoice
    Set rpt = rptApp.OpenReport(RptPath & "\rptCustomInvoice.rpt") '
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptProformas.PInvoice}='" & cmbInvNo.Text & "' ", rpt
    'frmPrevRpt.ShowReport
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshLV
End Sub

Private Sub cmdShowAllCustomers_Click()

    On Error GoTo err
    
    If Frame2.Visible = False Then
        Dim ITM As ListItem
        Dim rs As New ADODB.Recordset
        Dim dExchRate As Double, dTotal As Double
        With rs
            .Open "SELECT Curr,SUM(TotalCustomAmtToRcv) AS TotalAmt FROM VrptCustomInvoice WHERE FreeSamples=0 GROUP BY Curr", con, adOpenForwardOnly, adLockReadOnly
            LVAll.ListItems.Clear
            Do Until .EOF
                Set ITM = LVAll.ListItems.add(, , !Curr & "")
                ITM.ListSubItems.add , , Format(Val(!TotalAmt & ""), "#,##")
                dExchRate = GetSingleDoubleValue("ExchRate", "VCurrencyExchangeRates", " WHERE Currency='" & !Curr & "'")
                ITM.ListSubItems.add , , dExchRate
                ITM.ListSubItems.add , , Format(Val(!TotalAmt & "") * dExchRate, "#,##")
                dTotal = dTotal + (Val(!TotalAmt & "") * dExchRate)
                .MoveNext
            Loop
            .Close
        End With
        Set rs = Nothing
        Set ITM = LVAll.ListItems.add
        ITM.ListSubItems.add
        ITM.ListSubItems.add(, , "Total").Bold = True
        ITM.ListSubItems.add(, , Format(dTotal, "#,##")).Bold = True
        
        Frame2.Visible = True
        cmdShowAllCustomers.Caption = "Hide All Customers Balance"
    Else
        Frame2.Visible = False
        cmdShowAllCustomers.Caption = "Show All Customers Balance"
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub



Private Sub Form_Load()
   
    On Error GoTo err
    

    
    Call SaveLV(LV)
     
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'HICO'")
    cmbCust.id = "0"
    
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
        With cmbStatus
        
        .AddItem "<UnPaid>"
        .AddItem "<Paid>"
        .AddItem "<All>"
    End With
    
    cmbStatus.ListIndex = 0
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub



Private Sub Form_Resize()

    On Error Resume Next
    FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
    
End Sub

Private Sub CalculateTotals()

    With LV
        For i = 1 To .ListItems.count
            'dQty = dQty + Val(.TextMatrix(i, 4))
            dTotal = dTotal + Val(.ListItems(i).SubItems(7) & "")
        Next i
    End With
    'txttQty = dQty
    lblTotal = dTotal
    lblTotWords = ConvertInWords(CCur(dTotal), GetSingleStringValue("Curr", "ForeignCustomers", " WHERE CustCode='" & cmbCust.id & "'"))
    
End Sub


Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    
    Dim strRet As String
    
    
    If LV.ListItems.count = 0 Then
        strRet = ShowPopUpMenu1("&Refresh")
'    ElseIf LV.SelectedItem.ListSubItems(4).Tag = "" Then
'        strRet = ShowPopUpMenu1("&Refresh", "-", "Assign &FDBC No. To This Invoice.")
    ElseIf LV.SelectedItem.ListSubItems(2).Tag <> "" Then
        If LV.SelectedItem.ListSubItems(3).Tag = "" Then
            strRet = ShowPopUpMenu1("Post PRC")
        End If
    ElseIf Val(LV.SelectedItem.ListSubItems(8).Tag) > 0 Then
        If LV.SelectedItem.Text = "" Then
            strRet = ShowPopUpMenu1("&Refresh", "-", "&Receive Payments For This Invoice", "-", "&Print Bank Letter")
        Else
            strRet = ShowPopUpMenu1("&Refresh", "-", "&Receive Payments For This Invoice", "-", "&Print Payment Status For This Invoice", "-", "&Print Payment Status Of All", "Print Status Customerwise")
        End If
    'ElseIf LV.SelectedItem.SubItems(4) <> "" Then
    
    Else
        strRet = ShowPopUpMenu1("&Refresh", "-", "&Print Payment Status For This Invoice", "-", "&Print Payment Status Of All", "Print Status Customerwise")
    End If
    
    
    'strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Foreign Customers", "-", "Customer Catalog", "-", "Customer &Ledger", "-", "Assign &FDBC No.", "-", "&Receive Payments")
    Dim f As Form
    
    Select Case strRet
        Case "&Refresh"
            Call RefreshLV
        Case "Assign &FDBC No. To This Invoice."
        
            Set f = New frmFCustFDBCNo
            Load f
            Call f.ShowForm(LV.SelectedItem.Tag, LV.SelectedItem.ListSubItems(1).Tag)
            
        Case "&Receive Payments For This Invoice"
        
            Set f = New frmFCustRcvPymnts
            Load f
            Call f.ShowForm(LV.SelectedItem.Tag, LV.SelectedItem.ListSubItems(1).Tag, True)
            
        Case "&Print Bank Letter"
            Set f = New frmPrevRpt
            f.ShowReport "{VPaymentsRcvd.EntryID}=" & LV.SelectedItem.ListSubItems(2).Tag & " ", rptBankLetter
            
        Case "&Print Payment Status For This Invoice"
            
            Screen.MousePointer = vbHourglass
            Dim rpt As CRAXDDRT.Report 'New rptPaymentStatus
            Set rpt = rptApp.OpenReport(RptPath & "\rptPaymentStatus.rpt") '
            rpt.FormulaFields(1).Text = True
            Load frmPrevRpt
            frmPrevRpt.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.ListSubItems(1).Tag & "'", rpt
            'frmPrevRpt.ShowReport
            Screen.MousePointer = vbDefault
            
        Case "&Print Payment Status Of All"
        
            Screen.MousePointer = vbHourglass
            Set rpt = rptApp.OpenReport(RptPath & "\rptPaymentStatus.rpt") '
            rpt.FormulaFields(1).Text = True
            Load frmPrevRpt
            frmPrevRpt.ShowReport "{CustomInvoice.CustomInvoice} " & Left(InvsForPrint, Len(InvsForPrint) - 1) & "] ", rpt
            'frmPrevRpt.ShowReport
            Screen.MousePointer = vbDefault
        Case "Print Status Customerwise"
        
            Screen.MousePointer = vbHourglass
            Set rpt = rptApp.OpenReport(RptPath & "\PaymentStatusCustomerWise.rpt") '
            rpt.FormulaFields(1).Text = True
            Load frmPrevRpt
            frmPrevRpt.ShowReport "{CustomInvoice.CustomInvoice} " & Left(InvsForPrint, Len(InvsForPrint) - 1) & "] ", rpt
            'frmPrevRpt.ShowReport
            Screen.MousePointer = vbDefault
            
        Case "Post PRC"
            Call PostPRC
    End Select
    
End Sub

Private Sub PostPRC()
    
    Dim strInvoiceNo As String
    strInvoiceNo = LV.SelectedItem.ListSubItems(1).Tag
    Dim lInvoiceAuthCount As Long
    lInvoiceAuthCount = GetSingleLongValue("COUNT(*)", "CustomInvoiceAuthorization", " WHERE CustomInvoice='" & strInvoiceNo & "'")
    
    If lInvoiceAuthCount = 0 Then
        MsgBox "Invoice not Booked, Please Authorize the invoice.", vbInformation
        Exit Sub
    End If
    
    Dim lPRCCount As Long
    'lPRCCount = GetSingleLongValue("COUNT(*)", "PRC", " WHERE CustomInvoice='" & strInvoiceNo & "'")
    lPRCCount = GetSingleLongValue("COUNT(*)", "FCustPayments", " WHERE ISNULL(PRCVchrNo,'')<>'' AND EntryID=" & Val(LV.SelectedItem.ListSubItems(2).Tag))
    If lPRCCount > 0 Then
        MsgBox "PRC Already Posted.", vbInformation
        Exit Sub
    End If
    
    frmPRC.ShowMe Val(LV.SelectedItem.ListSubItems(2).Tag), LV.SelectedItem.ListSubItems(1).Tag, Val(LV.SelectedItem.SubItems(8)), LV.SelectedItem.SubItems(4), Val(LV.SelectedItem.ListSubItems(7).Tag)
End Sub
