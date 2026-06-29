VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmArticlewiseOrders 
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   12660
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
   ScaleHeight     =   8130
   ScaleWidth      =   12660
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
      Height          =   8085
      Left            =   60
      TabIndex        =   1
      Top             =   45
      Width           =   12585
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
         TabIndex        =   18
         Top             =   0
         Width           =   12570
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Articlewise Order Status"
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
            TabIndex        =   19
            Top             =   165
            Width           =   12450
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Articlewise Order Status"
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
            TabIndex        =   20
            Top             =   180
            Width           =   12435
         End
      End
      Begin VB.OptionButton OptSortBy 
         Caption         =   "Report By "
         Height          =   285
         Left            =   210
         TabIndex        =   16
         Top             =   7695
         Visible         =   0   'False
         Width           =   1125
      End
      Begin ComboList.Usercontrol1 cmbArticle 
         Height          =   285
         Left            =   210
         TabIndex        =   15
         Top             =   990
         Width           =   3885
         _ExtentX        =   6853
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
      Begin VB.CheckBox chkDT 
         Caption         =   "Check1"
         Height          =   195
         Left            =   7980
         TabIndex        =   14
         Top             =   780
         Width           =   210
      End
      Begin ComboList.Usercontrol1 cmbCountry 
         Height          =   285
         Left            =   6270
         TabIndex        =   10
         Top             =   990
         Width           =   1650
         _ExtentX        =   2910
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
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   4125
         TabIndex        =   0
         Top             =   990
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
      Begin MSComctlLib.ListView LV 
         Height          =   5925
         Left            =   210
         TabIndex        =   7
         Top             =   1305
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   10451
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1014
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cust Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Country"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Invoice No"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Item Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Item Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Quantity"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTFrom 
         Height          =   285
         Left            =   7950
         TabIndex        =   11
         Top             =   990
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   129957891
         CurrentDate     =   38364
      End
      Begin MSComCtl2.DTPicker DTTo 
         Height          =   285
         Left            =   9600
         TabIndex        =   13
         Top             =   990
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   129957891
         CurrentDate     =   38364
      End
      Begin MSForms.CommandButton cmdPrintWithValue 
         Height          =   360
         Left            =   7305
         TabIndex        =   21
         Top             =   7620
         Width           =   1665
         Caption         =   "Print with Value"
         PicturePosition =   327683
         Size            =   "2937;635"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintAllArticles 
         Height          =   360
         Left            =   5715
         TabIndex        =   17
         Top             =   7620
         Visible         =   0   'False
         Width           =   1665
         Caption         =   "Print All Articles"
         PicturePosition =   327683
         Size            =   "2937;635"
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
         Left            =   9015
         TabIndex        =   12
         Top             =   7620
         Width           =   1665
         Caption         =   "Print    "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   80
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
         Left            =   11250
         TabIndex        =   9
         Top             =   900
         Width           =   1185
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2090;635"
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
         Left            =   10740
         TabIndex        =   8
         Top             =   7620
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
         Left            =   2025
         TabIndex        =   6
         Top             =   7275
         Visible         =   0   'False
         Width           =   7995
         VariousPropertyBits=   746604575
         BackColor       =   15199215
         BorderStyle     =   1
         Size            =   "14102;476"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Quantity :"
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
         Left            =   10005
         TabIndex        =   5
         Top             =   7275
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Quantity (In Words) :"
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
         Left            =   210
         TabIndex        =   4
         Top             =   7275
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
         Left            =   10935
         TabIndex        =   3
         Top             =   7275
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmArticlewiseOrders.frx":0000
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
         Left            =   210
         TabIndex        =   2
         Top             =   735
         Width           =   11010
      End
   End
End
Attribute VB_Name = "frmArticlewiseOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PInvoice As String
Dim PrintSQL  As String
Dim PrintCond As String

Private Sub RefreshLV()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    Dim ITM As ListItem
    Dim LColor As Long
    Dim Cond As String
    Dim dTotal As Long
    
    Cond = ""
    
    If cmbArticle.ID <> "0" Then
        Cond = " WHERE CompItemCode='" & cmbArticle.ID & "'"
    End If
    
    If cmbCust.ID <> "0" Then
        If Cond = "" Then Cond = " WHERE " Else Cond = Cond & " AND "
        Cond = Cond & " CustCode='" & cmbCust.ID & "'"
    End If
    
    If cmbCountry.ID <> "0" Then
        If Cond = "" Then Cond = " WHERE " Else Cond = Cond & " AND "
        Cond = Cond & " Country='" & cmbCountry.ID & "'"
    End If
    
    If chkDT.Value = vbChecked Then
        If Cond = "" Then Cond = " WHERE " Else Cond = Cond & " AND "
        Cond = Cond & " DT Between '" & DTFrom & "' AND '" & DTTo & "'"
    End If
    PrintCond = Cond
    
    LV.ListItems.Clear
    PrintSQL = "SELECT * FROM VArticleWiseOrders " & Cond & " ORDER BY DT"
    'PrintSQL = "SELECT CustCode,Country,FCustomerOrders.OrderNo,DT,Qty,CompItemCode,ItemName FROM FOrderItems INNER JOIN FCustomerOrders ON FCustomerOrders.OrderNo=FOrderItems.OrderNo INNER JOIN Items ON Items.ItemID=FOrderItems.CompItemCode" & Cond
    With rs
        
        .Open PrintSQL, con, adOpenForwardOnly, adLockReadOnly
        
        dTotal = 0
        
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.ListSubItems.add , , ![CustCode] & ""
            ITM.ListSubItems.add , , ![Country] & ""
            ITM.ListSubItems.add , , ![OrderNo] & ""
            ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
            ITM.ListSubItems.add , , !CompItemCode & ""
            ITM.ListSubItems.add , , !ItemName & ""
            
            ITM.ListSubItems.add , , ![Qty] & ""
            dTotal = dTotal + Val(!Qty & "")
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    lblTotal = dTotal
    lblTotWords = ConvertInWords(dTotal, "")

    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub cmbCust_matched()
    
    cmbCountry.ClearVals
    cmbCountry.AddItem "<All Countries>", "0"
    If cmbCust.ID = "0" Then
        cmbCountry.AddVals con, "Distinct Country", "ForeignCustomers", "Country"
    Else
        cmbCountry.AddVals con, "Distinct Country", "ForeignCustomers", "Country", " Where CustCode='" & cmbCust.ID & "'"
    End If
    cmbCountry.ID = "0"
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdPrint_Click()
    
    On Error GoTo err
    If PrintSQL = "" Then Exit Sub
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
        Call PrintReport
        Exit Sub
    End If
    
    Dim strRet As String
    strRet = ShowPopUpMenu(FAll.Left + cmdPrint.Left, cmdPrint.Top + FAll.Top, "Print Order Report", "Print Order Report With &Detail")
    
    Dim rpt As CRAXDDRT.Report
    Dim f As New frmPrevRpt
    
    Dim ShowRpt As Boolean
    Set rpt = rptApp.OpenReport(RptPath & "\rptArticlewiseOrders.rpt")
    
    Dim strTitle As String
    
    If chkDT.Value = vbChecked Then
        strTitle = Format(DTFrom, "dd-MMM-yyyy") & " to " & Format(DTTo, "dd-MMM-yyyy")
    Else
        strTitle = "No Date Range"
    End If
    rpt.FormulaFields(1).Text = "'" & strTitle & "'"
    Select Case strRet
        Case "Print Order Report"
            rpt.FormulaFields(2).Text = True
        Case "Print Order Report With &Detail"
            rpt.FormulaFields(2).Text = False
        Case Else
            Exit Sub
    End Select
    
    Dim strCond As String, strCondSub As String
    strCond = ""
    
    If cmbArticle.ID <> "0" Then
        strCond = "{VArticlewiseSales.ItemID}='" & cmbArticle.ID & "' AND "
        strCondSub = "{ItemPriceHistory_Imported.Itemcode}='" & cmbArticle.ID & "' AND "
    Else
        strCondSub = "{ItemPriceHistory_Imported.Itemcode}='HideEveryThing' AND "
    End If
    
    If cmbCust.ID <> "0" Then
        strCond = strCond & "{VArticlewiseSales.CustCode}='" & cmbCust.ID & "' AND "
    End If
    
    If cmbCountry.ID <> "0" Then
        strCond = strCond & "{VArticlewiseSales.Country}='" & cmbCountry.ID & "' AND "
    End If
    
    If chkDT.Value = vbChecked Then
        strCond = strCond & "{VArticlewiseSales.DT}=#" & DTFrom.Value & "# To #" & DTTo.Value & "# AND "
        strCondSub = strCondSub & "{ItemPriceHistory_Imported.DT}=#" & DTFrom.Value & "# To #" & DTTo.Value & "# AND "
    End If
    
    If strCond <> "" Then
        strCond = Left(strCond, Len(strCond) - 5)
    End If
    
    If strCondSub <> "" Then
        strCondSub = Left(strCondSub, Len(strCondSub) - 5)
    End If
    
'    Dim rptSUB As CRAXDDRT.Report
'    Set rptSUB = rpt.OpenSubreport("ManualSalesData")
'    rptSUB.RecordSelectionFormula = strCondSub
    
    f.ShowReport strCond, rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub PrintReport(Optional iReportType As Integer = 0)

    Dim rpt As CRAXDDRT.Report
    
    If iReportType = 0 Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptArticlewiseOrders.rpt")
    ElseIf iReportType = 1 Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptArticlewiseOrders_Value.rpt")
    End If
    
    Dim f As New frmPrevRpt
    
    Dim strSelection As String
    
    strSelection = "1=1"
    If cmbArticle.ID <> "0" Then
        strSelection = strSelection & " AND {VArticlewiseOrders.ItemID}='" & cmbArticle.ID & "'"
    End If
    
    If cmbCust.ID <> "0" Then
        strSelection = strSelection & " AND {VArticlewiseOrders.CustCode}='" & cmbCust.ID & "'"
    End If
    
    If cmbCountry.ID <> "0" Then
        strSelection = strSelection & " AND {VArticlewiseOrders.Country}='" & cmbCountry.ID & "'"
    End If
    
    If chkDT.Value = vbChecked Then
        strSelection = strSelection & " AND {VArticlewiseOrders.DT}=#" & DTFrom & "# TO #" & DTTo & "#"
    End If
    If chkDT.Value = vbChecked Then
        strTitle = Format(DTFrom, "dd-MMM-yyyy") & " to " & Format(DTTo, "dd-MMM-yyyy")
    Else
        strTitle = "No Date Range"
    End If
    rpt.FormulaFields(1).Text = "'" & strTitle & "'"
    
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub cmdPrintAllArticles_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptArticlewiseOrders.rpt")
    Dim f As New frmPrevRpt
    
    
    f.ShowReport "{VArticlewiseSales.OrderDT}=#" & DTFrom.Value & "# TO #" & DTTo.Value & "#", rpt
    
End Sub

Private Sub cmdPrintWithValue_Click()

    If PrintSQL = "" Then Exit Sub
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
        Call PrintReport(1)
        Exit Sub
    End If
    
End Sub

Private Sub cmdRefresh_Click()
   
    Call RefreshLV
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
   
    On Error GoTo err
    
    Call SaveLV(LV)
    
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "Distinct CustCode", "ForeignCustomers", "CustCode")
    cmbCust.ID = "0"
    
    '\/\/\/\/\/\/\/\/\/
    cmbArticle.DropDownOffLine = True
    cmbArticle.AddItem "<All Articles>", "0"
    Call cmbArticle.AddVals(con, "ItemID + ' ' + ItemName + ' ' + ISNULL(ItemSize,'')", "Items", "ItemID")
    cmbArticle.ID = "0"
    
    Me.KeyPreview = True
    cmbCust.ListHeight = 4000
    cmbCountry.ListHeight = 4000
    cmbArticle.ListHeight = 6000
    
    DTFrom.Value = Date
    DTTo.Value = Date
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
    
End Sub

