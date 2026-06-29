VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmCustomPList 
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10860
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
   Icon            =   "frmCustomPList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   10860
   Visible         =   0   'False
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
      Height          =   9285
      Left            =   45
      TabIndex        =   8
      Top             =   30
      Width           =   10770
      Begin VB.TextBox txtEdit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6360
         TabIndex        =   20
         Top             =   2505
         Visible         =   0   'False
         Width           =   1000
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   1035
         Width           =   1740
         _ExtentX        =   3069
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
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3810
         MaxLength       =   150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1035
         Width           =   1005
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
         TabIndex        =   12
         Top             =   0
         Width           =   10560
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Packing List"
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
            TabIndex        =   13
            Top             =   165
            Width           =   10605
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Packing List"
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
            TabIndex        =   14
            Top             =   180
            Width           =   10590
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6855
         Left            =   285
         TabIndex        =   4
         Top             =   1350
         Width           =   10170
         _ExtentX        =   17939
         _ExtentY        =   12091
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Carton #"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Qnty"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   5715
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Lot No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Batch #"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   5175
         TabIndex        =   9
         Top             =   8535
         Width           =   5295
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   90
            TabIndex        =   19
            Top             =   195
            Width           =   1665
            Caption         =   "Print Packing List   "
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
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   3510
            TabIndex        =   7
            Top             =   195
            Width           =   1665
            Caption         =   "Close"
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
         Begin MSForms.CommandButton cmdAdd 
            Height          =   360
            Left            =   1800
            TabIndex        =   5
            Top             =   195
            Visible         =   0   'False
            Width           =   1665
            Caption         =   "New Packing List   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdEdit 
            Height          =   360
            Left            =   1800
            TabIndex        =   6
            Top             =   195
            Width           =   1665
            Caption         =   "Edit Packing List  "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   69
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin MSForms.ComboBox cmbCustomInvoice 
         Height          =   285
         Left            =   4830
         TabIndex        =   3
         Top             =   1035
         Width           =   1725
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3043;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
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
         Caption         =   " No. Of Pieces :"
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
         Left            =   3285
         TabIndex        =   18
         Top             =   8250
         Width           =   1275
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " No. Of Cartons :"
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
         TabIndex        =   17
         Top             =   8250
         Width           =   1425
      End
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   270
         TabIndex        =   16
         Top             =   8850
         Visible         =   0   'False
         Width           =   1545
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "2725;609"
         Value           =   "0"
         Caption         =   "Print On Save"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label lblTotCartons 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1695
         TabIndex        =   15
         Top             =   8250
         Width           =   1605
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Customer                     Country                       Currency       Invoice No."
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
         TabIndex        =   11
         Top             =   795
         Width           =   6270
      End
      Begin VB.Label lblTotalQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4545
         TabIndex        =   10
         Top             =   8250
         Width           =   1605
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2055
         TabIndex        =   1
         Top             =   1035
         Width           =   1725
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3043;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmCustomPList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String
Dim CustomPackingList As Boolean
Dim iColNo As Integer
Public Sub ShowForm(CustomPList As Boolean)
    CustomPackingList = CustomPList
    Me.Show
End Sub

Public Sub EditOrder(TempOrdNo As String)
    
    On Error GoTo err
    strOrderNo = TempOrdNo
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select * From VFCustomerOrders Where OrderNo='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        txtOrderNo = ![OrderNo] & ""
        OrderDT = ![DT] & ""
        cmbTradeTerms = ![TradeTerms] & ""
        
        cmbPayment = ![PaymentTerms] & ""
        cmbTrans = ![TransShipment] & ""
        cmbPartial = ![PartialShipment] & ""
        txtPackaging = ![Packaging] & ""
        DeliveryDT = ![DeliveryDT] & ""
        
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![CustItemID] & "'", .AbsolutePosition)
            ITM.Tag = ![ITemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            ITM.ListSubItems.add , , Left(![ITemID] & "", 2) '![CatID] & ""
            ITM.ListSubItems.add , , ![ITemID] & ""
            ITM.ListSubItems.add(, ![FullItemName] & "", ![FullItemName] & "").Tag = ![Description] & ""
            
            ITM.SubItems(4) = txtUnit
            
            ITM.SubItems(5) = Format(![Price] & "", "#0.00")
            ITM.SubItems(6) = Val(![Qty] & "")
            ITM.SubItems(7) = Val(![Qty] & "") * Val(![Price])
   
            .MoveNEXT
        Loop
        .Close
    End With
    

    txtQty = 0
    cmbItems.SetFocus
    
    Set rs = Nothing
    Me.Show
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    cmbCountry.Tag = cmbCust.Text & cmbCountry
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_GetPackingList"
        .Parameters("@Customer") = cmbCountry.Tag
        .Parameters("@CustomPackingList") = Abs(CustomPackingList)
        Set rs = .Execute
    End With
    
    cmbCustomInvoice.Clear
    With rs
        '.Open "Select PInvoice From VCustomPList3 Where CustCode+Country='" & cmbCountry.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            cmbCustomInvoice.AddItem .Fields(0) & ""
            cmbCustomInvoice.List(cmbCustomInvoice.ListCount - 1, 1) = .Fields(1) & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_matched()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Country,Curr From ForeignCustomers Where CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCustomInvoice_Change()
    Call cmbCustomInvoice_Click
End Sub

Private Sub cmbCustomInvoice_Click()
    
    On Error GoTo err
    If cmbCustomInvoice.MatchFound = False Then Exit Sub
    
    LV.ListItems.Clear
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim str As String
    Dim Qty As String
    With rs
        'Debug.Print "Start : " & Now
        .Open "Select * From VrptCustomPList Where CustomInvoice='" & cmbCustomInvoice.Text & "' Order By Cast(CartonFrom as int)", con, adOpenForwardOnly, adLockReadOnly
        'Debug.Print "rs Opened : " & Now
        
        Dim Arr As Variant
        Arr = rsExecute("Select TotalCartons,TotalQty From VTotalCartons Where CustomInvoice='" & cmbCustomInvoice.Text & "'")
        lblTotCartons.Caption = Arr(0) & ""
        lblTotalQty.Caption = Arr(1) & ""
        
        Do Until .EOF
            
            If Val(![CartonTo] & "") = 0 Then
                str = Format(![CartonFrom] & "", "00")
                
                Qty = ![Qty] & " Pcs"
            Else
                str = Format(![CartonFrom] & "", "00 to ") & Format(![CartonTo] & "", "00")
                Qty = ![Qty] & " Pcs Each"
            End If
            
            Set ITM = LV.ListItems.add(, , str)
            ITM.SubItems(1) = Qty
            ITM.SubItems(2) = ![CompItemCode] & ""
            ITM.SubItems(3) = ![Description] & ""
            
            ITM.SubItems(4) = !LotNo & ""
            ITM.SubItems(5) = !BatchNo & ""
            .MoveNEXT
        Loop
    
        .Close
    End With
    Set rs = Nothing
    'Debug.Print "End : " & Now
    
    Exit Sub
err:
    MsgBox err.Description
    
    
End Sub



Private Sub cmdAdd_Click()
    'frmNewCustomPListM.Show
    Dim retStr As String
    retStr = ShowPopUpMenu(FAll.Left + Frame6.Left + cmdAdd.Left + 300, Frame6.Top + FAll.Top + cmdAdd.Top - 200, "&Manual Packing List", "-", "&Auto Packing List")
    
    Select Case retStr
        Case "&Manual Packing List"
            frmNewCustomPListM.Show
        Case "&Auto Packing List"
            Load frmNewCustomPListA
            frmNewCustomPListA.ShowForm (False)
'        Case "&ReGenerate Packing List"
'            Load frmNewCustomPListA
'            frmNewCustomPListA.ShowForm (True)
    End Select
    
End Sub

Private Sub cmdEdit_Click()
            
    Exit Sub
    If cmbCustomInvoice.MatchFound = False Then Exit Sub
    
    Load frmNewCustomPListM_New
    frmNewCustomPListM_New.EditOrder cmbCustomInvoice
    
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
        
    Dim rpt As CRAXDDRT.Report
    If CustomPackingList Then
        'Dim retStr As String
        'retStr = ShowPopUpMenu(FAll.Left + Frame6.Left + cmdPrint.Left + 300, Frame6.Top + FAll.Top + cmdPrint.Top - 200, "Print On &Blank Page", "-", "Print On &Pre-Printed Page")
        'Select Case retStr
        '    Case "Print On &Blank Page"
        Set rpt = rptApp.OpenReport(RptPath & "\rptCustomPList.rpt") 'New rptCustomPList
        '    Case "Print On &Pre-Printed Page"
        '        Set rpt = New rptCustomPListPre
        '    Case Else
        '        Exit Sub
        'End Select
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptCommercialPList.rpt") 'New rptCommercialPList
    End If
    
    Screen.MousePointer = vbHourglass
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VCustomPList.CustomInvoice}='" & cmbCustomInvoice & "' ", rpt
    
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdSClear_Click()
    If Saved Then
        Unload Me
        Me.Show
    End If
End Sub

Private Sub cmdSClose_Click()
    If Saved Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
   
    On Error GoTo err
      
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where Active=1")
    OrderDT = Date
    DeliveryDT = DateAdd("m", 1, Date)
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Private Sub Form_Resize()
On Error Resume Next
   FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Function Saved() As Boolean
    
    Saved = False
    If LV.ListItems.count = 0 Then
        MsgBox "No Items Added To List.", vbInformation
        Exit Function
    End If
    
    For i = 0 To cmbItemID.ListCount - 1
        If Val(cmbItemID.List(i, 3)) > 0 Then
            MsgBox "Not All Of Custom Invoice Items Are Added To The List.", vbInformation
            Exit Function
        End If
    Next i
    
    Call StartTrans(con)
    '----------------------------------------------------------------
    If strOrderNo <> "" Then
        con.Execute "DELETE FROM CustomPList WHERE CustomInvoice='" & _
         strOrderNo & "'"
    End If
    
    'Now Save Items
    Dim strItemCode As String
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            con.Execute "INSERT INTO CustomPList(CustomInvoice,CartonFrom,CartonTo,Qty,AutoItemID,CustomInvoiceItemID,LotNo" & _
             ",BatchNo) VALUES('" & cmbCustomInvoice & "'," & Val(.key) & "," & Val(.ListSubItems(1).Tag) & "," & _
             Val(.SubItems(1)) & "," & Val(.ListSubItems(2).Tag) & "," & Val(.ListSubItems(3).Tag) & ",'" & .SubItems(4) & _
             "','" & .SubItems(5) & "')"
        End With
    Next
      
    con.CommitTrans
    
    Saved = True
    Exit Function

err:
    MsgBox err.Description
End Function

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top
        .Text = LV.SelectedItem.SubItems(iColNo - 1)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > Val(LV.ColumnHeaders(i).Left) And x < (Val(LV.ColumnHeaders(i).Left) + Val(LV.ColumnHeaders(i).Width)) Then
            iColNo = i
            Exit For
        End If
    Next
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(iColNo - 1) = txtEdit.Text
        If iColNo = 6 Then
            If LV.SelectedItem.Index < LV.ListItems.count Then
                LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                iColNo = 5
                Call LV_DblClick
                Exit Sub
            End If
        Else
            iColNo = 6
            Call LV_DblClick
            Exit Sub
        End If
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub
