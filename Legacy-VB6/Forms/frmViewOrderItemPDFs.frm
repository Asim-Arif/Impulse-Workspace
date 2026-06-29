VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmViewOrderItemPDFs 
   ClientHeight    =   9285
   ClientLeft      =   -120
   ClientTop       =   -120
   ClientWidth     =   14100
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
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   14100
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
      Height          =   9225
      Left            =   45
      TabIndex        =   2
      Top             =   30
      Width           =   14055
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   435
         Left            =   11640
         TabIndex        =   11
         Top             =   8700
         Width           =   1125
      End
      Begin ComboList.Usercontrol1 cmbOrderNo 
         Height          =   285
         Left            =   4785
         TabIndex        =   10
         Top             =   975
         Width           =   3225
         _ExtentX        =   5689
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
      Begin MSComDlg.CommonDialog CD1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "PDF Files Only|*.pdf"
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   435
         Left            =   12810
         TabIndex        =   8
         Top             =   8700
         Width           =   1125
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Top             =   975
         Width           =   2625
         _ExtentX        =   4630
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
         TabIndex        =   4
         Top             =   0
         Width           =   14040
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "View PDFs"
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
            TabIndex        =   5
            Top             =   165
            Width           =   13950
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "View PDFs"
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
            TabIndex        =   6
            Top             =   180
            Width           =   13935
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   7305
         Left            =   90
         TabIndex        =   9
         Top             =   1335
         Width           =   13860
         _ExtentX        =   24448
         _ExtentY        =   12885
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cust ID"
            Object.Width           =   1931
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   4128
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "P. Unit"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Price"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Qty"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Amount"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Sort"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Delivery"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Finishing"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Stamps"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Remarks"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSForms.ComboBox cmbCatID 
         Height          =   285
         Left            =   0
         TabIndex        =   7
         Top             =   345
         Visible         =   0   'False
         Width           =   1545
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2725;503"
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
         Caption         =   " Customer                                       Country                               Order No"
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
         Left            =   75
         TabIndex        =   3
         Top             =   735
         Width           =   7935
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2685
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   975
         Width           =   2100
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3704;503"
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
Attribute VB_Name = "frmViewOrderItemPDFs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String
Dim strDelList As String
Dim iColNo As Integer

Public Sub EditOrder(TempOrdNo As String)
    
    On Error GoTo err
    strOrderNo = TempOrdNo
    strDelList = ""
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT * FROM VrptOrders WHERE OrderNo='" & strOrderNo & "' ORDER BY SortNo", con, adOpenForwardOnly, adLockReadOnly
        
        
        Dim strDelivery As String
        LV.ListItems.Clear
        Do Until .EOF
            strDelivery = Format(!DeliveryDTItem, "dd-MMM-yyyy")
            
            Set ITM = LV.ListItems.add(, ![CompItemCode] & strDelivery & !QualityItem & !StampsItem & !Remarks & "", .AbsolutePosition)
            
            ITM.Tag = ![CompITemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            ITM.ListSubItems.add , , !ITemID & "" '![CatID] & ""
            ITM.ListSubItems.add , , ![CompITemID] & ""
            ITM.ListSubItems.add(, , ![ItemName] & "").Tag = ![Description] & ""
            
            ITM.ListSubItems.add , , !Unit & ""
            
            ITM.ListSubItems.add(, , Format("0", "#0.0000")).Tag = ![CustomPrice] & ""
            ITM.ListSubItems.add(, , Val(![Qty] & "")).Tag = Val(![Qty] & "") - Val(![InvQty] & "")
            
            ITM.ListSubItems.add(, , "0").Tag = ![ID] & ""
            
            ITM.ListSubItems.add , , Val(![SortNo] & "")
            '----------------------------------------------------------------
            ITM.ListSubItems.add(, , IIf(strDelivery = "01-Jan-1900", "", strDelivery)).Tag = Val(!DeliveryStatus & "")
            ITM.ListSubItems.add , , !QualityItem & ""
            ITM.ListSubItems.add(, , !StampsItem & "").Tag = !Remarks & "" 'GetSingleStringValue("Remarks", "FOrderItems_MoreDetails", " WHERE FOI_RefID=" & Val(!ID & ""))
            ITM.ListSubItems.add , , !Remarks & ""
            If Val(!FOI_RefID & "") > 0 Then
                ITM.ListSubItems(12).Tag = "Added"
            End If
            '----------------------------------------------------------------
            .MoveNEXT
        Loop
        .Close
        
    End With
    
    
    Set rs = Nothing
    
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
    
    cmbOrderNo.ClearVals
    cmbOrderNo.AddVals con, "OrderNo + ' {'+InternalRefNo+'}'", "FCustomerOrders", "OrderNo", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND OrderNo IN(SELECT OrderNo FROM VUnshippedOrderList)"
    
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

Private Sub cmbOrderNo_matched()
    Call EditOrder(cmbOrderNo.ID)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MarkingPlan.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VrptOrders.OrderNo}='" & cmbOrderNo.ID & "'", rpt
End Sub

Private Sub Form_Load()
       
    On Error GoTo err
    strOrderNo = ""
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where Active=1")

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

Private Sub LV_DblClick()
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "FOrderItems_PDF", " WHERE FOI_RefID=" & Val(LV.SelectedItem.ListSubItems(7).Tag))
    
    If lCount > 0 Then
        Call ShowTechDrawing("FOrderItems_PDF", "PDF", " WHERE FOI_RefID=" & Val(LV.SelectedItem.ListSubItems(7).Tag), "Order Item PDF", Me.CD1)
        Exit Sub
    End If
    
    lCount = GetSingleLongValue("COUNT(*)", "FCustomerCatalog", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND CompItemID='" & LV.SelectedItem.SubItems(2) & "' AND ArtworkPDF IS NOT NULL")
     
    If lCount = 0 Then
        MsgBox "No PDF found.", vbInformation
        Exit Sub
    End If
    
   
    Call ShowTechDrawing("FCustomerCatalog", "ArtworkPDF", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND CompItemID='" & LV.SelectedItem.SubItems(2) & "'", "Item PDF", Me.CD1)
    
End Sub


