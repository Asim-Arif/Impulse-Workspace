VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmNewSampleInvoice 
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10635
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
   ScaleHeight     =   8625
   ScaleWidth      =   10635
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
      Height          =   8520
      Left            =   45
      TabIndex        =   14
      Top             =   30
      Width           =   10560
      Begin ComboList.Usercontrol1 cmbCompany 
         Height          =   285
         Left            =   1785
         TabIndex        =   29
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
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
         TextForeColor   =   16711680
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   1785
         Width           =   2250
         _ExtentX        =   3969
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
      Begin ComboList.Usercontrol1 cmbItems 
         Height          =   285
         Left            =   285
         TabIndex        =   34
         Top             =   2340
         Width           =   5250
         _ExtentX        =   9260
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
      Begin VB.TextBox txtEditPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4155
         TabIndex        =   31
         Top             =   3030
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.TextBox txtEditQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5385
         TabIndex        =   30
         Top             =   3030
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Frame FNew 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   285
         TabIndex        =   22
         Top             =   2100
         Width           =   10230
         Begin VB.CommandButton cmdAddItem 
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   9630
            Picture         =   "frmNewSampleInvoice.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Add This Entry To Voucher"
            Top             =   15
            UseMaskColor    =   -1  'True
            Width           =   540
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5265
            Locked          =   -1  'True
            MaxLength       =   150
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   255
            Width           =   930
         End
         Begin VB.TextBox txtPrice 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   6225
            MaxLength       =   150
            TabIndex        =   23
            Top             =   255
            Width           =   930
         End
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   7185
            MaxLength       =   150
            TabIndex        =   7
            Top             =   255
            Width           =   930
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmNewSampleInvoice.frx":0532
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
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   9555
         End
         Begin VB.Label lblAmt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   8145
            TabIndex        =   24
            Top             =   255
            Width           =   1410
         End
         Begin MSForms.ComboBox cmbCatID 
            Height          =   285
            Left            =   0
            TabIndex        =   5
            Top             =   600
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
      End
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5685
         MaxLength       =   150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1125
      End
      Begin VB.TextBox txtOrderNo 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6825
         MaxLength       =   150
         TabIndex        =   3
         Top             =   1800
         Width           =   1470
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
         TabIndex        =   18
         Top             =   0
         Width           =   10560
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New Sample Invoice"
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
            Width           =   10605
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "New Sample Invoice"
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
            Width           =   10590
         End
      End
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9915
         Picture         =   "frmNewSampleInvoice.frx":05E0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Remove This Entry From Voucher"
         Top             =   2715
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   540
      End
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   8295
         TabIndex        =   4
         Top             =   1785
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   529
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
         Format          =   115146755
         CurrentDate     =   37250
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4755
         Left            =   285
         TabIndex        =   9
         Top             =   2715
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   8387
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
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cust ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   5997
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
            Object.Width           =   1499
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
            Object.Width           =   1852
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
         Left            =   4665
         TabIndex        =   15
         Top             =   7785
         Width           =   5280
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   3510
            TabIndex        =   13
            Top             =   195
            Width           =   1665
            Caption         =   "Cancel            "
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
         Begin MSForms.CommandButton cmdSClear 
            Height          =   360
            Left            =   90
            TabIndex        =   11
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
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
         Begin MSForms.CommandButton cmdSClose 
            Height          =   360
            Left            =   1800
            TabIndex        =   12
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Trade Terms :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5685
         TabIndex        =   33
         Top             =   765
         Width           =   1170
      End
      Begin MSForms.ComboBox cmbTradeTerms 
         Height          =   285
         Left            =   6885
         TabIndex        =   32
         Top             =   720
         Width           =   2325
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4101;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Company Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   330
         TabIndex        =   28
         Top             =   765
         Width           =   1410
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
         Left            =   7470
         TabIndex        =   27
         Top             =   7455
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
         TabIndex        =   26
         Top             =   7455
         Width           =   1830
      End
      Begin VB.Label lblTotWords 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2100
         TabIndex        =   21
         Top             =   7455
         Width           =   5385
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmNewSampleInvoice.frx":0B12
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
         TabIndex        =   17
         Top             =   1545
         Width           =   9555
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   8340
         TabIndex        =   16
         Top             =   7455
         Width           =   1500
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2565
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1785
         Width           =   3120
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5503;503"
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
Attribute VB_Name = "frmNewSampleInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String
Dim strDelList As String
Dim iColNo As Integer

Public Sub EditSampleInvoice(strSampleInvoiceNo As String)
    
    On Error GoTo err
    strOrderNo = strSampleInvoiceNo
    strDelList = ""
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT * FROM VSampleInvoice Where CustomInvoice='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        cmbCompany.ID = !CompanyRefID
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        txtOrderNo = ![CustomInvoice] & ""
        OrderDT = ![DT] & ""
        cmbTradeTerms = ![TradeTerms] & ""
        
'        cmbPayment = ![PaymentTerms] & ""
'        cmbTrans = ![TransShipment] & ""
'        cmbPartial = ![PartialShipment] & ""
'        txtPackaging = ![Packaging] & ""
'        DeliveryDT = ![DeliveryDT] & ""
'        txtQuality = !Quality & ""
'        txtInternalRefNo = Val(!InternalRefNo & "")
'        If IsNull(!StampDT) = False Then
'            DTStamp = !StampDT
'        End If
        
        Do Until .EOF
        
            Set ITM = LV.ListItems.add(, ![ItemCode] & "'", .AbsolutePosition)
            
            ITM.Tag = !ItemCode & "" '![CompITemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            ITM.ListSubItems.add , , !ItemCode & "" '![CatID] & ""
            ITM.ListSubItems.add , , ![ItemCode] & ""
            ITM.ListSubItems.add(, , ![ItemName] & "").Tag = "" '![Description] & ""
            
            ITM.ListSubItems.add , , !Unit & ""
            
            ITM.ListSubItems.add(, , Format(![Price] & "", "#0.00")).Tag = 0 '![CustomPrice] & ""
            ITM.ListSubItems.add(, , Val(![Qty] & "")).Tag = 0
            
            ITM.ListSubItems.add(, , Val(![Qty] & "") * Val(![Price])).Tag = ![EntryID] & ""
           
            .MoveNEXT
        Loop
        .Close
    End With
    
    Call GetTotal

    txtQty = 0
'    cmbItems.SetFocus
    
    Set rs = Nothing
    Me.Show
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub chkCust_Click()

    If cmbCatID.MatchFound Then
        Call FillItems
        For i = 1 To LV.ListItems.count
            If chkCust Then
                LV.ListItems(i).SubItems(2) = LV.ListItems(i).key
                LV.ListItems(i).ListSubItems(3) = LV.ListItems(i).ListSubItems(3).key
            Else
                LV.ListItems(i).SubItems(2) = LV.ListItems(i).Tag
                LV.ListItems(i).ListSubItems(3) = LV.ListItems(i).ListSubItems(3).Tag
            End If
        Next i
    End If
    
End Sub

Private Sub cmbCatID_Change()
    Call cmbCatID_Click
End Sub

Private Sub cmbCatID_Click()
    If cmbCatID.MatchFound = False Then Exit Sub
    Call FillItems
End Sub

Private Sub cmbCompany_matched()

    If strCompany = "Tecno" Then Exit Sub
    
    If strOrderNo <> "" Then Exit Sub
    Dim InvoiecNo As String
    InvoiceNo = GetSampleInvoiceNo(cmbCompany.ID)
    If bCourierInvoice Then
        txtOrderNo = strCourierInvoicePrefix & " " & Format(InvoiceNo, "000")
    Else
        txtOrderNo = InvoiceNo
    End If
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    With rs
        '.Open "Select Distinct CatID,CatDesc From VFCustomerCatalog Where CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCatID.Clear
        cmbCatID.AddItem "<All Catagories>"
        'Do Until .EOF
       '     cmbCatID.AddItem .Fields(1) & ""
      '      cmbCatID.List(cmbCatID.ListCount - 1, 1) = .Fields(0) & ""
     '       .MoveNext
    '    Loop
    '    .Close
        .Open "Select TradeTerms,PaymentTerms From ForeignCustomers Where CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbTradeTerms.Text = !TradeTerms & ""
        'cmbPayment.Text = !PaymentTerms & ""
        .Close
    End With
    
    Set rs = Nothing
    If cmbCatID.ListCount > 0 Then cmbCatID.ListIndex = 0
    
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
    
    If strCompany <> "Tecno" Then
        If strOrderNo = "" Then
            If Len(txtOrderNo) > 8 Then
                txtOrderNo = Left(txtOrderNo, 8) & Right(cmbCust.ID, 4)
            Else
                txtOrderNo = Left(txtOrderNo, 8) & Right(cmbCust.ID, 4)
            End If
        End If
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub


Private Sub cmbItems_matched()

    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT Unit,FOB From Items WHERE ItemID='" & cmbItems.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            txtUnit = ""
            txtPrice = ""
            txtPrice.Tag = ""
        Else
            txtUnit = !Unit & ""
            txtPrice = Val(!FOB & "")
            txtPrice.Tag = Val(!FOB & "")
        End If
        .Close
    End With
    Set rs = Nothing
    
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdAddItem_Click()
    
    On Error GoTo err
    
    If Not cmbItems.MatchFound Then
        MsgBox "Invalid Item Selected", vbInformation
        Exit Sub
    ElseIf Val(txtQty) = 0 Then
        MsgBox "Quantity To Order Cannot Be Zero", vbInformation
        Exit Sub
    End If

    Dim ITM As ListItem
    Dim SNo As Integer
    
    If LV.ListItems.count > 0 Then
        SNo = Val(LV.ListItems(LV.ListItems.count).Text) + 1
    Else
        SNo = 1
    End If
    
    Set ITM = LV.ListItems.add(, cmbItems.Text & "'", SNo)
    ITM.Tag = cmbItems.ID 'cmbItems.List(cmbItems.ListIndex, 2)
    
    ITM.ListSubItems.add , , "" 'cmbCustItemID.Text
    ITM.ListSubItems.add , , cmbItems.ID
    ITM.ListSubItems.add(, , cmbItems.Text).Tag = ""  'cmbCustItemID.Text
    
    ITM.ListSubItems.add , , txtUnit
    
    ITM.ListSubItems.add(, , Format(txtPrice, "#0.0000")).Tag = Val(txtPrice.Tag)
    ITM.ListSubItems.add(, , txtQty).Tag = "New"
    ITM.ListSubItems.add , , lblAmt
    
    
    Call GetTotal

    txtQty = 0
    'cmbCustItemID.SetFocus
    cmbItems.SetFocus
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If LV.SelectedItem.ListSubItems(6).Tag <> "New" Then
        'If Val(LV.SelectedItem.ListSubItems(6).Tag) <> Val(LV.SelectedItem.ListSubItems(6)) Then
        If Val(LV.SelectedItem.ListSubItems(6).Tag) <> 0 Then
            MsgBox "Can't Remove This Item.This Item Is In Use In Some Proforma.", vbInformation
            Exit Sub
        End If
    End If
    
    If Val(LV.SelectedItem.ListSubItems(7).Tag) <> 0 Then
        strDelList = strDelList & Val(LV.SelectedItem.ListSubItems(7).Tag) & ","
    End If
    
    LV.ListItems.Remove LV.SelectedItem.Index
    
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
    strOrderNo = ""
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where Active=1")
    OrderDT = ServerDate
    DeliveryDT = DateAdd("m", 2, ServerDate)
    DTStamp = ServerDate
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    cmbItems.ListHeight = 4000
    cmbItems.DropDownOffLine = True
    Call FillCmbs
        
    If strCompany = "Tecno" Then
        Dim lInvoiceNo As Long
        lInvoiceNo = GetSingleLongValue("MAX(CAST(RIGHT(CustomInvoice,LEN(CustomInvoice)-6) AS INT))", "SampleInvoice")
        lInvoiceNo = lInvoiceNo + 1
        txtOrderNo = "TCI/S/" & Format(lInvoiceNo, "0000")
    End If
   
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Public Sub AddNew()
    txtInternalRefNo = GetInternalRefNo
    Me.Show
End Sub

Private Function GetInternalRefNo() As Long
 
    On Error GoTo err
    
    Dim lRefNo As Long
    lRefNo = GetSingleLongValue("MAX(InternalRefNo)", "FCustomerOrders")
     
    If lRefNo = 0 Then
        lRefNo = Val(InputBox("Please Enter the Starting Number for Internal Ref. #", "Internal Reference No.", "596"))
    Else
        lRefNo = lRefNo + 1
    End If
    
    GetInternalRefNo = lRefNo
    
    Exit Function
err:
    MsgBox err.Description

    
End Function

Private Sub Form_Resize()
On Error Resume Next
   FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub

Private Sub lblPymtTerms_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("PaymentTerms", "Payment Terms", "Payment Terms")
End Sub

Private Sub LV_Click()
    If LV.ListItems.count > 0 Then
        Call LV_ItemClick(LV.SelectedItem)
    End If
End Sub

Private Sub LV_DblClick()

    If iColNo = 6 Then
'        Dim strInvoiceNo As String
'        strInvoiceNo = UsedInPInvoice()
'        If strInvoiceNo <> "" Then
'            MsgBox "Article has been used in " & strInvoiceNo & ".Can't Edit Price Now.'", vbInformation
'            Exit Sub
'        End If
        'Editing of Price
        With txtEditPrice
            .Alignment = 1
            .Move LV.Left + LV.ColumnHeaders(6).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(6).Width
            .Text = LV.SelectedItem.SubItems(5)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    ElseIf iColNo = 9 Then
        With txtEditQty
            .Move LV.Left + LV.ColumnHeaders(9).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(9).Width
            .Text = LV.SelectedItem.SubItems(8)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    Else
        With txtEditQty
            .Move LV.Left + LV.ColumnHeaders(7).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(7).Width
            .Text = LV.SelectedItem.SubItems(6)
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    End If
    
End Sub

Private Function UsedInPInvoice() As String

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim lRefID As Long
    lRefID = Val(LV.SelectedItem.ListSubItems(7).Tag)
    With rs
        .Open "Select Top 1 CustomInvoice From CustomInvoiceItems Where RefID In(Select EntryID From FProformaOrders Where OrderEntryID=" & lRefID & ")", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            UsedInPInvoice = ""
        Else
            UsedInPInvoice = .Fields(0)
        End If
        .Close
    End With
    
    Exit Function
    
err:
    MsgBox err.Description
    
End Function

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With cmdDelete
        .Move .Left, LV.Top + Item.Top
        .Visible = True
    End With
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > LV.ColumnHeaders(i).Left And x < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Sub OrderDT_Change()
    DeliveryDT.Value = DateAdd("m", 2, OrderDT.Value)
End Sub

Private Sub txtEditPrice_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        txtEditPrice.Visible = False
    ElseIf KeyAscii = 13 Then
        LV.SelectedItem.SubItems(5) = Val(txtEditPrice)
        LV.SelectedItem.SubItems(7) = Val(txtEditPrice) * Val(LV.SelectedItem.SubItems(6))
        txtEditPrice.Visible = False
    End If

End Sub

Private Sub txtEditQty_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        txtEditQty.Visible = False
    ElseIf KeyAscii = 13 Then
        If iColNo = 9 Then
            If Val(txtEditQty) <= 0 Then
                MsgBox "Invalid Sort No.", vbInformation
                txtEditQty.SetFocus
                Exit Sub
            End If
            LV.SelectedItem.SubItems(8) = Val(txtEditQty)
            If LV.SelectedItem.Index < LV.ListItems.count Then
                Dim iIndex As Integer
                iIndex = LV.SelectedItem.Index + 1
                LV.ListItems(iIndex).Selected = True
                With txtEditQty
                    .Top = LV.SelectedItem.Top + LV.Top
                    .Text = LV.SelectedItem.SubItems(8)
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End With
            Else
                txtEditQty.Visible = False
            End If
            'txtEditQty.Visible = False
        Else
            If Val(txtEditQty) <= 0 Or Val(txtEditQty) < Val(LV.SelectedItem.ListSubItems(6).Tag) Then
                MsgBox "Invalid Qty.", vbInformation
                txtEditQty.SetFocus
                Exit Sub
            End If
            LV.SelectedItem.SubItems(6) = Val(txtEditQty)
            LV.SelectedItem.SubItems(7) = Val(txtEditQty) * Val(LV.SelectedItem.SubItems(5))
            txtEditQty.Visible = False
        End If
    End If
    
End Sub

Private Sub txtPrice_Change()
   lblAmt = Format(Val(txtQty) * Val(txtPrice), "#0.00")
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii, True)
End Sub

Private Sub txtQty_Change()
   lblAmt = Format(Val(txtQty) * Val(txtPrice), "#0.00")
End Sub

Private Sub txtQty_GotFocus()
    cmdAddItem.Default = True
    With txtQty
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtQty_LostFocus()
    cmdAddItem.Default = False
End Sub
Private Sub txtQty_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii)
End Sub


Private Sub GetTotal()

    Dim i As Integer
    Dim total As Double
    total = 0

    For i = 1 To LV.ListItems.count
        total = total + Val(LV.ListItems(i).ListSubItems(7).Text)
    Next i

    lblTotal = Format(total, "#0.00")
    lblTotWords = ConvertInWords(lblTotal, txtCurrency.Text)
    
End Sub



Private Function Saved() As Boolean
            
    On Error GoTo err
        
    Saved = False
    If cmbCompany.MatchFound = False Then
        MsgBox "Please Select Company Name.", vbInformation
        Exit Function
    ElseIf txtOrderNo = "" Then
        MsgBox "No Invoice No.", vbInformation
        txtOrderNo.SetFocus
        Exit Function
    End If
    
    If LV.ListItems.count = 0 Then
        MsgBox "No Items Added to This Sample Invoice", vbInformation
        Exit Function
    End If
    
    Call StartTrans(con)
    
    
    If strOrderNo <> "" Then
        con.Execute "UPDATE SampleInvoice SET DT='" & _
         OrderDT & "',TradeTerms='" & cmbTradeTerms & "',CompanyRefID=" & cmbCompany.ID & _
          " WHERE CustomInvoice='" & strOrderNo & "'"
             
        If strDelList <> "" Then
            strDelList = Left(strDelList, Len(strDelList) - 1)
             
            con.Execute "DELETE FROM SampleInvoiceItems WHERE EntryID In(" & strDelList & ")"
        End If
    Else
        con.Execute "INSERT INTO SampleInvoice(CustCode,Country," & _
         "CustomInvoice,DT,TradeTerms,CompanyRefID) VALUES('" & _
         cmbCust.ID & "','" & cmbCountry & "','" & txtOrderNo & _
         "','" & OrderDT & "','" & cmbTradeTerms & "'," & cmbCompany.ID & ")"
    End If
   
    'Now Save Items
    Dim strItemCode As String
    Dim cmd As New ADODB.Command
    For i = 1 To LV.ListItems.count
        strItemCode = Left(LV.ListItems(i).key, Len(LV.ListItems(i).key) - 1)
SaveNewItem:
            If strCompany <> "All Seas" Then
                'Assign to Customer....
                With cmd
                    Set .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                    .CommandText = "SP_AssItem"
                    .Parameters("@CustCode").Value = cmbCust.ID
                    .Parameters("@CustCountry").Value = cmbCountry
                    .Parameters("@ItemID").Value = LV.ListItems(i).Tag
                    .Execute Saved
                End With
                
            End If
            If Val(LV.ListItems(i).ListSubItems(7).Tag) = 0 Then
                con.Execute "INSERT INTO SampleInvoiceItems (CustomInvoice,ItemCode," & _
                 "Price,Qty) VALUES('" & txtOrderNo & "','" & LV.ListItems(i).Tag & "'," & _
                 Val(LV.ListItems(i).SubItems(5)) & "," & Val(LV.ListItems(i).SubItems(6)) & ")"
            Else
                con.Execute "UPDATE SampleInvoiceItems SET Qty=" & Val(LV.ListItems(i).SubItems(6)) & _
                 ",Price=" & Val(LV.ListItems(i).SubItems(5)) & " WHERE EntryID=" & Val(LV.ListItems(i).ListSubItems(7).Tag)
            End If
    Next
      
    con.CommitTrans
    
    Saved = True
     
    Exit Function

err:
    MsgBox err.Description
    
End Function

Private Sub PrintOrder(strOrderNo As String)

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptFCustOrders.rpt") '
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{CustOrders.OrderNo}='" & strOrderNo & "' ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillCmbs()
    
    Call AddToCombo(cmbTradeTerms, "Description", "TradeTerms", " Order By ID")
'    With cmbTradeTerms
'        .Clear
'        .AddItem "FOB"
'        .AddItem "CIF (Air)"
'        .AddItem "CIF (Sea)"
'        .AddItem "C&F (Air)"
'        .AddItem "C&F (Sea)"
'        .ListIndex = 0
'    End With
    
'    With cmbPartial
'        .AddItem "Allowed"
'        .AddItem "Not Allowed"
'        .ListIndex = 0
'    End With
'
'    With cmbTrans
'        .AddItem "Allowed"
'        .AddItem "Not Allowed"
'        .ListIndex = 0
'    End With
    
    'Call AddToCombo(cmbPayment, "Description", "PaymentTerms")
    
    cmbCompany.ListHeight = 6000
    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID", " Order By CompanyName"
    cmbCompany.ID = 1
    
End Sub

Private Sub FillItems()

    On Error GoTo err
    
    cmbItems.ClearVals
    cmbItems.AddVals con, "'{'+ItemID+'} '+ItemName", "Items", "ItemID"
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
