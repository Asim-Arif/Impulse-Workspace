VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmNewCustomPListM_New 
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11820
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
   Icon            =   "frmNewCustomPListM_New.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   11820
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
      Height          =   9300
      Left            =   45
      TabIndex        =   16
      Top             =   30
      Width           =   11745
      Begin VB.TextBox txtEdit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7335
         TabIndex        =   31
         Top             =   3855
         Visible         =   0   'False
         Width           =   1000
      End
      Begin ComboList.Usercontrol1 cmbItem 
         Height          =   285
         Left            =   3390
         TabIndex        =   4
         Top             =   1695
         Width           =   6225
         _ExtentX        =   10980
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
         TextBackColor   =   12640511
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
         Left            =   615
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
      Begin VB.TextBox txtFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   615
         MaxLength       =   150
         TabIndex        =   1
         Text            =   "1"
         Top             =   1695
         Width           =   1005
      End
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
         Left            =   11100
         Picture         =   "frmNewCustomPListM_New.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Add This Entry To List."
         Top             =   1470
         UseMaskColor    =   -1  'True
         Width           =   540
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   5
         Top             =   1695
         Width           =   705
      End
      Begin VB.TextBox txtRen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   10290
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1695
         Width           =   750
      End
      Begin VB.TextBox txtTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1590
         MaxLength       =   150
         TabIndex        =   2
         Top             =   1695
         Width           =   915
      End
      Begin VB.TextBox txtCartonQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2490
         MaxLength       =   150
         TabIndex        =   3
         Top             =   1695
         Width           =   915
      End
      Begin VB.CommandButton cmdNextCrtn 
         Caption         =   "Next Crtn"
         Height          =   525
         Left            =   75
         Picture         =   "frmNewCustomPListM_New.frx":0974
         TabIndex        =   28
         Top             =   1455
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4140
         MaxLength       =   150
         TabIndex        =   9
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
         TabIndex        =   20
         Top             =   0
         Width           =   11745
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
            ForeColor       =   &H00000000&
            Height          =   390
            Index           =   8
            Left            =   60
            TabIndex        =   21
            Top             =   165
            Width           =   11595
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
            TabIndex        =   22
            Top             =   180
            Width           =   11580
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
         Left            =   11115
         Picture         =   "frmNewCustomPListM_New.frx":0EA6
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Remove This Entry From List."
         Top             =   1965
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   540
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6255
         Left            =   615
         TabIndex        =   11
         Top             =   1965
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   11033
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Carton From"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Carton To"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qnty"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Description"
            Object.Width           =   5715
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Lot No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
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
         Left            =   4020
         TabIndex        =   17
         Top             =   8535
         Width           =   7020
         Begin MSForms.CommandButton cmdSavenStay 
            Height          =   360
            Left            =   3510
            TabIndex        =   32
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Stay"
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   116
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdCancel 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   5220
            TabIndex        =   15
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
            TabIndex        =   13
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
            TabIndex        =   14
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmNewCustomPListM_New.frx":13D8
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
         Left            =   615
         TabIndex        =   30
         Top             =   1455
         Width           =   10425
      End
      Begin MSForms.ComboBox cmbItemID 
         Height          =   285
         Left            =   8085
         TabIndex        =   29
         Top             =   795
         Visible         =   0   'False
         Width           =   2340
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4128;503"
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
         Caption         =   " to"
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
         Height          =   285
         Index           =   3
         Left            =   3015
         TabIndex        =   27
         Top             =   8730
         Visible         =   0   'False
         Width           =   285
      End
      Begin MSForms.ComboBox cmbCustomInvoice 
         Height          =   285
         Left            =   5175
         TabIndex        =   10
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
         Left            =   8175
         TabIndex        =   26
         Top             =   8250
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
         Left            =   615
         TabIndex        =   25
         Top             =   8250
         Width           =   1830
      End
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   600
         TabIndex        =   24
         Top             =   8850
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
      Begin VB.Label lblTotWords 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2430
         TabIndex        =   23
         Top             =   8250
         Width           =   5760
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
         Left            =   615
         TabIndex        =   19
         Top             =   795
         Width           =   6270
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   9045
         TabIndex        =   18
         Top             =   8250
         Width           =   1995
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2385
         TabIndex        =   8
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
Attribute VB_Name = "frmNewCustomPListM_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String
Dim iColNo As Integer

Public Sub EditOrder(TempOrdNo As String)
    
    On Error GoTo err
    strOrderNo = TempOrdNo
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim str As String
    Dim Qty As String
    
    
    With rs
        '.Open "Select * From VCustomPList Where PInvoice='" & strOrderNo & "' Order By Cast(CartonFrom as int)", con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT * FROM VrptCustomPList WHERE CustomInvoice='" & strOrderNo & "' Order By Cast(CartonFrom as int)", con, adOpenForwardOnly, adLockReadOnly
        
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        
        With cmbCustomInvoice
            .Clear
            .AddItem strOrderNo
            '.List(0, 1) = rs![PInvoiceID] & ""
            .ListIndex = 0
            .Locked = True
        End With
        
        LV.ListItems.Clear
        Do Until .EOF
            If Val(![CartonTo] & "") = 0 Then
                str = Format(![CartonFrom] & "", "00")
                Qty = ![Qty] & " Pcs"
            Else
                str = Format(![CartonFrom] & "", "00 to ") & Format(![CartonTo] & "", "00")
                Qty = ![Qty] & " Pcs Each"
            End If
            
            
            'Set ITM = LV.ListItems.Add(, Val(![CartonFrom] & "") & "'", Str)
            
            Set ITM = LV.ListItems.add(, , !CartonFrom & "")
            
            For i = 0 To cmbItemID.ListCount - 1
                'If Val(cmbItemID.List(i, 1)) = ![AutoItemID] Then
                'If Val(cmbItemID.List(i, 2)) = ![OrderItemID] Then
                If cmbItemID.List(i, 0) = !ITemID Then
                    ITM.Tag = i   'To Remember Item Index In The ComboBox.
                    cmbItemID.List(i, 3) = Val(cmbItemID.List(i, 3)) - ((!CartonTo - !CartonFrom + 1) * Val(!Qty & ""))
                    Exit For
                End If
            Next i
            ITM.ListSubItems.add , , !CartonTo & ""
            ITM.ListSubItems.add , , Qty
            ITM.ListSubItems(1).Tag = Val(![Qty] & "")
            'ITM.ListSubItems.Add(, , ![ITemID] & "").Tag = ![AutoItemID] & ""
            ITM.ListSubItems.add(, , ![ItemCode] & "").Tag = ![OrderItemID] & ""
            'ITM.ListSubItems.Add(, , ![FullItemName] & "").Tag = ![CustomInvoiceItemID] & ""
            ITM.ListSubItems.add(, , ![Description] & "").Tag = ![OrderItemID] & ""
            
            ITM.ListSubItems.add , , !LotNo & ""
            ITM.ListSubItems.add , , !BatchNo & ""
            'ITM.ListSubItems.Add , , Qty
            'ITM.ListSubItems(1).Tag = Val(txtTo)
            
            'ITM.ListSubItems.Add(, , cmbItemID).Tag = cmbItemID.List(cmbItemID.ListIndex, 1)    'AutoItemID
            
            'ITM.ListSubItems.Add(, , cmbItemDesc).Tag = cmbItemID.List(cmbItemID.ListIndex, 2)    'CustomInvoiceItemID
                
            'cmbItemID.List(cmbItemID.ListIndex, 3) = cmbItemID.List(cmbItemID.ListIndex, 3) - lTotalQty
            
            .MoveNEXT
        Loop
        .Close
    End With
    txtQty = 0
    
    
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
        .CommandText = "SP_CustomInvoiceForPList"
        
        .Parameters("@Customer").Value = cmbCountry.Tag
        
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
    End With
    
    cmbCustomInvoice.Clear
    With rs
        Do Until .EOF
            cmbCustomInvoice.AddItem .Fields(0) & ""
            'cmbCustomInvoice.List(.AbsolutePosition - 1, 1) = .Fields(1) & ""
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

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub

Private Sub cmbCustomInvoice_Change()
    Call cmbCustomInvoice_Click
End Sub

Private Sub cmbCustomInvoice_Click()

    On Error GoTo err
    If cmbCustomInvoice.MatchFound = False Then Exit Sub
    
    If strOrderNo = "" Then LV.ListItems.Clear
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "SELECT ItemID,SUM(Qty) AS Qty,MAX(CustomInvoiceItemsEntryID) AS CustomInvoiceItemsEntryID FROM VrptCustomInvoiceDetail WHERE CustomInvoice='" & cmbCustomInvoice.Text & "' GROUP BY ItemID ORDER BY ItemID", con, adOpenForwardOnly, adLockReadOnly
        cmbItemID.Clear

        Do Until .EOF
            cmbItemID.AddItem !ITemID & ""
            cmbItemID.List(.AbsolutePosition - 1, 1) = !ITemID & ""
            cmbItemID.List(.AbsolutePosition - 1, 2) = Val(!CustomInvoiceItemsEntryID & "")
            cmbItemID.List(.AbsolutePosition - 1, 3) = Val(!Qty & "")
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    cmbItem.ClearVals
    cmbItem.AddVals con, "'{' + ItemID + '} {' + CompItemID + '} ' + Description", "VrptCustomInvoiceDetail", "MAX(CustomInvoiceItemsEntryID)", " WHERE CustomInvoice='" & cmbCustomInvoice.Text & "' GROUP BY ItemID,CompItemID,Description ORDER BY ItemID"
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbItemDesc_Change()
    Call cmbItemDesc_Click
End Sub

Private Sub cmbItemDesc_Click()
    If cmbItemDesc.MatchFound = False Then Exit Sub
    cmbItemID.ListIndex = cmbItemDesc.ListIndex
End Sub

Private Sub cmbItem_GotFocus()
    cmbItem.ID = ""
End Sub

Private Sub cmbItem_matched()

    Dim lInvEntryID As Long
    lInvEntryID = cmbItem.ID
    
    Dim i As Integer
    For i = 0 To cmbItemID.ListCount - 1
        If cmbItemID.List(i, 2) = lInvEntryID Then
            cmbItemID.ListIndex = i
            Exit For
        End If
    Next
    
End Sub

Private Sub cmbItemID_Change()
    Call cmbItemID_Click
End Sub

Private Sub cmbItemID_Click()
    If cmbItemID.MatchFound = False Then Exit Sub
    txtRen = cmbItemID.List(cmbItemID.ListIndex, 3)
End Sub

Private Sub cmdNextCrtn_Click()

    Dim isAdded As Boolean
    
    For i = 1 To LV.ListItems.count
        If Val(LV.ListItems(i).Text) = Val(txtFrom) Then
            isAdded = True
            Exit For
        End If
    Next i
    
    If isAdded Then txtFrom = Val(txtFrom) + 1
    
End Sub

Private Sub cmdSavenStay_Click()

    If Saved Then
        Dim strInvNo As String
        strInvNo = cmbCustomInvoice
        Call EditOrder(strInvNo)
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdAddItem_Click()
    
    On Error GoTo err
    
    
    If Val(txtQty) > Val(cmbItemID.List(cmbItemID.ListIndex, 3)) Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    End If
    
    If Not cmbItemID.MatchFound Or cmbItemID = "" Then
        MsgBox "Invalid Item Selected", vbInformation
        Exit Sub
    
    ElseIf Val(txtQty) = 0 Then
        MsgBox "Quantity Cannot Be Zero", vbInformation
        Exit Sub
    End If

    If Val(txtFrom) = 0 Or Val(txtTo) = 0 Then
        MsgBox "Invalid Carton #", vbInformation
        Exit Sub
    End If
    
    Dim ITM As ListItem
    Dim str As String
    Dim Qty As String
    Dim lTotalQty As Long
    Dim lNextCarton As Long
    
    'If Val(txtTo) = 0 Then
        str = Format(Val(txtFrom), "00")
        Qty = Val(txtQty) & " Pcs"
        lTotalQty = Val(txtQty)
        lNextCarton = Val(txtFrom) + 1
    'Else
    '    Str = Format(Val(txtFrom), "00") & " to " & Format(Val(txtTo), "00")
    '    Qty = Val(txtQty) & " Pcs Each"
    '    lTotalQty = Val(txtQty) * (Val(txtTo) - Val(txtFrom) + 1)
    '    lNextCarton = Val(txtTo) + 1
    'End If
    
    If lTotalQty > Val(cmbItemID.List(cmbItemID.ListIndex, 3)) Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    Set ITM = LV.ListItems.add(, , str)
    
    If err.Number = "35602" Then
        MsgBox "This Item Is Present Againg This Carton #.", vbInformation
        Exit Sub
    End If
    On Error GoTo 0
    
    ITM.Tag = cmbItemID.ListIndex
    
    ITM.ListSubItems.add , , Val(txtTo)
    ITM.ListSubItems.add , , Qty
    ITM.ListSubItems(1).Tag = Val(txtCartonQty)
    ITM.ListSubItems(2).Tag = Val(txtFrom) & cmbItemID & "'"
    
    ITM.ListSubItems.add(, , cmbItemID).Tag = cmbItemID.List(cmbItemID.ListIndex, 1)     'AutoItemID
    
    ITM.ListSubItems.add(, , cmbItem.Text).Tag = cmbItemID.List(cmbItemID.ListIndex, 2)    'CustomInvoiceItemID
        
    cmbItemID.List(cmbItemID.ListIndex, 3) = cmbItemID.List(cmbItemID.ListIndex, 3) - lTotalQty
    
    ITM.ListSubItems.add
    ITM.ListSubItems.add , , GetSingleStringValue("BatchNo", "CustomInvoiceItems", " WHERE EntryID=" & Val(cmbItemID.List(cmbItemID.ListIndex, 2)))
    
    txtQty = 0
    
    'txtFrom = lNextCarton
    'txtTo = ""
    Call cmbItemID_Click
    txtFrom.SetFocus
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    With LV.SelectedItem
        If InStr(.Text, "to") Then
            cmbItemID.List(.Tag, 3) = cmbItemID.List(.Tag, 3) + Val(.SubItems(1)) * Val(.ListSubItems(1).Tag - Val(.ListSubItems(2).Tag) + 1)
        Else
'           cmbItemID.List(.Tag, 3) = cmbItemID.List(.Tag, 3) + Val(.SubItems(1))
            cmbItemID.List(.Tag, 3) = cmbItemID.List(.Tag, 3) + (Val(.SubItems(1)) - Val(.Text) + 1) * Val(.SubItems(2))
        End If
        LV.ListItems.Remove .Index
       
    End With
    
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
      
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'HICO'")
    OrderDT = Date
    DeliveryDT = DateAdd("m", 1, Date)
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    
    cmbItem.ListHeight = 3000
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub



Private Sub LV_Click()
    If LV.ListItems.count > 0 Then
        Call LV_ItemClick(LV.SelectedItem)
    End If
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With cmdDelete
        .Move .Left, LV.Top + Item.Top
        .Visible = True
    End With
End Sub

Private Sub LV_DblClick()
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    If iColNo = 7 Then
        If strCompany <> "IAA" Then
            MsgBox "Please Edit Batch No. from Custom Invoice.", vbInformation
            Exit Sub
        End If
    End If
    
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


Private Sub txtCartonQty_Change()
    txtQty = (Val(txtTo) - Val(txtFrom) + 1) * Val(txtCartonQty)
End Sub

Private Sub txtCartonQty_GotFocus()
    With txtCartonQty
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtFrom_GotFocus()
    With txtFrom
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtQty_GotFocus()
    With txtQty
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
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
    
'    For i = 0 To cmbItemID.ListCount - 1
'        If Val(cmbItemID.List(i, 3)) > 0 Then
'            MsgBox "Not All Of Custom Invoice Items Are Added To The List.", vbInformation
'            Exit Function
'        End If
'    Next i
    
    Call StartTrans(con)
    '----------------------------------------------------------------
    If strOrderNo <> "" Then
        con.Execute "DELETE FROM CustomPList WHERE CustomInvoice='" & cmbCustomInvoice & "'"
    End If
    
    'Now Save Items
    Dim strItemCode As String
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            con.Execute "INSERT INTO CustomPList(CustomInvoice,CartonFrom,CartonTo,Qty,OrderItemID,LotNo,BatchNo) VALUES('" & _
             cmbCustomInvoice & "'," & Val(.Text) & "," & Val(.SubItems(1)) & "," & Val(.ListSubItems(1).Tag) & _
             ",'" & Val(.ListSubItems(4).Tag) & "','" & .SubItems(5) & "','" & .SubItems(6) & "')"
            
        End With
    Next
      
    con.CommitTrans
    
    Saved = True
    Exit Function

err:
    MsgBox err.Description
End Function

Private Sub txtTo_GotFocus()
    With txtTo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTo_Validate(cancel As Boolean)
    If Val(txtTo) < Val(txtFrom) Then
        MsgBox "Invalid Entry", vbInformation
        cancel = True
        Exit Sub
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(iColNo - 1) = txtEdit.Text
        If iColNo = 7 Then
            If LV.SelectedItem.Index < LV.ListItems.count Then
                LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                iColNo = 6
                Call LV_DblClick
                Exit Sub
            End If
        Else
            iColNo = 7
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

