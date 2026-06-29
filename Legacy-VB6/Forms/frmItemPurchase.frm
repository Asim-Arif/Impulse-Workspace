VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMatPurchase 
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10830
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      Height          =   6720
      Left            =   60
      TabIndex        =   18
      Top             =   15
      Width           =   10650
      Begin VB.ComboBox cmbSnos 
         Height          =   1935
         Left            =   7575
         Style           =   1  'Simple Combo
         TabIndex        =   9
         Top             =   2250
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   1335
         TabIndex        =   34
         Top             =   2370
         Visible         =   0   'False
         Width           =   4320
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   1680
         TabIndex        =   33
         Top             =   3390
         Visible         =   0   'False
         Width           =   4320
      End
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   15
         TabIndex        =   30
         Top             =   0
         Width           =   10635
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Item Purchase"
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
            TabIndex        =   31
            Top             =   165
            Width           =   10500
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Item Purchase"
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
            TabIndex        =   32
            Top             =   180
            Width           =   10485
         End
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         DataField       =   "RptTime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6360
         MaxLength       =   150
         TabIndex        =   7
         Top             =   1920
         Width           =   1230
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
         Left            =   9960
         Picture         =   "frmItemPurchase.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Add This Entry To Voucher"
         Top             =   1695
         UseMaskColor    =   -1  'True
         Width           =   540
      End
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   9945
         Picture         =   "frmItemPurchase.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Remove This Entry From Voucher"
         Top             =   2280
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   660
         Left            =   4680
         TabIndex        =   19
         Top             =   5955
         Width           =   5280
         Begin MSForms.CommandButton cmdSClear 
            Height          =   360
            Left            =   90
            TabIndex        =   15
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
            TabIndex        =   16
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmItemPurchase.frx":0A64
            Accelerator     =   83
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
            TabIndex        =   17
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
      End
      Begin VB.TextBox txtDesc 
         DataField       =   "RptTime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   6360
         MaxLength       =   150
         TabIndex        =   3
         Top             =   1335
         Width           =   3585
      End
      Begin VB.TextBox txtInvNo 
         DataField       =   "RptTime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   5160
         MaxLength       =   150
         TabIndex        =   4
         Top             =   1335
         Width           =   1200
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   8520
         MaxLength       =   150
         TabIndex        =   14
         Top             =   5415
         Width           =   1410
      End
      Begin MSComctlLib.ListView LV 
         Height          =   2880
         Left            =   360
         TabIndex        =   12
         Top             =   2250
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   5080
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item ID"
            Object.Width           =   3545
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item"
            Object.Width           =   7488
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Price"
            Object.Width           =   2064
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2293
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtPicker1 
         Height          =   300
         Left            =   6960
         TabIndex        =   0
         Top             =   750
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dddd  dd/MM/yyyy"
         Format          =   22609923
         CurrentDate     =   37250
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         DataField       =   "RptTime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7590
         MaxLength       =   150
         TabIndex        =   8
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmItemPurchase.frx":0BC2
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
         Left            =   330
         TabIndex        =   29
         Top             =   1680
         Width           =   9600
      End
      Begin MSForms.ComboBox cmbVendAccTitle 
         Height          =   330
         Left            =   2025
         TabIndex        =   2
         Top             =   1335
         Width           =   3120
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "5503;582"
         ListWidth       =   8819
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblRcptID 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2115
         TabIndex        =   28
         Top             =   780
         Width           =   1230
      End
      Begin VB.Label lblAmt 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8505
         TabIndex        =   10
         Top             =   1920
         Width           =   1410
      End
      Begin MSForms.ComboBox cmbItem 
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   1920
         Width           =   4335
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "7646;556"
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Purchase Amount :"
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
         Height          =   195
         Index           =   6
         Left            =   6420
         TabIndex        =   27
         Top             =   5190
         Width           =   2070
      End
      Begin MSForms.ComboBox cmbItemID 
         Height          =   315
         Left            =   330
         TabIndex        =   5
         Top             =   1920
         Width           =   1695
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2990;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbVendAccno 
         Height          =   330
         Left            =   330
         TabIndex        =   1
         Top             =   1335
         Width           =   1695
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2990;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Item Purchase # :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   3
         Left            =   360
         TabIndex        =   26
         Top             =   810
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   5
         Left            =   6345
         TabIndex        =   25
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Vender Acc No          Vender Name                                            Bill #                  Description"
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
         Left            =   345
         TabIndex        =   24
         Top             =   1095
         Width           =   9585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Net Purchase Amount :"
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
         Height          =   195
         Index           =   2
         Left            =   6570
         TabIndex        =   23
         Top             =   5715
         Width           =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Discount :"
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
         Height          =   195
         Index           =   1
         Left            =   7665
         TabIndex        =   22
         Top             =   5460
         Width           =   825
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   8520
         TabIndex        =   21
         Top             =   5685
         Width           =   1410
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   8520
         TabIndex        =   20
         Top             =   5160
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmMatPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DontDropDown As Boolean
Dim DontShowList As Boolean

Private Sub cmbItem_LostFocus()
    List2.Visible = False
End Sub

Private Sub cmbSnos_GotFocus()
    Me.KeyPreview = False
End Sub

Private Sub cmbSnos_Keydown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 And cmbSnos <> "" Then
        If FindInCombo(cmbSnos, Trim(cmbSnos), True) Then
            MsgBox "This Serial No Already Added...", vbInformation
            cmbSnos.SelStart = 0
            cmbSnos.SelLength = Len(cmbSnos)
         
        ElseIf cmbSnos.Tag Like "*¤" & Trim(cmbSnos) & "¤*" Then
   
            MsgBox "This Serial No Already Exists...", vbInformation
            cmbSnos.SelStart = 0
            cmbSnos.SelLength = Len(cmbSnos)
        Else
            cmbSnos.AddItem cmbSnos.Text
            cmbSnos = ""
        End If
        If cmbSnos.ListCount = Val(txtQty) Then cmdAddItem.SetFocus
    ElseIf KeyCode = 46 Then
        If FindInCombo(cmbSnos, cmbSnos.Text, True) Then
            If MsgBox("Remove Sno : " & cmbSnos, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            cmbSnos.RemoveItem cmbSnos.ListIndex
        End If
    ElseIf KeyCode = 27 Then
        cmbSnos.Visible = False
        txtQty.SetFocus
    End If
    
End Sub

Private Sub cmbSnos_LostFocus()
Me.KeyPreview = True
End Sub



Private Sub cmbSnos_Validate(Cancel As Boolean)

    If cmbSnos.ListCount <> Val(txtQty) Then
        MsgBox "You Must Enter Serial No's For " & Val(txtQty) & " Items.", vbInformation
        Cancel = True
    End If

End Sub

Private Sub cmbVendAccTitle_LostFocus()
    List1.Visible = False
End Sub

Private Sub List2_Click()

    If UCase(Screen.ActiveControl.Name) <> "CMBITEM" Then
        cmbItem = List2
        List2.Visible = False
        cmbItem.SetFocus
    End If

End Sub

Private Sub CMBITEM_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode = vbKeyDown Then
        If List2.ListIndex < List2.ListCount - 1 Then List2.ListIndex = List2.ListIndex + 1
        KeyCode = 0
        DontShowList = True
    ElseIf KeyCode = vbKeyUp Then
        If List2.ListIndex > 0 Then List2.ListIndex = List2.ListIndex - 1
        KeyCode = 0
        DontShowList = True
    ElseIf KeyCode = 13 And List2.Visible Then
        cmbItem = List2
        List2.Visible = False
        cmbItem.SetFocus
    End If

End Sub

Private Sub CMBITEM_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If DontShowList Then DontShowList = False Else ShowList List2, cmbItem, cmbItem
End Sub

Private Sub List1_Click()

    If UCase(Screen.ActiveControl.Name) <> "CMBVENDACCTITLE" Then
        cmbVendAccTitle = List1
        List1.Visible = False
        cmbVendAccTitle.SetFocus
    End If


End Sub

Private Sub cmbVendAccTitle_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDown Then
   If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex + 1
   KeyCode = 0
   DontShowList = True
ElseIf KeyCode = vbKeyUp Then
   If List1.ListIndex > 0 Then List1.ListIndex = List1.ListIndex - 1
   KeyCode = 0
   DontShowList = True
ElseIf KeyCode = 13 And List1.Visible Then
   cmbVendAccTitle = List1
   List1.Visible = False
   cmbVendAccTitle.SetFocus
End If

End Sub

Private Sub cmbVendAccTitle_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If DontShowList Then DontShowList = False Else ShowList List1, cmbVendAccTitle, cmbVendAccTitle
End Sub



Private Sub cmbItemID_Change()
If DontDropDown Then Exit Sub
    If cmbItemID.MatchFound Then Call cmbItemID_Click
End Sub
Private Sub cmbItemID_Click()
   If cmbItem.ListIndex <> cmbItemID.ListIndex Then cmbItem.ListIndex = cmbItemID.ListIndex
   Call GetItemData
End Sub
Private Sub cmbItemID_Validate(Cancel As Boolean)
    If Not cmbItemID.MatchFound Then cmbItemID.ListIndex = cmbItem.ListIndex
End Sub
Private Sub cmbItem_Change()
'If DontDropDown Then Exit Sub
    If cmbItem.MatchFound Then Call cmbItem_Click
End Sub
Private Sub cmbItem_Click()
    If cmbItemID.ListIndex <> cmbItem.ListIndex Then cmbItemID.ListIndex = cmbItem.ListIndex
End Sub
Private Sub cmbItem_Validate(Cancel As Boolean)
If Not cmbItem.MatchFound Then cmbItem.ListIndex = cmbItemID.ListIndex
End Sub


Private Sub cmbVendAccNo_Change()
    If cmbVendAccno.MatchFound Then Call cmbVendAccNo_Click Else If Not cmbVendAccno <> "" Then cmbVendAccno.DropDown
    
End Sub
Private Sub cmbVendAccNo_Click()
    If cmbVendAccTitle.ListIndex <> cmbVendAccno.ListIndex Then cmbVendAccTitle.ListIndex = cmbVendAccno.ListIndex
    
End Sub
Private Sub cmbVendAccNo_Validate(Cancel As Boolean)
    If Not cmbVendAccno.MatchFound Then cmbVendAccno.ListIndex = cmbVendAccTitle.ListIndex
End Sub
Private Sub cmbVendAccTitle_Change()
    If cmbVendAccTitle.MatchFound Then Call cmbVendAccTitle_Click 'Else If cmbVendAccTitle <> "" Then cmbVendAccTitle.DropDown
End Sub
Private Sub cmbVendAccTitle_Click()
    If cmbVendAccno.ListIndex <> cmbVendAccTitle.ListIndex Then cmbVendAccno.ListIndex = cmbVendAccTitle.ListIndex
End Sub
Private Sub cmbVendAccTitle_Validate(Cancel As Boolean)
If Not cmbVendAccTitle.MatchFound Then cmbVendAccTitle.ListIndex = cmbVendAccno.ListIndex
End Sub

Private Sub cmdAddItem_Click()

If Not cmbItemID.MatchFound Or cmbItemID = "" Then
   MsgBox "Invalid Item Selected", vbInformation
   Exit Sub
ElseIf Val(txtPrice) = 0 Then
   MsgBox "Price Cannot Be Zero", vbInformation
   Exit Sub
ElseIf Val(txtQty) = 0 Then
   MsgBox "Quantity Cannot Be Zero", vbInformation
   Exit Sub
ElseIf cmbSnos.Visible And Val(txtQty) <> cmbSnos.ListCount Then
   MsgBox "You Must Enter Serial No's For " & Val(txtQty) & " Items.", vbInformation
   Exit Sub
End If


Dim itm As ListItem, Snos() As String
Set itm = LV.ListItems.Add(, cmbItemID, cmbItemID)

   If cmbSnos.Visible Then
      ReDim Snos(cmbSnos.ListCount - 1)
      For i = 0 To cmbSnos.ListCount - 1
         Snos(i) = cmbSnos.List(i)
      Next
   End If
   
   itm.ListSubItems.Add(, , cmbItem).Tag = Join(Snos, "¤")
   itm.SubItems(2) = Format(txtPrice, "#0.00")
   itm.SubItems(3) = txtQty
   itm.SubItems(4) = lblAmt
   
   itm.Tag = con.Execute("Select AssetAccNo From Items Where ItemID='" & cmbItemID & "'").Fields(0)
   
cmbItemID.RemoveItem cmbItemID.ListIndex
cmbItem.RemoveItem cmbItem.ListIndex

lblTotal = Format(GetTotal, "#0.00")

cmbItemID = ""
cmbItem = ""

txtQty = ""
txtPrice = ""
cmbSnos.Visible = False

txtQty.Visible = True

cmbItemID.SetFocus

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub CmdDelete_Click()

cmbItemID.AddItem LV.SelectedItem
cmbItem.AddItem LV.SelectedItem.SubItems(1)

LV.ListItems.Remove LV.SelectedItem.Index
lblTotal = Format(GetTotal, "#0.00")
End Sub

Private Sub cmdSClear_Click()
If Saved Then
   If MsgBox("Do You Want To Enter The Payment For This Purchase ?", vbQuestion + vbYesNo) = vbYes Then
      'Dim f As New frmCashPV
      With f
      .cmbAccNo = cmbVendAccno
      .Show
      .SetFocus
      End With
      Unload Me
   Else
      Unload Me
      Me.Show
   End If
End If

End Sub

Private Sub cmdSClose_Click()
If Saved Then
   
   If MsgBox("Do You Want To Enter The Payment For This Purchase ?", vbQuestion + vbYesNo) = vbYes Then
   'Dim f As New frmCashPV
      With f
      .cmbAccNo = cmbVendAccno
      .Show
      .SetFocus
      End With
   End If
      Unload Me

End If
End Sub

Private Sub DtPicker1_Change()
    Call setDateRange(DTPicker1)
    lblRcptID = GetNextStockID(DTPicker1, "PCH")
End Sub

Private Sub Form_Load()
   'CmbSnos is Show or Hidden Depending Upon
   'That Item Has Serial No's Or Not, Tag is Set as Flag
   txtQty.Tag = False
   
   Call AddToCombo(cmbVendAccno, "AccNo", "VVenders", "Order By AccNo")
   Call AddToCombo(cmbVendAccTitle, "AccTitle", "VVenders", "Order By AccNo")
   Call AddToCombo(cmbItemID, "MaterialID", "Material", " Order By MaterialID")
   Call AddToCombo(cmbItem, "MaterialName", "Material", "  Order By MaterialID")
   
   'lblRcptID = "PCH-" & Format(GetMax("Val(Right(EntryID,5))", "StockInOut", "Where left(EntryID,4)='PCH-'") + 1, "00000")
   
   DTPicker1 = ServerDate
Call DtPicker1_Change
   
End Sub

Private Sub Form_Resize()
On Error Resume Next
   FALL.Move (Me.ScaleWidth - FALL.Width) / 2, (Me.ScaleHeight - FALL.Height) / 2
End Sub

Private Sub lblTotal_Change()
   Call txtDiscount_Change
End Sub

Private Sub LV_Click()
    If LV.ListItems.Count > 0 Then
        Call LV_ItemClick(LV.SelectedItem)
    End If
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
        With cmdDelete
            .Move .Left, LV.Top + Item.Top
            .Visible = True
        End With
End Sub

Private Sub txtDiscount_Change()
   lblNetTotal = Format(Val(lblTotal) - Val(txtDiscount), "0.00")
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Sub txtDiscount_Validate(Cancel As Boolean)
   If Val(txtDiscount) >= Val(lblTotal) And Val(lblTotal) <> 0 Then
      MsgBox "Invalid Discount.", vbInformation
      txtDiscount = "0.00"
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
cmdAddItem.Default = Not CBool(txtQty.Tag)
End Sub
Private Sub txtQty_LostFocus()
   cmdAddItem.Default = False
   cmbSnos.Visible = CBool(txtQty.Tag)
   If cmbSnos.Visible Then
      cmbSnos.Clear
      cmbSnos.SetFocus
   End If
End Sub
Private Sub txtQty_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii)
End Sub


Private Function GetTotal(Optional ItemsWithAccNo As String) As Double

Dim i As Integer
Dim total As Double
total = 0

For i = 1 To LV.ListItems.Count
   If ItemsWithAccNo = "" Then
      total = total + Val(LV.ListItems(i).ListSubItems(4).Text)
   Else
      If LV.ListItems(i).Tag = ItemsWithAccNo Then total = total + Val(LV.ListItems(i).ListSubItems(4).Text)
   End If
Next i

GetTotal = total

End Function



Private Function Saved() As Boolean

   Saved = False
   
   If LV.ListItems.Count = 0 Then
      MsgBox "No Items Added.", vbInformation
      Exit Function
   ElseIf Not cmbVendAccno.MatchFound Then
      MsgBox "Invalid Vender Selected", vbInformation
      Exit Function
   End If
   
   Dim INextSno As Double
   
   
   con.BeginTrans
   
      con.Execute "Insert Into StockInOut (EntryID,CustVend,Dated,InvNo,Description,Discount) values('" & lblRcptID & "','" & cmbVendAccno & "',#" & DTPicker1 & "#,'" & txtInvNo & "','" & txtDesc & "'," & Val(txtDiscount) & ")"
      INextSno = getNextSno(DTPicker1, "VInOutDetail")
      For i = 1 To LV.ListItems.Count
         PrevBal = getStockBalance(LV.ListItems(i), DTPicker1)
         con.Execute "Insert into InOutDetail (EntryID,Sno,ItemID,Price,Qty,InHand) Values ('" & _
         lblRcptID & "'," & INextSno & ",'" & LV.ListItems(i) & "'," & Val(LV.ListItems(i).SubItems(2)) & "," & Val(LV.ListItems(i).SubItems(3)) & "," & PrevBal + Val(LV.ListItems(i).SubItems(3)) & ")"
         'With LV.ListItems(i).ListSubItems(1)
         '   If .Tag <> "" Then
         '      Dim Snos() As String
         '      Snos = Split(.Tag, "¤")
         '         For j = LBound(Snos) To UBound(Snos)
         '            Con.Execute "Insert Into ITemSno(ITemID,Sno,EntryID) Values('" & LV.ListItems(i) & "','" & Snos(j) & "','" & lblRcptID & "')"
         '         Next
         '   End If
         'End With
         
         Call adjustItemLedger(LV.ListItems(i), INextSno, Val(LV.ListItems(i).SubItems(3)))
         
         INextSno = INextSno + 1
      Next
      
      If AddVoucher Then
        con.CommitTrans
        Saved = True
      Else
        con.RollbackTrans
        Saved = False
      End If
      
   
Exit Function

err:
   Saved = False
   MsgBox "ERROR : " & err.Number & vbNewLine & err.Description
   con.RollbackTrans
End Function






Private Function AddVoucher() As Boolean

On Error GoTo err
    
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    Dim VchrNo As String
    Dim Amt As Double
    Dim Desc As String
    VchrNo = getNextVchrNo(DTPicker1, "JV")
    INextSno = getNextSno(DTPicker1.value)
    'Remember Sno For Voucher Details
    SnoToUse = INextSno
     
     For i = 1 To LV.ListItems.Count
        Desc = Desc & IIf((i = LV.ListItems.Count) And (i > 1), " and " & LV.ListItems(i).SubItems(1), LV.ListItems(i).SubItems(1) & ",")
     Next
     
      Desc = "Bill # :" & txtInvNo & " Purchase of " & Desc
    
    'Credit Entry
    
    '''''''' Vender Account '''''''''
    
    Amt = Val(lblNetTotal)
    Bal = getBalance(cmbVendAccno, DTPicker1.value)
    Bal = Bal - Amt
    
    con.Execute "insert into Vouchers Values(" & INextSno & ",#" & _
        DTPicker1.value & "#,'" & VchrNo & "','" & cmbVendAccno & _
        "','" & Desc & "',0," & Amt & "," & Bal & ",'',0)"
    
    Call adjustLedger(cmbVendAccno, DTPicker1.value, INextSno, -Amt)
    
    INextSno = INextSno + 1
    
    '''''''' Disacount '''''''''
    Dim DisAcc As String
    DisAcc = DISCOUNTACC
    
    If Val(txtDiscount) = 0 Then GoTo EndOfDiscount     'No Discount Given
    
    Amt = Val(txtDiscount)
    Bal = getBalance(DisAcc, DTPicker1.value)
    Bal = Bal - Amt
    
    'Credit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",#" & _
        DTPicker1.value & "#,'" & VchrNo & "','" & DisAcc & _
        "','Discount On Purchase # :" & lblRcptID & "',0," & Amt & "," & Bal & ",'',0)"
    Call adjustLedger(DisAcc, DTPicker1.value, INextSno, -Amt)
    
    INextSno = INextSno + 1
    
EndOfDiscount:
    
    
    
    
    
    ' Debit Entries
    
    
    
  Dim rs As New ADODB.Recordset
  
  With rs
      .Open "Select Distinct AssetAccNo From Items Where ItemID in (" & GetSelList(LV, ITEM_TEXT, True) & ") and assetAccno<>''", con, adOpenForwardOnly, adLockReadOnly
      ReDim AccBal(.RecordCount, 1)
      
      For i = 0 To .RecordCount - 1
            Bal = getBalance(![AssetAccno], DTPicker1.value)
            Amt = GetTotal(![AssetAccno])
            Bal = Bal + Amt
           
            con.Execute "insert into Vouchers Values(" & INextSno & ",#" & DTPicker1.value & "#,'" & _
            VchrNo & "','" & ![AssetAccno] & "','Purchase From " & cmbVendAccTitle & " Bill # : " & txtInvNo & "'," & _
            Amt & ",0," & Bal & ",'',0)"
             
            
            Call adjustLedger(![AssetAccno], DTPicker1, INextSno, Amt)
            INextSno = INextSno + 1
         .MoveNext
   Next
  End With
    
    AddVoucher = True
    
    Exit Function
err:
    AddVoucher = False
    MsgBox "Error In AddVoucher :" & err.Number & vbNewLine & err.Description
    
End Function


Private Sub GetItemData()
Dim rs As New ADODB.Recordset
With rs
.Open "Select PchPrice,HasSNo From ITems Where ITemID='" & cmbItemID & "'", con, adOpenForwardOnly, adLockReadOnly
   txtPrice = Format(Val(![PchPRice]), "#0.00")
   txtQty.Tag = ![HasSno]

.Close

If CBool(txtQty.Tag) Then cmbSnos.Tag = GetExistingSnosStr(cmbItemID)




End With
Set rs = Nothing
End Sub
