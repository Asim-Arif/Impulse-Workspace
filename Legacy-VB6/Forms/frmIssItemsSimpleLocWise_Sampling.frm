VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmIssItemsSimpleLocWise_Sampling 
   ClientHeight    =   8280
   ClientLeft      =   1110
   ClientTop       =   60
   ClientWidth     =   11805
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbIssuance 
      Height          =   285
      Left            =   5220
      TabIndex        =   31
      Top             =   1365
      Visible         =   0   'False
      Width           =   5925
      _ExtentX        =   10451
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
   Begin VB.OptionButton Opts 
      Caption         =   "Issuance Against Lot"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   2850
      TabIndex        =   33
      Top             =   1260
      Width           =   2070
   End
   Begin VB.OptionButton Opts 
      Caption         =   "Issuance Against PO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   450
      TabIndex        =   32
      Top             =   1260
      Width           =   2070
   End
   Begin VB.TextBox txtSampleRemarks 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   405
      TabIndex        =   30
      Top             =   7770
      Visible         =   0   'False
      Width           =   5505
   End
   Begin ComboList.Usercontrol1 cmbMaterial 
      Height          =   285
      Left            =   405
      TabIndex        =   4
      Top             =   2025
      Width           =   9855
      _ExtentX        =   17383
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
   Begin ComboList.Usercontrol1 cmbLocation 
      Height          =   285
      Left            =   3165
      TabIndex        =   7
      Top             =   2580
      Width           =   4635
      _ExtentX        =   8176
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
   Begin ComboList.Usercontrol1 cmbStore 
      Height          =   285
      Left            =   405
      TabIndex        =   6
      Top             =   2580
      Width           =   2745
      _ExtentX        =   4842
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
   Begin VB.CheckBox chkSampleIssuance 
      Caption         =   "Sample Issuance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   420
      TabIndex        =   26
      Top             =   7515
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   405
      TabIndex        =   9
      Top             =   3075
      Width           =   10140
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   9900
      TabIndex        =   24
      Top             =   345
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox txtInHand 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9375
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2580
      Width           =   960
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10320
      TabIndex        =   8
      Top             =   2580
      Width           =   825
   End
   Begin VB.TextBox txtUnit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10275
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2025
      Width           =   870
   End
   Begin VB.TextBox txtOrderNo 
      Height          =   285
      Left            =   9405
      TabIndex        =   3
      Top             =   885
      Width           =   1740
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   285
      Left            =   7905
      TabIndex        =   20
      Top             =   885
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   197132291
      CurrentDate     =   38535
   End
   Begin VB.TextBox txtAttn 
      Height          =   300
      Left            =   4845
      TabIndex        =   2
      Top             =   885
      Width           =   3030
   End
   Begin VB.TextBox txtDept 
      Height          =   300
      Left            =   2130
      TabIndex        =   1
      Top             =   885
      Width           =   2700
   End
   Begin VB.TextBox txtIssNo 
      Height          =   300
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   885
      Width           =   1725
   End
   Begin VB.Frame Frame6 
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
      Left            =   6060
      TabIndex        =   14
      Top             =   7485
      Width           =   5280
      Begin MSForms.CommandButton cmdSave 
         Height          =   360
         Left            =   75
         TabIndex        =   17
         Top             =   195
         Width           =   1665
         Caption         =   "Save & New   "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   110
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
         Left            =   3510
         TabIndex        =   16
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
      Begin MSForms.CommandButton cmdSaveClose 
         Height          =   360
         Left            =   1785
         TabIndex        =   15
         Top             =   180
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
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   720
      TabIndex        =   11
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Items"
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
         Left            =   3720
         TabIndex        =   12
         Top             =   135
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Items"
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
         Left            =   3735
         TabIndex        =   13
         Top             =   150
         Width           =   1695
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3990
      Left            =   405
      TabIndex        =   18
      Top             =   3495
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   7038
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S No."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Material ID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Material Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Unit"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Order No."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Type"
         Object.Width           =   1764
      EndProperty
   End
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   405
      TabIndex        =   27
      Top             =   2025
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
   Begin MSForms.ComboBox cmbType 
      Height          =   285
      Left            =   7800
      TabIndex        =   34
      Top             =   2580
      Width           =   1575
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2778;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   168
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   285
      Left            =   2175
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2025
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
   Begin MSForms.Label Label5 
      Height          =   240
      Left            =   405
      TabIndex        =   29
      Top             =   2850
      Width           =   10140
      BackColor       =   11517387
      Caption         =   " Remarks"
      Size            =   "17886;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4_1 
      Height          =   240
      Left            =   405
      TabIndex        =   25
      Top             =   2355
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmIssItemsSimpleLocWise_Sampling.frx":0000
      Size            =   "18944;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbPrinter 
      Height          =   300
      Left            =   3660
      TabIndex        =   23
      Top             =   7785
      Visible         =   0   'False
      Width           =   2250
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3969;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   405
      TabIndex        =   21
      Top             =   1800
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmIssItemsSimpleLocWise_Sampling.frx":00CD
      Size            =   "18944;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   525
      Left            =   10590
      TabIndex        =   10
      Top             =   2850
      Width           =   555
      Caption         =   "Add"
      PicturePosition =   2
      Size            =   "979;926"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   240
      Left            =   405
      TabIndex        =   19
      Top             =   645
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmIssItemsSimpleLocWise_Sampling.frx":01AA
      Size            =   "18944;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmIssItemsSimpleLocWise_Sampling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bSample As Boolean
Private Sub cmbLocation_Change()
    Call cmbLocation_Click
End Sub

Private Sub cmbLocation_Click()
    
End Sub

Private Sub chkSampleIssuance_Click()
    If chkSampleIssuance.Value = vbChecked Then
        txtSampleRemarks.Visible = True
    Else
        txtSampleRemarks.Visible = True
    End If
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()

    cmbMaterial.ClearVals
    cmbMaterial.DropDownOffLine = True
    'cmbMaterial.AddVals con, "ItemName + ' {' + ItemID + '}'", "Items", "ItemID", " Order By ItemID"
    If cmbCust.ID = "0" Then
        If bSample Then
            'cmbMaterial.AddVals con, "'{'+ItemID+'} '+ItemName  ", "Items", "ItemID", " WHERE Sample=1 ORDER BY ItemID"
            cmbMaterial.AddVals con, "'{'+ItemID+'} '+ItemName  ", "Items", "ItemID", " ORDER BY ItemID"
        Else
            cmbMaterial.AddVals con, "'{'+ItemID+'} '+ItemName  ", "Items", "ItemID", " ORDER BY ItemID"
        End If
    Else
        cmbMaterial.AddVals con, "'{' + CompItemID + '} ' + ItemID + ' {' + ISNULL(OCR,'') + '} ' + CompDesc", "VFCustomerCatalog", "CompItemID", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'"
    End If
    
End Sub

Private Sub cmbCust_matched()
    
    If cmbCust.ID = "0" Then
        cmbCountry.Clear
        cmbCountry.AddItem "<Comp>"
        cmbCountry.ListIndex = 0
        cmbCountry.Enabled = False
        Exit Sub
    Else
        cmbCountry.Enabled = True
    End If
    
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
    Set rs = Nothing
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
End Sub

Private Sub cmbLocation_matched()

'    cmbLocation.Tag = GetSingleLongValue("NetQty", "VItemsWithShelfWiseStock", " WHERE ItemID='" & cmbMaterial.ID & "' AND Shelf_RefID=" & cmbLocation.ID)
'
'    txtInHand = cmbLocation.Tag
    
'    Dim lEntryID As Long
'    lEntryID = GetSingleLongValue("MAX(EntryID)", "VRcvItemsSimpleDetail_Placement", " WHERE ItemID='" & cmbMaterial.ID & "' AND Shelf_RefID=" & cmbLocation.ID)
'    txtRemarks.Tag = lEntryID
    'txtRemarks = GetSingleStringValue("Remarks", "RcvItemsSimpleDetail_Placement", " WHERE EntryID IN(SELECT MAX(EntryID) FROM VRcvItemsSimpleDetail_Placement WHERE ItemID='" & cmbMaterial.ID & "' AND Shelf_RefID=" & cmbLocation.ID & ")")
    
End Sub


Private Sub cmbLocation_UnMatched()
    txtInHand = "0"
End Sub

Private Sub cmbMaterial_matched()

    Dim rs As New ADODB.Recordset
    
    
    With rs
        .Open "SELECT Unit,InHand From Items Where ItemID='" & cmbMaterial.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        txtUnit = !Unit & ""
        .Close
        
        Dim lQty As Long
        lQty = GetSingleLongValue("SUM(NetQty)", "VItemsWithShelfWiseStock", " WHERE ItemID='" & cmbMaterial.ID & "'")
        txtUnit.Tag = lQty
        txtInHand = lQty
        
        
    End With
    Set rs = Nothing
    
    cmbStore.ClearVals
    'cmbStore.AddVals con, "StoreName", "VStoreShelfs", "Store_RefID", " WHERE Store_RefID IN(SELECT Store_RefID FROM Users_Stores WHERE UserID=" & UserID & ") AND EntryID IN(SELECT Shelf_RefID FROM VItemsWithShelfWiseStock WHERE ItemID='" & cmbMaterial.ID & "' AND NetQty>0)"
    cmbStore.AddVals con, "StoreName", "VStoreShelfs", "Store_RefID", " WHERE EntryID IN(SELECT Shelf_RefID FROM VItemsWithShelfWiseStock WHERE ItemID='" & cmbMaterial.ID & "' AND NetQty>0)"
    
    'cmbLocation.AddVals con, "StoreName + '{' + RackNo + '} {' + ShelfNo + '} {' + CAST(NetQty AS VARCHAR(50)) + '}'", "VItemsWithShelfWiseStock", "Shelf_RefID", " WHERE ItemID='" & cmbMaterial.ID & "' AND NetQty>0"
    
    'Call AddToCombo(cmbLocation, "StoreName + '(' + RackNo + ') (' + ShelfNo + ') (' + CAST(NetQty AS VARCHAR(50)) + ')'", "VItemsWithShelfWiseStock", " WHERE ItemID='" & cmbMaterial.ID & "' AND NetQty>0", , "Shelf_RefID,NetQty")
'    txtInHand = getMaterialBalance(cmbMaterial.ID, DT)
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbPrinter_Change()
    txtAttn = cmbPrinter.Text
End Sub

Private Sub cmbStore_matched()

    cmbLocation.ClearVals
    cmbLocation.AddVals con, "RackNo + ' {' + ShelfNo + '} {' + CAST(NetQty AS VARCHAR(50)) + '}'", "VItemsWithShelfWiseStock", "Shelf_RefID", " WHERE ItemID='" & cmbMaterial.ID & "' AND NetQty>0 AND Shelf_RefID IN(SELECT EntryID FROM VStoreShelfs WHERE Store_RefID=" & cmbStore.ID & ")"
    
End Sub

Private Sub cmbType_Change()
    Call cmbType_Click
End Sub

Private Sub cmbType_Click()

    cmbLocation.Tag = GetSingleLongValue("NetQty", "VItemsWithShelfWiseStock_Sample", " WHERE ItemID='" & cmbMaterial.ID & "' AND Sample_Type=" & cmbType.ListIndex & " AND Shelf_RefID=" & cmbLocation.ID)
    txtInHand = cmbLocation.Tag
    
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    
    If cmbMaterial.MatchFound = False Then Exit Sub
    If cmbLocation.MatchFound = False Then Exit Sub
    
    If Val(txtQty) > Val(cmbLocation.Tag) Then
        MsgBox "Invalid Qty. Only " & Val(cmbLocation.Tag) & " Available at " & cmbLocation.Text, vbInformation
        Exit Sub
    End If
    
    If cmbType.MatchFound = False Then
        MsgBox "Please Select Type", vbInformation
        Exit Sub
    End If
    
    If Val(txtQty) <= 0 Then Exit Sub
    If Val(txtQty) > Val(txtUnit.Tag) Then
        MsgBox "Can't Issue This Quantity." & vbNewLine & "Only " & txtUnit.Tag & " " & txtUnit & " Is Available In The Stock.", vbInformation
        Exit Sub
    End If
    
    Dim ITM As ListItem
    Set ITM = LV.ListItems.add(, cmbMaterial.ID & "'", LV.ListItems.count + 1)
    ITM.Tag = Val(txtUnit.Tag)
    ITM.ListSubItems.add(, , cmbMaterial.ID).Tag = cmbMaterial.Tag 'Rate For Printers
    ITM.ListSubItems.add , , cmbMaterial.Text
    ITM.ListSubItems.add , , txtUnit
    ITM.ListSubItems.add , , Val(txtQty)
    ITM.ListSubItems.add , , txtOrderNo
    ITM.ListSubItems.add(, , cmbLocation.Text).Tag = cmbLocation.ID
    ITM.ListSubItems.add , , txtRemarks
    ITM.ListSubItems.add(, , cmbType.Text).Tag = cmbType.ListIndex
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Saved Then
        Unload Me
'        LV.ListItems.Clear
'        Call DT_Change
        Me.Show
    End If
End Sub

Private Sub cmdSaveClose_Click()

    If Saved Then Unload Me
    
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
     
    If DT > GetServerDate Then
        MsgBox "Sorry, You cannot issue in future dates.", vbCritical
        Exit Function
    End If
    
    If Trim(txtDept) = "" Then
        MsgBox "Please Feed Department.", vbInformation
        Exit Function
    ElseIf Trim(txtAttn) = "" Then
        MsgBox "Please Feed Attention Person.", vbInformation
        Exit Function
    End If
    
    Dim strSampleRemarks As String
    strSampleRemarks = ""
    If chkSampleIssuance.Value = vbChecked Then
        If txtSampleRemarks.Text = "" Then
            MsgBox "Please Enter Sample Issuance Remarks.", vbInformation
            Exit Function
        Else
            strSampleRemarks = txtSampleRemarks
        End If
    End If
    
    Dim Dated As Date
    Dim dQty As Double, dBal As Double
    Dim lRefID As Long
    Dated = DT
    
    Dim strIssType As String, strVI_RefID As String, strLotNo As String, strProcessID As String
    If bSample Then
        If Opts(0).Value = False And Opts(1).Value = False Then
            MsgBox "Please Select Issuance.", vbInformation
            Exit Function
        End If
        If Opts(0).Value Then
            strIssType = "0"
            strVI_RefID = cmbIssuance.ID
            strLotNo = "NULL"
            strProcessID = "NULL"
        ElseIf Opts(1).Value Then
            strIssType = "1"
            strVI_RefID = "NULL"
            strLotNo = "'" & cmbIssuance.ID & "'"
            strProcessID = GetSingleLongValue("ProcessID", "VRunningLots", " WHERE LotNo='" & cmbIssuance.ID & "'")
        End If
        
    Else
        strIssType = "NULL"
        strVI_RefID = "NULL"
        strLotNo = "NULL"
        strProcessID = "NULL"
    End If
    DT_Change   'To ReAssign Issuance No.
    Call StartTrans(con)
    
    con.Execute "INSERT INTO IssItemsSimple(IssNo,DT,OrderNo,Department,AttnPerson,UserName,ComputerName,SampleIssuance,SampleRemarks,IssuanceType,VI_RefID,LotNo,ProcessID) VALUES('" & _
     txtIssNo & "','" & DT & " " & ServerTime & "','" & txtOrderNo & "','" & txtDept.Text & "','" & _
     txtAttn & "','" & CurrentUserName & "','" & strComputerName & "'," & Abs(bSample) & ",'" & strSampleRemarks & "'," & strIssType & "," & strVI_RefID & "," & strLotNo & "," & strProcessID & ")"
    
    lRefID = con.Execute("Select Max(EntryID) From IssItemsSimple").Fields(0).Value & ""
    
    Dim strMaterialID As String
    Dim lIISD_EntryID As Long
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            strMaterialID = Left(.key, Len(.key) - 1)
            dQty = Val(.SubItems(4))
            dBal = getItemBalance(strMaterialID, Dated)
            dBal = dBal - Val(.SubItems(4))
            theDesc = "Item Issued To " & txtAttn & " Of " & txtDept
            con.Execute "INSERT INTO ItemLedgerSimple(ItemID,DT,Description,Qty,Balance,OrderNo,Location,Remarks) Values('" & _
             strMaterialID & "','" & Dated & "','" & theDesc & "'," & -dQty & "," & dBal & ",'" & .SubItems(5) & "','" & .SubItems(6) & "','" & .SubItems(7) & "')"
            lRefID = Val(con.Execute("Select Max(EntryID) From ItemLedgerSimple").Fields(0).Value & "")
            
            Call adjustItemLedger(strMaterialID, Dated, -dQty)

            con.Execute "INSERT INTO IssItemsSimpleDetail(IssNo,ItemID,Qty,InHand,LedgerRefID,OrderNo,Location,Remarks,Sample_Type) Values('" & _
             txtIssNo & "','" & strMaterialID & "'," & .SubItems(4) & "," & Val(.Tag) - Val(.SubItems(4)) & "," & lRefID & _
             ",'" & .SubItems(5) & "','" & .SubItems(6) & "','" & .SubItems(7) & "'," & Val(.ListSubItems(8).Tag) & ")"
             
            lIISD_EntryID = GetSingleLongValue("MAX(EntryID)", "IssItemsSimpleDetail")
            con.Execute "INSERT INTO IssItemsSimpleDetail_More(IISD_RefID,IssdQty,Shelf_RefID,Remarks) VALUES(" & _
             lIISD_EntryID & "," & Val(.SubItems(4)) & "," & Val(.ListSubItems(6).Tag) & ",'" & .SubItems(7) & "')"
            
            'Update Remarks in RcvItemsSimpleDetail_Placement; Last Entry of selected material @ selected Location...
            Dim lEntryID As Long
            lEntryID = GetSingleLongValue("MAX(EntryID)", "VRcvItemsSimpleDetail_Placement", " WHERE ItemID='" & strMaterialID & "' AND Sample_Type=" & Val(.ListSubItems(8).Tag) & " AND Shelf_RefID=" & Val(.ListSubItems(6).Tag))
            con.Execute "UPDATE RcvItemsSimpleDetail_Placement SET Remarks='" & .SubItems(7) & "' WHERE EntryID=" & lEntryID
            
            Call UpdateItemLocation(strMaterialID, Val(.ListSubItems(6).Tag), dQty, Val(.ListSubItems(8).Tag))
            'con.Execute "Update Material Set InHand=InHand-" & Val(.SubItems(4)) & " Where MaterialID='" & .key & "'"
            
        End With
    Next
    con.CommitTrans
    Saved = True
    Exit Function
err:
    MsgBox err.Description
    
End Function
Private Sub DT_Change()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Max(Cast(Right(IssNo,3) As Int)) From IssItemsSimple Where Convert(Char(8),DT,1)='" & Format(DT, "MM/dd/yy") & "'", con, adOpenForwardOnly, adLockReadOnly
        
        txtIssNo = "ISS-" & Format(DT, "yyMMdd-") & Format(Val(.Fields(0) & "") + 1, "000")
        
        .Close
    End With
    Set rs = Nothing
    
    If cmbMaterial.MatchFound Then
        txtInHand = getItemBalance(cmbMaterial.ID, DT)
        txtUnit.Tag = Val(txtInHand)
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()

    'DT = Date
    'Call DT_Change
    'Call RestrictIssRcvDates(DT)
    DT = Date
    Call DT_Change
    If CurrentUserName = "Administrator" Then
        DT.Enabled = True
    Else
        DT.Enabled = False
    End If
    
    cmbCust.ListHeight = 4000
    cmbCust.AddItem "<Comp>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " WHERE Active=1")
    cmbCust.ID = "0"
    
    cmbMaterial.ListHeight = 2400
    cmbLocation.ListHeight = 2400
'    cmbMaterial.DropDownOffLine = True
'    cmbMaterial.AddVals con, "ItemName + ' {' + ItemID + '}'", "Items", "ItemID", " Order By ItemID"
    
    'cmbDept.AddVals con, "Name", "Departments", "DeptID", " Where Active=1 Order By Name"
    'AddToCombo cmbPrinter, "PrinterName", "Printers", , , "EntryID"
    bSample = False
    
    
    With cmbType
        .AddItem "Master-1"
        .AddItem "Customer"
        .AddItem "IAA"
        .AddItem "Master-2"
    End With
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    Dim TitleBottom As Long

    TitleBottom = cmbMaterial.Top + cmbMaterial.Height + 50 '+ fTitle.Height
 
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Do you want to Remove ?", vbQuestion + vbYesNo) = vbYes Then
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
End Sub

Public Sub ShowMe(Optional p_bSample As Boolean)
    bSample = p_bSample
    If bSample Then
        cmbCust.ID = "0"
        cmbCust.Enabled = False
        Call cmbCountry_Click
        cmbIssuance.Visible = True
        Opts(0).Visible = True
        Opts(1).Visible = True
    End If
    Me.Show
End Sub

Private Sub Opts_Click(Index As Integer)
    
    cmbIssuance.ClearVals
    If Index = 0 Then 'Issuance against PO
        cmbIssuance.AddVals con_ServerSide, "TOP 30000 RecieptID+' {'+VenderName+'}'", "VendIssued INNER JOIN Makers ON VendIssued.VendID=Makers.VendID", "EntryID", " ORDER BY EntryID DESC"
    Else
        cmbIssuance.AddVals con_ServerSide, "TOP 30000 LotNo+' {'+Description+'}'", "VRunningLots", "LotNo"
    End If
    
End Sub
