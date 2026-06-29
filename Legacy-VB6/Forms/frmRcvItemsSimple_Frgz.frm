VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmRcvItemsSimple_Frgz 
   ClientHeight    =   7560
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
   ScaleHeight     =   7560
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbMaterial 
      Height          =   285
      Left            =   4755
      TabIndex        =   17
      Top             =   1455
      Width           =   5730
      _ExtentX        =   10107
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
      Left            =   3435
      TabIndex        =   27
      Top             =   2010
      Width           =   6840
      _ExtentX        =   12065
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
   Begin VB.TextBox txtMilNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8865
      TabIndex        =   34
      Top             =   2535
      Width           =   1635
   End
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   405
      TabIndex        =   29
      Top             =   1455
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
   Begin VB.TextBox txtBatchNo_Manual 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7245
      TabIndex        =   26
      Top             =   2535
      Width           =   1605
   End
   Begin VB.TextBox txtLotNo_Manual 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5625
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   2535
      Width           =   1605
   End
   Begin VB.CheckBox chkSample 
      Caption         =   "Sample Receiving"
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
      TabIndex        =   32
      Top             =   6870
      Width           =   1995
   End
   Begin ComboList.Usercontrol1 cmbStore 
      Height          =   285
      Left            =   1275
      TabIndex        =   28
      Top             =   2010
      Width           =   2145
      _ExtentX        =   3784
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
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   405
      TabIndex        =   24
      Top             =   2535
      Width           =   5205
   End
   Begin VB.TextBox txtLotNo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3900
      TabIndex        =   22
      Top             =   1455
      Width           =   840
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   9735
      TabIndex        =   21
      Top             =   285
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtInHand 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2010
      Width           =   840
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10290
      TabIndex        =   18
      Top             =   2010
      Width           =   855
   End
   Begin VB.TextBox txtUnit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10500
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1455
      Width           =   645
   End
   Begin VB.TextBox txtOrderNo 
      Height          =   285
      Left            =   8145
      TabIndex        =   13
      Top             =   885
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   285
      Left            =   6645
      TabIndex        =   12
      Top             =   885
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   121110531
      CurrentDate     =   38535
   End
   Begin VB.TextBox txtAttn 
      Height          =   300
      Left            =   4365
      TabIndex        =   11
      Top             =   885
      Width           =   2250
   End
   Begin VB.TextBox txtDept 
      Height          =   285
      Left            =   2130
      TabIndex        =   10
      Top             =   885
      Width           =   2235
   End
   Begin VB.TextBox txtIssNo 
      Height          =   300
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   9
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
      TabIndex        =   3
      Top             =   6810
      Width           =   5280
      Begin MSForms.CommandButton cmdSave 
         Height          =   360
         Left            =   75
         TabIndex        =   6
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
         TabIndex        =   5
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
         Left            =   1800
         TabIndex        =   4
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
      TabIndex        =   0
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receive Items"
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
         TabIndex        =   1
         Top             =   135
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receive Items"
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
         TabIndex        =   2
         Top             =   150
         Width           =   2115
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3915
      Left            =   420
      TabIndex        =   7
      Top             =   2880
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   6906
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S No."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Material ID"
         Object.Width           =   2117
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
         SubItemIndex    =   4
         Text            =   "Qty"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Order #"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Remarks"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Lot No."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Batch No."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Mil No."
         Object.Width           =   1499
      EndProperty
   End
   Begin MSForms.ComboBox cmbSampleRemarks 
      Height          =   315
      Left            =   420
      TabIndex        =   33
      Top             =   7095
      Visible         =   0   'False
      Width           =   5415
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "9551;556"
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
      TabIndex        =   31
      Top             =   2310
      Width           =   10095
      BackColor       =   11517387
      Caption         =   $"frmRcvItemsSimple_Frgz.frx":0000
      Size            =   "17806;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   285
      Left            =   2160
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1455
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
   Begin MSForms.Label Label4 
      Height          =   240
      Left            =   405
      TabIndex        =   23
      Top             =   1755
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmRcvItemsSimple_Frgz.frx":00AF
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
      Left            =   5625
      TabIndex        =   20
      Top             =   885
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
      TabIndex        =   16
      Top             =   1200
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmRcvItemsSimple_Frgz.frx":0181
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
      Height          =   555
      Left            =   10515
      TabIndex        =   15
      Top             =   2310
      Width           =   630
      Caption         =   "Add"
      PicturePosition =   2
      Size            =   "1111;979"
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
      TabIndex        =   8
      Top             =   645
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmRcvItemsSimple_Frgz.frx":0241
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
Attribute VB_Name = "frmRcvItemsSimple_Frgz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bSample As Boolean

Private Sub chkSampleIssuance_Click()
    
End Sub

Private Sub chkSample_Click()
    If chkSample.Value = vbChecked Then
        cmbSampleRemarks.Visible = True
        Call AddToCombo(cmbSampleRemarks, "'{' + IssNo + '} ' + SampleRemarks", "IssItemsSimple", " WHERE SampleIssuance=1 AND SampleRemarks<>'' AND IssNo NOT IN(SELECT IssNo FROM RcvItemsSimple WHERE IssNo IS NOT NULL)", , "IssNo")
    Else
        cmbSampleRemarks.Visible = False
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

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    cmbMaterial.ClearVals
    cmbMaterial.DropDownOffLine = True
    'cmbMaterial.AddVals con, "ItemName + ' {' + ItemID + '}'", "Items", "ItemID", " Order By ItemID"
    If cmbCust.ID = "0" Or cmbCust.ID = "Stock" Then
        cmbMaterial.AddVals con, "'{'+ItemID+'} '+ItemName", "Items", "ItemID", " ORDER BY ItemID"
    Else
        cmbMaterial.AddVals con, "'{' + CompItemID + '} ' + ItemID + ' {' + ISNULL(OCR,'') + '} ' + CompDesc", "VFCustomerCatalog", "CompItemID", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'"
    End If
    'cmbMaterial.AddVals con, "'{' + CompItemID + '} ' + ItemID + ' {' + ISNULL(OCR,'') + '} ' + CompDesc", "VFCustomerCatalog", "CompItemID", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'"
End Sub

Private Sub cmbLocation_matched()

    txtRemarks = GetSingleStringValue("Remarks", "RcvItemsSimpleDetail_Placement", " WHERE EntryID IN(SELECT MAX(EntryID) FROM VRcvItemsSimpleDetail_Placement WHERE ItemID='" & cmbMaterial.ID & "' AND Shelf_RefID=" & cmbLocation.ID & ")")
    
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
'    Call AddToCombo(cmbLocation, "Location + ' (' + CAST(NetQty AS VARCHAR(50)) + ')'", "VItemsWithLocationwiseStock", " WHERE ItemID='" & cmbMaterial.ID & "' AND NetQty>0", , "Location,NetQty")
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
    cmbLocation.AddVals con, "RackNo + ' {' + ShelfNo + '}'", "VStoreShelfs", "EntryID", " WHERE Store_RefID=" & cmbStore.ID & " ORDER BY StoreName,RackNo,ShelfNo"
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    
    If cmbMaterial.MatchFound = False Then Exit Sub
    If cmbLocation.MatchFound = False Then
        MsgBox "Please Select Location.", vbInformation
        Exit Sub
    End If
    If txtLotNo_Manual = "" Then
        MsgBox "Please Enter Lot No.", vbInformation
        Exit Sub
    End If
    
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(LotNo_Manual)", "RcvItemsSimpleDetail", " WHERE LotNo_Manual='" & txtLotNo_Manual & "'")
    If lCount > 0 Then
        MsgBox "Lot No. Exists.", vbInformation
        Exit Sub
    End If
    If txtBatchNo_Manual <> "" Then
        lCount = GetSingleLongValue("COUNT(BatchNo_Manual)", "RcvItemsSimpleDetail", " WHERE BatchNo_Manual='" & txtBatchNo_Manual & "'")
        If lCount > 0 Then
            MsgBox "Batch No. Exists.", vbInformation
            Exit Sub
        End If
    End If
    
    If Val(txtQty) <= 0 Then Exit Sub
'    If Val(txtQty) > Val(txtUnit.Tag) Then
'        MsgBox "Can't Receive This Quantity." & vbNewLine & "Only " & txtUnit.Tag & " " & txtUnit & " Is Available In The Stock.", vbInformation
'        Exit Sub
'    End If
    
    Dim ITM As ListItem
    Set ITM = LV.ListItems.add(, cmbMaterial.ID & "'", LV.ListItems.count + 1)
    ITM.Tag = Val(txtUnit.Tag)
    
    ITM.ListSubItems.add(, , cmbMaterial.ID).Tag = cmbMaterial.Tag 'Rate For Printers
    
    ITM.ListSubItems.add(, , cmbMaterial.Text).Tag = IIf(cmbMaterial.Enabled = False, Val(txtLotNo), "0")
    
    ITM.ListSubItems.add , , txtUnit
    ITM.ListSubItems.add , , Val(txtQty)
    ITM.ListSubItems.add , , txtOrderNo
    
    ITM.ListSubItems.add(, , cmbLocation.Text).Tag = cmbLocation.ID
    ITM.ListSubItems.add , , txtRemarks
    ITM.ListSubItems.add , , txtLotNo_Manual
    ITM.ListSubItems.add , , txtBatchNo_Manual
    ITM.ListSubItems.add , , txtMilNo
    
    txtRemarks = ""
    txtLotNo_Manual = Val(txtLotNo_Manual) + 1
    txtBatchNo_Manual = ""
    
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
        MsgBox "Sorry, You cannot Receive in future dates.", vbCritical
        Exit Function
    End If
    If Trim(txtAttn) = "" Then
        MsgBox "Please feed Received from.", vbInformation
        Exit Function
    End If
    Dim strIssNo As String
    strIssNo = ""
    If chkSample.Value = vbChecked Then
        If cmbSampleRemarks.MatchFound = False Then
            MsgBox "Please Select Sample Issuance.", vbInformation
            Exit Function
        Else
            strIssNo = cmbSampleRemarks.List(cmbSampleRemarks.ListIndex, 1)
        End If
    End If
    
    Dim Dated As Date
    Dim dQty As Double, dBal As Double
    Dim lRefID As Long, lRIS_EntryID As Long
    Dated = DT
    
    Call DT_Change  'to ReAssign Rcv No.
    
    Call StartTrans(con)
    
    con.Execute "INSERT INTO RcvItemsSimple(RcvNo,DT,OrderNo,RcvdBy,RcvdFrom,UserName,ComputerName,IssNo,CustCode) VALUES('" & _
     txtIssNo & "','" & DT & " " & ServerTime & "','" & txtOrderNo & "','" & txtDept.Text & "','" & _
     txtAttn & "','" & CurrentUserName & "','" & strComputerName & "','" & strIssNo & "','" & cmbCust.ID & "')"
    
    lRefID = con.Execute("SELECT Max(EntryID) From RcvItemsSimple").Fields(0).Value & ""
    lRIS_EntryID = lRefID
    
    Dim strMaterialID As String
    Dim lRISD_EntryID As Long
    Dim strNewLotNo As String
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            strMaterialID = Left(.key, Len(.key) - 1)
            dQty = Val(.SubItems(4))
            dBal = getItemBalance(strMaterialID, Dated)
            dBal = dBal + Val(.SubItems(4))
            If Val(.ListSubItems(2).Tag) > 0 Then
                theDesc = GetSingleStringValue("CustomerPO", "VInternalProductionOrderDetail", " WHERE EntryID=" & Val(.ListSubItems(2).Tag))
                theDesc = theDesc & " (" & Val(.ListSubItems(2).Tag) & ")"
            Else
                theDesc = "Item Received From " & txtAttn & " Of " & txtDept
            End If
            
            con.Execute "INSERT INTO ItemLedgerSimple(ItemID,DT,Description,Qty,Balance,OrderNo,Location,Remarks) Values('" & _
             strMaterialID & "','" & Dated & "','" & theDesc & "'," & dQty & "," & dBal & ",'" & .SubItems(5) & "','" & .SubItems(6) & "','" & .SubItems(7) & "')"
             
            lRefID = Val(con.Execute("SELECT MAX(EntryID) FROM ItemLedgerSimple").Fields(0).Value & "")
            
            Call adjustItemLedger(strMaterialID, Dated, dQty)
            strNewLotNo = getNextLotNo
            con.Execute "INSERT INTO RcvItemsSimpleDetail(RcvNo,ItemID,Qty,InHand,LedgerRefID,OrderNo,Location,Remarks,LotNo_Manual,BatchNo_Manual,Mill_Certificate_No) VALUES('" & _
             txtIssNo & "','" & strMaterialID & "'," & .SubItems(4) & "," & Val(.Tag) - Val(.SubItems(4)) & "," & lRefID & _
             ",'" & .SubItems(5) & "','" & .SubItems(6) & "','" & .SubItems(7) & "','" & strNewLotNo & "','" & .SubItems(9) & "','" & .SubItems(10) & "')"
             
            con.Execute "INSERT INTO Lots_List(LotNo,ItemID,Lot_Type,Reference_LotNo,Batch_No,Mill_Certificate_No) VALUES('" & strNewLotNo & "','" & strMaterialID & "',2,'','" & .SubItems(9) & "','" & .SubItems(10) & "')"
            
            lRISD_EntryID = GetSingleLongValue("MAX(EntryID)", "RcvItemsSimpleDetail")
            
            con.Execute "INSERT INTO RcvItemsSimpleDetail_Placement(RISD_RefID,Shelf_RefID,RcvdQty,Remarks) VALUES(" & _
             lRISD_EntryID & "," & Val(.ListSubItems(6).Tag) & "," & Val(.SubItems(4)) & ",'" & .SubItems(7) & "')"
             
            If Val(.ListSubItems(2).Tag) > 0 Then
                Call ReceiveLot(Val(.ListSubItems(2).Tag), strMaterialID, CLng(iFinishedPackagingProcessID), Dated)
            End If
            'con.Execute "Update Material Set InHand=InHand-" & Val(.SubItems(4)) & " Where MaterialID='" & .key & "'"
            
        End With
    Next
    con.CommitTrans
    Saved = True
        
    Call Print_Packed_Opening_Slip(txtIssNo)
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub ReceiveLot(lIPOLotNo As Long, strItemID As String, lProcessID As Long, myDT As Date)
    
    Dim rs As New ADODB.Recordset
    Dim lMakerID As Long, lIssMainEntryID As Long, lIssEntryID As Long, lQty As Long, lLotNo As String
    Dim strOrderNo As String
    With rs
        .Open "SELECT * FROM VVendIssuanceIPOLots WHERE ItemCode='" & strItemID & "' AND ProcessID=" & lProcessID & " AND IPODL_RefID=" & lIPOLotNo, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            lMakerID = Val(!VendID & "")
            lIssMainEntryID = Val(!MainEntryID & "")
            lIssEntryID = Val(!EntryID & "")
            lQty = Val(!IssQty & "")
            lLotNo = Val(!LotNo & "")
            strOrderNo = !OrderNo & ""
        End If
        .Close
    End With
    If lMakerID = 0 Then lMakerID = iFactoryMakerID
    
    con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID) Values(" & lMakerID & _
     ",'" & myDT & " " & ServerTime & "',''," & UserID & "," & lProcessID & "," & lIssMainEntryID & ")"
    
    Dim iRefID As Integer
    iRefID = Val(con.Execute("SELECT MAX(EntryID) FROM VendReceived").Fields(0).Value & "")
    
    'Get The Next Process ID For Current Item...
    Dim lEntryID As Long
    lEntryID = iRefID
   
    Dim strRcvID As String
    Dim lRcvID As Long
    lRcvID = GetSingleLongValue("MAX(CAST(RIGHT(RecieptID,Len(RecieptID)-10) As Int))", "VVendRcvdItems", " WHERE YEAR(DT)=" & year(myDT) & " AND MONTH(DT)=" & Month(myDT) & " AND DAY(DT)=" & Day(myDT) & " AND LEFT(RecieptID,3)='RCV'")
    lRcvID = lRcvID + 1
    strRcvID = "RCV-" & Format(myDT, "ddMMyy") & lRcvID
     
    Dim cmd As New ADODB.Command
    If lQty > 0 Then
        '2. Now If Rcvd After Issued Then
        'Update The RcvdQty In VendIssdDetails....
        '-----------------------------------------------------------------
        With cmd
            .ActiveConnection = con
            .CommandType = adCmdStoredProc
        
            .CommandText = "SP_InsertVendRcvdNew"
            .Parameters("@ItemCode").Value = strItemID
            .Parameters("@ProcessID").Value = lProcessID
            .Parameters("@Qty").Value = lQty
            .Parameters("@WasteQty").Value = 0 'Val(LV.ListItems(i).SubItems(5) & "")
            .Parameters("@iRefID").Value = iRefID
            .Parameters("@VendID").Value = lMakerID
            .Parameters("@LotNo").Value = lLotNo
            .Parameters("@RcvDT").Value = myDT 'LV.ListItems(i).SubItems(7)
            
            .Parameters("@RecID").Value = strRcvID
            .Parameters("@OrderNo").Value = strOrderNo
            .Parameters("@DeliveryRefID").Value = lIssMainEntryID
            .Parameters("@CountedBy").Value = ""
            .Parameters("@Issue_RefID").Value = lIssEntryID
            .Parameters("@Weight").Value = 0 'Val(LV.ListItems(i).ListSubItems(7))
            .Parameters("@ReqAuth").Value = 0 'iAuthRequired 'UserReqAuth
            .Parameters("@UserID").Value = UserID
                        
            .Execute
        End With
        
    End If
   
End Sub


Private Sub DT_Change()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Max(Cast(Right(RcvNo,3) As Int)) From RcvItemsSimple Where Convert(Char(8),DT,1)='" & Format(DT, "MM/dd/yy") & "'", con, adOpenForwardOnly, adLockReadOnly
        
        txtIssNo = "RCV-" & Format(DT, "yyMMdd-") & Format(Val(.Fields(0) & "") + 1, "000")
        
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
    
    'cmbCust.AddItem "<Comp>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " ")
    'cmbCust.ID = "0"
    
    cmbMaterial.ListHeight = 2400
    cmbLocation.ListHeight = 2400
    'cmbMaterial.AddVals con, "ItemID", "Items", "ItemID", " Order By ItemID"
'    cmbMaterial.DropDownOffLine = True
'    cmbMaterial.AddVals con, "ItemName + ' {' + ItemID + '}'", "Items", "ItemID", " ORDER BY ItemID"
    
    txtDept.Text = CurrentUserName
    'cmbDept.AddVals con, "Name", "Departments", "DeptID", " Where Active=1 Order By Name"
    'AddToCombo cmbPrinter, "PrinterName", "Printers", , , "EntryID"
'    Dim rs As New ADODB.Recordset
'    With rs
'        .Open "SELECT * FROM VStoreShelfs ORDER BY StoreName,RackNo,ShelfNo", con, adOpenForwardOnly, adLockReadOnly
'        cmbLocation.Clear
'        Do Until .EOF
'            cmbLocation.AddItem !StoreName & " [" & !RackNo & "] [" & !ShelfNo & "]"
'            cmbLocation.List(.AbsolutePosition - 1, 1) = !EntryID
'            .MoveNEXT
'        Loop
'        .Close
'    End With
    'cmbLocation.AddVals con, "StoreName + ' {' + RackNo + '} {' + ShelfNo + '}'", "VStoreShelfs", "EntryID", " ORDER BY StoreName,RackNo,ShelfNo"
    cmbStore.ListHeight = 2400
    cmbStore.AddVals con, "StoreName", "Stores", "EntryID", " WHERE EntryID IN(SELECT Store_RefID FROM Users_Stores WHERE UserID=" & UserID & ") ORDER BY StoreName"
    'cmbStore.AddVals con, "StoreName", "Stores", "EntryID", " WHERE EntryID=2 ORDER BY StoreName"
    
    txtLotNo_Manual = getNextLotNo
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
        LV.ListItems.Remove LV.SelectedItem.Index
    End If
End Sub

Private Sub txtLotNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtLotNo) > 0 Then
            Dim rs As New ADODB.Recordset
            With rs
                .Open "SELECT * FROM VVendIssuanceIPOLots WHERE IPODL_RefID=" & Val(txtLotNo) & _
                 " AND EntryID=(SELECT MAX(EntryID) FROM VendIssdDetail WHERE IPODL_RefID=" & Val(txtLotNo) & _
                 ") AND ProcessID=" & iFinishedPackagingProcessID & " AND EntryID NOT IN(SELECT Issue_RefID " & _
                 "FROM VendRcvdDetail)", con, adOpenForwardOnly, adLockReadOnly
                
                If .EOF = False Then
                    cmbMaterial.Enabled = False
                    cmbMaterial.ID = !ItemCode & ""
                    txtQty.Locked = True
                    txtQty.Text = Val(!IssQty & "")
                End If
                
            End With
        Else
            cmbMaterial.Enabled = False
            txtQty.Text = ""
        End If
    End If
End Sub

Public Sub ShowMe(Optional p_bSample As Boolean)
    bSample = p_bSample
    If bSample Then
        cmbCust.ID = "0"
        cmbCust.Enabled = False
        Call cmbCountry_Click
'        cmbIssuance.Visible = True
'        Opts(0).Visible = True
'        Opts(1).Visible = True
    End If
    Me.Show
End Sub
