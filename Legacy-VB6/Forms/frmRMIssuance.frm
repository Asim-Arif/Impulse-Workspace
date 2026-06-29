VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmRMIssuance 
   ClientHeight    =   7560
   ClientLeft      =   1170
   ClientTop       =   120
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
   Begin VB.TextBox txtIssNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   660
      Width           =   2040
   End
   Begin ComboList.Usercontrol1 cmbAuthBy 
      Height          =   285
      Left            =   6300
      TabIndex        =   27
      Top             =   1290
      Width           =   2595
      _ExtentX        =   4577
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
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   3540
      TabIndex        =   26
      Top             =   1290
      Width           =   2760
      _ExtentX        =   4868
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
      Enabled         =   0   'False
      Appearance      =   0
   End
   Begin ComboList.Usercontrol1 cmbIssuance 
      Height          =   285
      Left            =   8895
      TabIndex        =   19
      Top             =   1290
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
   Begin ComboList.Usercontrol1 cmbMaker 
      Height          =   285
      Left            =   405
      TabIndex        =   18
      Top             =   1290
      Width           =   3135
      _ExtentX        =   5530
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
   Begin VB.TextBox txtProcess 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EF&
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   1815
      Width           =   3420
   End
   Begin VB.TextBox txtIssdQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EF&
      Height          =   285
      Left            =   10380
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1815
      Width           =   765
   End
   Begin VB.TextBox txtItemName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EF&
      Height          =   285
      Left            =   1875
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1815
      Width           =   5085
   End
   Begin VB.TextBox txtItemCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EF&
      Height          =   285
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1815
      Width           =   1470
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5925
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2355
      Width           =   795
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5520
      TabIndex        =   15
      Top             =   4650
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print When Saved"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   420
      TabIndex        =   7
      Top             =   6870
      Width           =   2070
   End
   Begin ComboList.Usercontrol1 cmbStore 
      Height          =   285
      Left            =   6750
      TabIndex        =   1
      Top             =   2355
      Width           =   1350
      _ExtentX        =   2381
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
   Begin VB.TextBox txtInHand 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8745
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2355
      Width           =   960
   End
   Begin ComboList.Usercontrol1 cmbMaterial 
      Height          =   285
      Left            =   405
      TabIndex        =   0
      Top             =   2355
      Width           =   5505
      _ExtentX        =   9710
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
   Begin VB.TextBox txtRMQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9735
      TabIndex        =   2
      Top             =   2355
      Width           =   825
   End
   Begin VB.TextBox txtUnit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3E7EF&
      Height          =   285
      Left            =   8115
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2355
      Width           =   615
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
      Left            =   5865
      TabIndex        =   11
      Top             =   6810
      Width           =   5280
      Begin MSForms.CommandButton cmdSave 
         Height          =   360
         Left            =   75
         TabIndex        =   4
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
         TabIndex        =   6
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
         TabIndex        =   5
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
      TabIndex        =   8
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Material"
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
         TabIndex        =   9
         Top             =   135
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Material"
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
         TabIndex        =   10
         Top             =   150
         Width           =   2220
      End
   End
   Begin MSComctlLib.ListView LVRMIss 
      Height          =   4095
      Left            =   435
      TabIndex        =   24
      Top             =   2700
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   7223
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
      NumItems        =   7
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
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Store"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Unit"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Rate"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   4290
      TabIndex        =   30
      Top             =   660
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   503
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
      Format          =   117112835
      CurrentDate     =   37384
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Date :"
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
      Index           =   1
      Left            =   3750
      TabIndex        =   31
      Top             =   690
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Issuance No. :"
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
      Index           =   0
      Left            =   405
      TabIndex        =   28
      Top             =   690
      Width           =   1155
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   2
      Left            =   405
      TabIndex        =   20
      Top             =   1575
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmRMIssuance.frx":0000
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
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   1
      Left            =   405
      TabIndex        =   17
      Top             =   1050
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmRMIssuance.frx":00CA
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
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   0
      Left            =   405
      TabIndex        =   13
      Top             =   2100
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmRMIssuance.frx":0176
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
      Height          =   345
      Left            =   10590
      TabIndex        =   3
      Top             =   2355
      Width           =   555
      Caption         =   "Add"
      PicturePosition =   2
      Size            =   "979;609"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmRMIssuance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strIssNo As String
Dim strDelList As String

Private Sub cmbIssuance_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM VVendIssued WHERE EntryID=" & cmbIssuance.ID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtItemCode = !ITemID & ""
            txtItemName = !ItemName & ""
            txtProcess = !Description & ""
            txtIssdQty = Val(!TotalIssQty & "")
        Else
            txtItemCode = "-"
            txtItemName = "-"
            txtProcess = "-"
            txtIssdQty = "-"
        End If
        .Close
    End With
        
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbmaker_matched()

    cmbIssuance.ClearVals
    If strCompany = "Tecno" Then
        cmbIssuance.AddVals con, "VendIssued.RecieptID + ' (' + CAST(IPODL_RefID AS VARCHAR(50)) + ')'", "VendIssued INNER JOIN VendIssdDetail ON VendIssued.EntryID=VendIssdDetail.RefID", "VendIssued.EntryID", " WHERE VendID=" & cmbMaker.ID & " ORDER BY VendIssued.EntryID DESC"
    Else
        cmbIssuance.AddVals con_ServerSide, "TOP 10000 RecieptID", "VendIssued", "EntryID", " WHERE VendID=" & cmbMaker.ID & " ORDER BY EntryID DESC"
    End If
    
    If cmbMaker.ID = iFactoryMakerID Then
        cmbEmp.Enabled = True
    Else
        cmbEmp.Enabled = False
    End If
    
End Sub

Private Sub cmbMaterial_matched()
    
    Dim iRMGroup As Integer, iForgingCount As Integer
    iRMGroup = GetSingleLongValue("GroupID", "RM", " WHERE RMID=" & cmbMaterial.ID)
    iForgingCount = GetSingleLongValue("COUNT(Group_ID)", "RMGroupIDsForForging", " WHERE Group_ID=" & iRMGroup)
    
    Dim rs As New ADODB.Recordset
    If iForgingCount > 0 Then
        cmbStore.ClearVals
        cmbStore.Enabled = False
        With rs
            .Open "SELECT RequiresMaleGrinding,Unit,QtyInStock,Rate,AnnealingStock,MachiningStock,MakerRate,MaleGrindingStock FROM RM WHERE RMID='" & cmbMaterial.ID & "'", con, adOpenForwardOnly, adLockReadOnly
            txtPrice = Val(!MakerRate & "")
            txtUnit = !Unit & ""
            txtPrice.Tag = "1"  'Forging
            
            If !RequiresMaleGrinding Then
                txtInHand.Tag = Val(!MaleGrindingStock & "")
                txtInHand = Val(!MaleGrindingStock & "")
                txtUnit.Tag = Val(!MaleGrindingStock & "")
            Else
                txtInHand.Tag = Val(!MachiningStock & "")
                txtInHand = Val(!MachiningStock & "")
                txtUnit.Tag = Val(!MachiningStock & "")
            End If
            .Close
        End With
    Else
        cmbStore.Enabled = True
        cmbStore.ClearVals
        cmbStore.AddVals con, "StoreName", "VRMStatusStorewise", "StoreName", " WHERE RMID=" & cmbMaterial.ID & " AND NetQty>0"
        txtPrice.Tag = "0"
        
    
        With rs
            .Open "SELECT Unit,QtyInStock,MakerRate FROM RM WHERE RMID='" & cmbMaterial.ID & "'", con, adOpenForwardOnly, adLockReadOnly
         
            txtPrice = Val(!MakerRate & "")
            txtUnit = !Unit & ""
            '        txtUnit.Tag = Val(!QtyInStock & "")
            txtInHand.Tag = Val(!QtyInStock & "")   'Total Qty from All Stores....
        
            .Close
        
        End With
        Set rs = Nothing
        
    End If
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub cmbPrinter_Change()
    txtAttn = cmbPrinter.Text
End Sub

Private Sub cmbStore_matched()

    Dim dStoreQty As Double
    dStoreQty = GetSingleDoubleValue("NetQty", "VRMStatusStorewise", " WHERE RMID=" & cmbMaterial.ID & " AND StoreName='" & cmbStore.Text & "'")
    txtUnit.Tag = dStoreQty
    txtInHand = dStoreQty
    
End Sub

Private Sub cmdAdd_Click()
    
    On Error GoTo err
    
    If cmbMaterial.MatchFound = False Then Exit Sub
    If Val(txtRMQty) <= 0 Then Exit Sub
    
    Dim lSF_RefID As Long
    lSF_RefID = 0 'GetSingleLongValue("SF_RefID", "RM", " WHERE RMID=" & cmbMaterial.ID)
    
    If lSF_RefID = 0 Then   'If Regual Material then Apply check, No Check for SemiFinish materials...
        If Val(txtRMQty) > Val(txtUnit.Tag) Then
            MsgBox "Can't Issue This Quantity." & vbNewLine & "Only " & txtUnit.Tag & " " & txtUnit & " Is Available In The Store.", vbInformation
            Exit Sub
        End If
    End If
    
    Dim ITM As ListItem
    Set ITM = LVRMIss.ListItems.add(, cmbMaterial.ID & "'", LVRMIss.ListItems.count + 1)
    ITM.Tag = Val(txtInHand.Tag)
    ITM.ListSubItems.add(, , cmbMaterial.ID).Tag = cmbMaterial.Tag ''Material Type
    ITM.ListSubItems.add(, , cmbMaterial.Text).Tag = ""
    ITM.ListSubItems.add(, , cmbStore.Text).Tag = txtPrice.Tag
    ITM.SubItems(4) = txtUnit
    ITM.SubItems(5) = Val(txtRMQty)
    ITM.SubItems(6) = Val(txtPrice)
    
    cmbStore.ClearVals
    cmbMaterial.ID = ""
    txtPrice = ""
    txtUnit = ""
    txtQty = ""
    txtInHand = ""
    txtInHand.Tag = ""
    cmbMaterial.SetFocus
    
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
    End If
End Sub

Private Sub cmdSaveClose_Click()

    If Saved Then Unload Me
    
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
    If cmbMaker.MatchFound = False Then
        MsgBox "Please Select Maker.", vbInformation
        Exit Function
    End If
    
    If cmbMaker.ID <> iFactoryMakerID Then
        If cmbIssuance.MatchFound = False Then
            MsgBox "Please Select Issuance.", vbInformation
            Exit Function
        End If
    End If
    
    Dim strIssNo As String
    strIssNo = getIssuanceNo
    
    If strIssNo <> txtIssNo Then
        MsgBox txtIssNo & " is already used, Software will use " & strIssNo & " for this entry.", vbInformation
    End If
    
    Call StartTrans(con)
    con.Execute "INSERT INTO RawMaterialIssuance(IssNo,DT,UserName,ComputerName,VI_RefID,IssdToEmpID,AuthorizedByEmpID) VALUES('" & _
         strIssNo & "','" & DTPicker1.Value & "','" & CurrentUserName & "','" & strComputerName & "'," & Val(cmbIssuance.ID) & _
         ",'" & cmbEmp.ID & "','" & cmbAuthBy.ID & "')"
      
    Dim strMaterialID As String, dRate As Double
    Dim bRequiresMaleGrinding As Boolean
    For i = 1 To LVRMIss.ListItems.count
        With LVRMIss.ListItems(i)
            strMaterialID = GetSingleStringValue("RMID1", "RM", " WHERE RMID=" & Val(.key))
            dQty = Val(.SubItems(5))
            'If Val(.ListSubItems(3).Tag) = 0 Then
                'dRate = GetSingleDoubleValue("Rate", "RM", " WHERE RMID=" & Val(.key))
                dRate = Val(.SubItems(6))
                '-------------------------------------------------------------------------
                con.Execute "INSERT INTO RawMaterialIssuanceDetail(IssNo,RMID1,Qty,InHand,StoreName,Price) VALUES('" & _
                 strIssNo & "','" & strMaterialID & "'," & dQty & "," & Val(.Tag) - dQty & ",'" & .SubItems(3) & _
                 "'," & dRate & ")"
                 
                If Val(.ListSubItems(3).Tag) = 1 Then 'Forging
                    bRequiresMaleGrinding = GetSingleBooleanValue("RequiresMaleGrinding", "RM", " WHERE RMID=" & Val(.key))
                    If bRequiresMaleGrinding Then
                        con.Execute "UPDATE RM SET QtyInStock=QtyInStock-" & dQty & ",MaleGrindingStock=MaleGrindingStock-" & dQty & " WHERE RMID=" & Val(.key)
                    Else
                        con.Execute "UPDATE RM SET QtyInStock=QtyInStock-" & dQty & ",MachiningStock=MachiningStock-" & dQty & " WHERE RMID=" & Val(.key)
                    End If
                Else
                    con.Execute "UPDATE RM SET QtyInStock=QtyInStock-" & dQty & " WHERE RMID=" & Val(.key)
                End If
            'End If
        End With
    Next
    con.CommitTrans
    Saved = True
    
    If chkPrint.Value = vbChecked Then
        Dim f As New frmPrevRpt
        Dim rpt As CRAXDDRT.Report
        Set rpt = rptApp.OpenReport(RptPath & "\IssuanceNote.rpt")
        f.ShowReport "{IssItemsSimple.IssNo}='" & strIssNo & "' ", rpt
    End If
    
    Exit Function
err:
    MsgBox err.Description
    
End Function


Private Sub DTPicker1_Change()
    txtIssNo = getIssuanceNo
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    'DT = Date
    'Call DT_Change
    'Call RestrictIssRcvDates(DT)
'    DT = Date
'    Call DT_Change
'    If CurrentUserName = "Administrator" Then
'        DT.Enabled = True
'    Else
'        DT.Enabled = False
'    End If
    cmbMaker.ListHeight = 4000
    cmbMaker.AddVals con, "VendID1 + ' ' + VenderName", "Makers", "VendID"
    
    cmbIssuance.ListHeight = 4000
    cmbMaterial.ListHeight = 4000
    
    
    cmbMaterial.AddVals con, "RMName + ' {' + RMID1 + '}'", "RM", "RMID", " WHERE InActive=0 ORDER BY RMID"
    
    cmbEmp.ListHeight = 4000
    cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name", "Employees", "EmpID", " ORDER BY EmpID"
    
    cmbAuthBy.ListHeight = 4000
    cmbAuthBy.AddVals con, "'{' + EmpID + '} ' + Name", "Employees", "EmpID", " ORDER BY EmpID"
    

    'AddToCombo cmbPrinter, "PrinterName", "Printers", , , "EntryID"
    strIssNo = ""
    strDelList = ""
    
    DTPicker1.Value = Date
    Call DTPicker1_Change
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    Dim TitleBottom As Long

    TitleBottom = cmbMaterial.Top + cmbMaterial.Height + 50 '+ fTitle.Height
 
End Sub

Public Sub EditIssuance(p_strIssNo As String)

    On Error GoTo err
    strIssNo = p_strIssNo
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM IssItemsSimple WHERE IssNo='" & strIssNo & "'", con, adOpenForwardOnly, adLockReadOnly
        DT.Value = Format(!DT, "dd-MMM-yyyy")
        txtAttn = !AttnPerson & ""
        txtIssNo = !IssNo
        cmbRestaurant.ID = Val(!Restaurant_RefID & "")
        .Close
        .Open "SELECT * FROM VIssItemsSimpleDetail WHERE IssNo='" & strIssNo & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !RMID & "'", LV.ListItems.count + 1)
            ITM.Tag = Val(Val(!QtyInStock & ""))
            ITM.ListSubItems.add(, , !RMID).Tag = ""
            ITM.ListSubItems.add(, , !RMName + "{ " & !RMID1 & "}").Tag = ""
            ITM.ListSubItems.add(, , !StoreName & "").Tag = !EntryID
            ITM.SubItems(4) = !Unit & ""
            ITM.SubItems(5) = Val(!Qty & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Me.Show
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_DblClick()
    If LV.ListItems.count = 0 Then Exit Sub
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(6).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(6).Width
        .Text = Val(LV.SelectedItem.SubItems(5))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If Val(LV.SelectedItem.ListSubItems(3).Tag) <> 0 Then
            strDelList = strDelList & Val(LV.SelectedItem.ListSubItems(3).Tag) & ","
        End If
        LV.ListItems.Remove LV.SelectedItem.Index
    End If
    
End Sub

Private Sub LVRMIss_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVRMIss.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Do you want to Remove ?", vbQuestion + vbYesNo) = vbYes Then
            LVRMIss.ListItems.Remove LVRMIss.SelectedItem.Index
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtEdit) <= 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        LV.SelectedItem.SubItems(5) = Val(txtEdit)
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Function getIssuanceNo() As String

    Dim rs As New ADODB.Recordset, strIssNo As String
    With rs
        .Open "SELECT MAX(CAST(RIGHT(IssNo,3) AS INT)) FROM RawMaterialIssuance WHERE DT='" & DTPicker1.Value & "'", con, adOpenForwardOnly, adLockReadOnly
        
        strIssNo = "ISS-" & Format(DTPicker1.Value, "yyMMdd-") & Format(Val(.Fields(0) & "") + 1, "000")
        
        .Close
    End With
    Set rs = Nothing
    getIssuanceNo = strIssNo
    
End Function
