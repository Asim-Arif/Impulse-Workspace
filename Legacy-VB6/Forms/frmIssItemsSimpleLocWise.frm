VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmIssItemsSimpleLocWise 
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
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   9900
      TabIndex        =   21
      Top             =   345
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox txtInHand 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8775
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1980
      Width           =   960
   End
   Begin ComboList.Usercontrol1 cmbMaterial 
      Height          =   285
      Left            =   405
      TabIndex        =   17
      Top             =   1425
      Width           =   10755
      _ExtentX        =   18971
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
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9720
      TabIndex        =   18
      Top             =   1980
      Width           =   825
   End
   Begin VB.TextBox txtUnit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1980
      Width           =   870
   End
   Begin VB.TextBox txtOrderNo 
      Height          =   285
      Left            =   9405
      TabIndex        =   13
      Top             =   885
      Width           =   1740
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   285
      Left            =   7905
      TabIndex        =   12
      Top             =   885
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   82444291
      CurrentDate     =   38535
   End
   Begin VB.TextBox txtAttn 
      Height          =   300
      Left            =   4845
      TabIndex        =   11
      Top             =   885
      Width           =   3030
   End
   Begin VB.TextBox txtDept 
      Height          =   300
      Left            =   2130
      TabIndex        =   10
      Top             =   885
      Width           =   2700
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
         Left            =   1785
         TabIndex        =   4
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
      TabIndex        =   0
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
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   150
         Width           =   2220
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4455
      Left            =   420
      TabIndex        =   7
      Top             =   2355
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   7858
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Order No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSForms.Label Label4 
      Height          =   240
      Left            =   405
      TabIndex        =   23
      Top             =   1755
      Width           =   10140
      BackColor       =   11517387
      Caption         =   $"frmIssItemsSimpleLocWise.frx":0000
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
   Begin MSForms.ComboBox cmbLocation 
      Height          =   285
      Left            =   405
      TabIndex        =   22
      Tag             =   "PaymentTerms"
      Top             =   1980
      Width           =   7515
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "13256;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbPrinter 
      Height          =   300
      Left            =   2265
      TabIndex        =   20
      Top             =   6945
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
      Caption         =   " Material"
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
      TabIndex        =   15
      Top             =   1740
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
      TabIndex        =   8
      Top             =   645
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmIssItemsSimpleLocWise.frx":00C6
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
Attribute VB_Name = "frmIssItemsSimpleLocWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbMaterial_matched()
    cmbMaterialID.ID = cmbMaterial.ID
End Sub

Private Sub cmbMaterialID_matched()

    Dim rs As New ADODB.Recordset
    
    cmbMaterial.ID = cmbMaterialID.ID
    
    With rs
        .Open "SELECT Unit,InHand From Items Where ItemID='" & cmbMaterialID.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        txtUnit = !Unit & ""
        txtUnit.Tag = Val(!InHand & "")
        txtInHand = Val(!InHand & "")
        .Close
        
    End With
    Set rs = Nothing
    
    Call AddToCombo(cmbLocation, "Location + ' (' + CAST(NetQty AS VARCHAR(50)) + ')'", "VItemsWithLocationwiseStock", " WHERE ItemID='" & cmbMaterialID.ID & "' AND NetQty>0", , "Location,NetQty")
'    txtInHand = getMaterialBalance(cmbMaterialID.ID, DT)
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbPrinter_Change()
    txtAttn = cmbPrinter.Text
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    
    If cmbMaterial.MatchFound = False Then Exit Sub
    If cmbLocation.MatchFound = False Then Exit Sub
    If Val(txtQty) > Val(cmbLocation.List(cmbLocation.ListIndex, 2)) Then
        MsgBox "Invalid Qty. Only " & Val(cmbLocation.List(cmbLocation.ListIndex, 2)) & " Available at " & cmbLocation.List(cmbLocation.ListIndex, 1), vbInformation
        Exit Sub
    End If
    If Val(txtQty) <= 0 Then Exit Sub
    If Val(txtQty) > Val(txtUnit.Tag) Then
        MsgBox "Can't Issue This Quantity." & vbNewLine & "Only " & txtUnit.Tag & " " & txtUnit & " Is Available In The Stock.", vbInformation
        Exit Sub
    End If
    
    Dim ITM As ListItem
    Set ITM = LV.ListItems.add(, cmbMaterialID.ID & "'", LV.ListItems.count + 1)
    ITM.Tag = Val(txtUnit.Tag)
    ITM.ListSubItems.add(, , cmbMaterialID.ID).Tag = cmbMaterial.Tag 'Rate For Printers
    ITM.SubItems(2) = cmbMaterial.Text
    ITM.SubItems(3) = txtUnit
    ITM.SubItems(4) = Val(txtQty)
    ITM.SubItems(5) = txtOrderNo
    ITM.SubItems(6) = cmbLocation.List(cmbLocation.ListIndex, 1)
    
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
    
    Dim Dated As Date
    Dim dQty As Double, dBal As Double
    Dim lRefID As Long
    Dated = DT
    
    DT_Change   'To ReAssign Issuance No.
    Call StartTrans(con)
    
    con.Execute "INSERT INTO IssItemsSimple(IssNo,DT,OrderNo,Department,AttnPerson,UserName,ComputerName) Values('" & _
     txtIssNo & "','" & DT & " " & ServerTime & "','" & txtOrderNo & "','" & txtDept.Text & "','" & _
     txtAttn & "','" & CurrentUserName & "','" & strComputerName & "')"
    
    lRefID = con.Execute("Select Max(EntryID) From IssItemsSimple").Fields(0).Value & ""
    
    Dim strMaterialID As String
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            strMaterialID = Left(.key, Len(.key) - 1)
            dQty = Val(.SubItems(4))
            dBal = getItemBalance(strMaterialID, Dated)
            dBal = dBal - Val(.SubItems(4))
            theDesc = "Item Issued To " & txtAttn & " Of " & txtDept
            con.Execute "INSERT INTO ItemLedgerSimple(ItemID,DT,Description,Qty,Balance,OrderNo,Location) Values('" & _
             strMaterialID & "','" & Dated & "','" & theDesc & "'," & -dQty & "," & dBal & ",'" & .SubItems(5) & "','" & .SubItems(6) & "')"
            lRefID = Val(con.Execute("Select Max(EntryID) From ItemLedgerSimple").Fields(0).Value & "")
            
            Call adjustItemLedger(strMaterialID, Dated, -dQty)

            con.Execute "INSERT INTO IssItemsSimpleDetail(IssNo,ItemID,Qty,InHand,LedgerRefID,OrderNo,Location) Values('" & _
             txtIssNo & "','" & strMaterialID & "'," & .SubItems(4) & "," & Val(.Tag) - Val(.SubItems(4)) & "," & lRefID & _
             ",'" & .SubItems(5) & "','" & .SubItems(6) & "')"
             
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
    
    If cmbMaterialID.MatchFound Then
        txtInHand = getItemBalance(cmbMaterialID.ID, DT)
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
    cmbMaterial.ListHeight = 2400
    cmbMaterialID.ListHeight = 2400
   
    cmbMaterialID.AddVals con, "ItemID", "Items", "ItemID", " Order By ItemID"
    cmbMaterial.AddVals con, "ItemName + ' {' + ItemID + '}'", "Items", "ItemID", " Order By ItemID"

    'cmbDept.AddVals con, "Name", "Departments", "DeptID", " Where Active=1 Order By Name"
    'AddToCombo cmbPrinter, "PrinterName", "Printers", , , "EntryID"
    
    
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
