VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmItems_Rcv_Sampling 
   Caption         =   "Sample Receiving"
   ClientHeight    =   9135
   ClientLeft      =   1170
   ClientTop       =   450
   ClientWidth     =   12375
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
   ScaleHeight     =   9135
   ScaleWidth      =   12375
   WindowState     =   2  'Maximized
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
      Left            =   900
      TabIndex        =   17
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Receiving"
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
         Left            =   4080
         TabIndex        =   18
         Top             =   135
         Width           =   2715
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Receiving"
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
         Left            =   4095
         TabIndex        =   19
         Top             =   150
         Width           =   2715
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8550
      Left            =   75
      TabIndex        =   8
      Top             =   555
      Width           =   12270
      Begin ComboList.Usercontrol1 cmbMaterial 
         Height          =   285
         Left            =   60
         TabIndex        =   0
         Top             =   360
         Width           =   10590
         _ExtentX        =   18680
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
      Begin ComboList.Usercontrol1 cmbIssuance 
         Height          =   285
         Left            =   5040
         TabIndex        =   31
         Top             =   840
         Width           =   7095
         _ExtentX        =   12515
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
         Caption         =   "Receiving Against PO"
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
         Left            =   105
         TabIndex        =   30
         Top             =   765
         Width           =   2250
      End
      Begin VB.OptionButton Opts 
         Caption         =   "Receiving Against Lot"
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
         Left            =   2505
         TabIndex        =   29
         Top             =   765
         Width           =   2250
      End
      Begin VB.TextBox txtTotal 
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
         Height          =   315
         Left            =   6225
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   8040
         Width           =   1560
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   330
         Left            =   11295
         TabIndex        =   7
         Top             =   1425
         Width           =   825
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   285
         Left            =   10665
         TabIndex        =   5
         Top             =   360
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   102563843
         CurrentDate     =   40758
      End
      Begin VB.TextBox txtValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   8220
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txtRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   315
         TabIndex        =   2
         Top             =   8220
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txtLotNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3315
         TabIndex        =   4
         Top             =   8220
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox txtOpQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   105
         TabIndex        =   1
         Top             =   1485
         Width           =   1470
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   435
         Left            =   8220
         TabIndex        =   15
         Top             =   7965
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4545
         ScaleHeight     =   285
         ScaleWidth      =   2160
         TabIndex        =   13
         Top             =   3015
         Visible         =   0   'False
         Width           =   2190
         Begin VB.TextBox txtLocation 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   975
            TabIndex        =   16
            Top             =   -15
            Width           =   1200
         End
         Begin VB.TextBox txtQty 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   -15
            TabIndex        =   14
            Top             =   -15
            Width           =   1000
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   435
         Left            =   9570
         TabIndex        =   12
         Top             =   7965
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   435
         Left            =   10905
         TabIndex        =   11
         Top             =   7965
         Width           =   1275
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5790
         Left            =   60
         TabIndex        =   9
         Top             =   2130
         Width           =   12105
         _ExtentX        =   21352
         _ExtentY        =   10213
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Material"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Issd. Qty"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Balance"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Location"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Type"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Index           =   8
         Left            =   8940
         TabIndex        =   28
         Top             =   1290
         Width           =   420
      End
      Begin MSForms.ComboBox cmbType 
         Height          =   285
         Left            =   8925
         TabIndex        =   27
         Top             =   1485
         Width           =   2295
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4048;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   168
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Index           =   7
         Left            =   10665
         TabIndex        =   25
         Top             =   180
         Width           =   405
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
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
         Left            =   2250
         TabIndex        =   24
         Top             =   8010
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Index           =   5
         Left            =   780
         TabIndex        =   23
         Top             =   8010
         Visible         =   0   'False
         Width           =   405
      End
      Begin MSForms.ComboBox cmbLocation 
         Height          =   285
         Left            =   1590
         TabIndex        =   6
         Top             =   1485
         Width           =   7320
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "12912;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   168
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot No."
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
         Index           =   4
         Left            =   3675
         TabIndex        =   22
         Top             =   8010
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Left            =   1605
         TabIndex        =   21
         Top             =   1275
         Width           =   720
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
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
         Left            =   105
         TabIndex        =   20
         Top             =   1275
         Width           =   300
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
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
         Index           =   0
         Left            =   60
         TabIndex        =   10
         Top             =   150
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmItems_Rcv_Sampling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bDontStartRefreshing As Boolean
Dim iColNo As Integer

Private Sub RefreshList()
    
    On Error GoTo err
    
    If cmbMaterial.MatchFound = False Then Exit Sub
    'txtUOM = GetSingleStringValue("Unit", "Items", " WHERE ItemID='" & cmbMaterial.ID & "'")
    
    'cmbMaterial.Tag = GetSingleBooleanValue("FinStock", "VMaterial", " WHERE MaterialID='" & cmbMaterial.ID & "'")
    
'    If cmbMaterial.Tag Then
'        txtLotNo.Text = ""
'        txtLotNo.Enabled = True
'    Else
        txtLotNo.Text = ""
        txtLotNo.Enabled = False
'    End If
    
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim dTotal As Double
    With rs
'        If CBool(cmbMaterial.Tag) Then
'            .Open "SELECT VendRcvdDetail.*,Location,Qty AS QtyPlaced,QtyIssued FROM VendRcvdDetail INNER JOIN VendRcvdDetailPackaging ON VendRcvdDetail.EntryID=VendRcvdDetailPackaging.Rcv_RefID INNER JOIN VendRcvdDetailPackagingDetail ON VendRcvdDetailPackaging.EntryID=VendRcvdDetailPackagingDetail.VRDP_RefID WHERE LEFT(RcvID,13)='Opening-Stock' AND MaterialID='" & cmbMaterial.ID & "'", con, adOpenForwardOnly, adLockReadOnly
'        Else
        .Open "SELECT RcvItemsSimpleDetail.*,Shelf_RefID,RcvdQty,IssdQty FROM RcvItemsSimpleDetail INNER JOIN RcvItemsSimpleDetail_Placement " & _
         "ON RcvItemsSimpleDetail_Placement.RISD_RefID=RcvItemsSimpleDetail.EntryID WHERE LEFT(RcvNo,13)='Opening-Stock' AND ItemID='" & _
         cmbMaterial.ID & "'", con, adOpenForwardOnly, adLockReadOnly
'        End If
        LV.ListItems.Clear
        
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, , cmbMaterial.Text)
            ITM.Tag = !EntryID & ""
             
            ITM.ListSubItems.add , , Val(!RcvdQty & "")
            ITM.ListSubItems.add , , Val(!IssdQty & "")
            ITM.ListSubItems.add , , Val(!RcvdQty & "") - Val(!IssdQty & "")
            ITM.ListSubItems.add , , !Shelf_RefID & ""
            ITM.ListSubItems.add , , cmbType.List(Val(!Sample_Type & ""))
            dTotal = dTotal + Val(ITM.SubItems(1))
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ToolTipText = ITM.ListSubItems(i).Text
            Next
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    txtTotal = dTotal
    
    Exit Sub

err:
   MsgBox err.Description
   
End Sub

Private Sub cmbIssuance_matched()

    Dim rs As New ADODB.Recordset
    Dim strCondition As String
    If Opts(0).Value Then
        strCondition = " WHERE IssuanceType=0"
        strCondition = strCondition & " AND VI_RefID=" & cmbIssuance.ID
    Else
        strCondition = " WHERE IssuanceType=1"
        strCondition = strCondition & " AND LotNo='" & cmbIssuance.ID & "'"
    End If
    strCondition = strCondition & " AND ItemID='" & cmbMaterial.ID & "'"
    
    Dim lQtyRcvd As Long, lQtyIssd As Long
    lQtyRcvd = GetSingleLongValue("SUM(Qty)", "RcvItemsSimple INNER JOIN RcvItemsSimpleDetail ON RcvItemsSimple.RcvNo=RcvItemsSimpleDetail.RcvNo", strCondition)
    
    With rs
        .Open "SELECT Sample_Type,Qty FROM IssItemsSimple INNER JOIN IssItemsSimpleDetail ON IssItemsSimple.IssNo=IssItemsSimpleDetail.IssNo " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            lQtyIssd = Val(!Qty & "")
            cmbType.ListIndex = Val(!Sample_Type & "")
        Else
            lQtyIssd = 0
        End If
        .Close
    End With
    txtOpQty.Tag = lQtyIssd - lQtyRcvd
    
End Sub

Private Sub cmbLocation_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode.Value = 13 Then
    
        Call cmdAdd_Click
        
    End If
    
End Sub

Private Sub cmbMaterial_matched()
    'Call RefreshList
End Sub

Private Sub cmbVend_matched()
    Call RefreshList
End Sub

Private Sub cmbVend_UnMatched()
    LV.ListItems.Clear
End Sub

Private Sub cmdAdd_Click()

    If cmbMaterial.MatchFound = False Then
        MsgBox "Please Select Material.", vbInformation
        Exit Sub
    ElseIf Val(txtOpQty) <= 0 Or Val(txtOpQty) > Val(txtOpQty.Tag) Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    ElseIf Trim(cmbLocation) = "" Then
        MsgBox "Invalid Location.", vbInformation
        Exit Sub
    ElseIf cmbType.MatchFound = False Then
        MsgBox "Invalid Type.", vbInformation
        Exit Sub
    ElseIf cmbIssuance.MatchFound = False Then
        MsgBox "Please Select Issuance.", vbInformation
        Exit Sub
    
    End If
    
    Dim lLocationCount As Long
    lLocationCount = GetSingleLongValue("SUM(RcvdQty-IssdQty)", "RcvItemsSimpleDetail_Placement", " WHERE Shelf_RefID=" & cmbLocation.List(cmbLocation.ListIndex, 1))
    
'    If (lLocationCount + Val(txtOpQty)) > 20 Then
'        MsgBox "Can't Store more than 20 on selected location. Already Store Qty :" & lLocationCount, vbInformation
'        Exit Sub
'    End If
        
    Dim strIssType As String, strVI_RefID As String, strLotNo As String, strProcessID As String
    
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
        
    Dim strRcvIDToUse As String, iLen As Integer
    strRcvIDToUse = "Opening-Stock-" & cmbType
    iLen = Len(strRcvIDToUse)
    
    Dim strRcvID As String, iRowAffect As Integer
    strRcvID = GetNextRcvNo(DT) 'GetSingleStringValue("RcvNo", "RcvItemsSimple", " WHERE LEFT(RcvNo," & iLen & ")='" & strRcvIDToUse & "' AND DT='" & DT & "'")
        
    con.Execute "INSERT INTO RcvItemsSimple(RcvNo,DT,OrderNo,RcvdBy,RcvdFrom,UserName,ComputerName,Sample,IssuanceType,VI_RefID,LotNo,ProcessID) VALUES('" & _
     strRcvID & "','" & DT & "','','','Opening','" & CurrentUserName & "','" & strComputerName & "',1," & strIssType & "," & strVI_RefID & "," & strLotNo & "," & strProcessID & ")"
        
    'First Check If Data is present in VendRcvdDetail against Opening-Stock RcvID....
    Dim lRcvID As Long, dRate As Double, lPlacementEntryID As Long
    lRcvID = 0 'GetSingleLongValue("EntryID", "RcvItemsSimpleDetail", " WHERE RcvNo='" & strRcvID & "' AND ItemID='" & cmbMaterial.ID & "'")
        
    Call StartTrans(con)
    If lRcvID = 0 Then
        con.Execute "INSERT INTO RcvItemsSimpleDetail(RcvNo,ItemID,Qty,Sample_Type) VALUES('" & _
         strRcvID & "','" & cmbMaterial.ID & "'," & Val(txtOpQty) & "," & cmbType.ListIndex & ")"
            
        lRcvID = GetSingleLongValue("MAX(EntryID)", "RcvItemsSimpleDetail", " WHERE ItemID='" & cmbMaterial.ID & "' AND RcvNo='" & strRcvID & "'")
        lPlacementEntryID = 0
    Else
        con.Execute "UPDATE RcvItemsSimpleDetail SET Qty=Qty+" & Val(txtOpQty) & " WHERE EntryID=" & lRcvID
        lPlacementEntryID = GetSingleLongValue("EntryID", "RcvItemsSimpleDetail_Placement", " WHERE Shelf_RefID=" & cmbLocation.List(cmbLocation.ListIndex, 1) & " AND Rcvd_RefID=" & lRcvID)
    End If
                
    If lPlacementEntryID = 0 Then
        con.Execute "INSERT INTO RcvItemsSimpleDetail_Placement(RISD_RefID,Shelf_RefID,RcvdQty) VALUES(" & _
        lRcvID & "," & cmbLocation.List(cmbLocation.ListIndex, 1) & "," & Val(txtOpQty) & ")"
    Else
        con.Execute "UPDATE MaterialLocationwiseStatus SET RcvdQty=RcvdQty+" & Val(txtOpQty) & " WHERE EntryID=" & lPlacementEntryID
    End If
        
        'con.Execute "UPDATE Material SET InHand=InHand+" & Val(txtOpQty) & "WHERE MaterialID='" & cmbMaterial.ID & "'"
        
    con.CommitTrans
        
    Call RefreshList
        
    cmbLocation = ""
    txtOpQty = ""
    txtLotNo = ""
        
'        If cmbMaterial.Tag Then
'            '''''
'            txtOpQty.SetFocus
'        Else
            'txtUOM = ""
            cmbMaterial.ID = ""
            cmbMaterial.SetFocus
'        End If
        
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Dim strCondition As String
    Set rpt = rptApp.OpenReport(RptPath & "\MaterialPlacementPendencies.rpt")
    strCondition = "ROUND({VVendItemsRcvd.QtyPassed},2)>ROUND({VVendItemsRcvd.QtyPlaced},2) AND LEFT({VVendItemsRcvd.MaterialID},2)<>'32' AND {VVendItemsRcvd.Void}=FALSE"
    If cmbMaterial.ID <> "0" Then
        strCondition = strCondition & " AND {VVendItemsRcvd.MaterialID}='" & cmbMaterial.ID & "'"
    End If
    
    Dim f As New frmPrevRpt
    rpt.FormulaFields.GetItemByName("ForEmail").Text = False
    f.ShowReport strCondition, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdSave_Click()
    
    If Saved Then
        Call RefreshList
    End If
    
End Sub



Private Sub Form_Load()
    
    bDontStartRefreshing = True
    
    'cmbMaterial.AddItem "<All Materials>", "0"
    cmbMaterial.DropDownOffLine = True
    cmbMaterial.AddVals con, "'{' + ItemID + '} ' + ItemName", "Items", "ItemID", " WHERE InActive=0 ORDER BY ItemID"
    'cmbMaterial.ID = "0"
    
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
    Next
    bDontStartRefreshing = False
    
    DT.Value = Date
    
    'Call AddToCombo(cmbLocation, "DISTINCT Location", "MaterialLocationWiseStatus", " WHERE ISNULL(Location,'')<>''")
    Call AddToCombo(cmbLocation, "StoreName+' '+RackNo+' '+ShelfNo", "VStoreShelfs", , False, "EntryID")
    Call RefreshList
    
    With cmbType
        .AddItem "Master-1"
        .AddItem "Customer"
        .AddItem "Mateen"
        .AddItem "Master-2"
    End With
End Sub

Private Sub Form_Resize()
    txtTotal.Left = LV.ColumnHeaders(2).Left
End Sub

Private Sub LV_DblClick()
    
    Exit Sub
    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    
    PicEdit.Move LV.Left + LV.ColumnHeaders(8).Left, LV.Top + LV.SelectedItem.Top
    PicEdit.Visible = True
    txtQty.Text = ""
    txtQty.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PicEdit.Visible = False
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    iColNo = 0
    'Dim i As Integer
    
End Sub

Private Function Saved() As Boolean

    Saved = False
    On Error GoTo err
    
    Call StartTrans(con)
    
    Dim i As Integer
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
        
            If Val(.SubItems(7)) > 0 Then
                con.Execute "INSERT INTO MaterialLocationwiseStatus(Rcvd_RefID,Location,QtyPlaced,UserName,MachineName) VALUES(" & _
                 Val(.Tag) & ",'" & .SubItems(8) & "'," & Val(.SubItems(7)) & ",'" & CurrentUserName & "','" & strComputerName & "')"
                 
                con.Execute "UPDATE Material SET InHand=InHand+" & Val(.SubItems(7)) & "WHERE MaterialID='" & .SubItems(1) & "'"
            End If
            
        End With
    Next
        
    con.CommitTrans
    Saved = True
     
    Screen.MousePointer = vbDefault
    
    Exit Function
    
err:
    MsgBox err.Description
    con.RollbackTrans
    Screen.MousePointer = vbDefault
    
End Function

Private Sub Opts_Click(Index As Integer)
    
    cmbIssuance.ClearVals
    If Index = 0 Then 'Issuance against PO
        cmbIssuance.AddVals con_ServerSide, "TOP 30000 RecieptID+' {'+VenderName+'}'", "VendIssued INNER JOIN Makers ON VendIssued.VendID=Makers.VendID", "EntryID", " WHERE VendIssued.EntryID IN(SELECT VI_RefID FROM IssItemsSimple INNER JOIN IssItemsSimpleDetail ON IssItemsSimple.IssNo=IssItemsSimpleDetail.IssNo WHERE ItemID='" & cmbMaterial.ID & "') ORDER BY EntryID DESC"
    Else
        cmbIssuance.AddVals con_ServerSide, "TOP 30000 LotNo+' {'+Description+'}'", "VRunningLots", "LotNo", " WHERE LotNo IN(SELECT LotNo FROM IssItemsSimple INNER JOIN IssItemsSimpleDetail ON IssItemsSimple.IssNo=IssItemsSimpleDetail.IssNo WHERE ItemID='" & cmbMaterial.ID & "')"
    End If

End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Dim ITM As ListItem
        Set ITM = LV.SelectedItem
        If Val(txtQty) > Val(ITM.SubItems(6)) Or Val(txtQty) = 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        ElseIf Trim(txtLocation) = "" Then
            MsgBox "Invalid Location.", vbInformation
            Exit Sub
        End If
        
        ITM.SubItems(7) = Val(txtQty)
        ITM.SubItems(8) = txtLocation
        
        LV.SetFocus
        
        PicEdit.Visible = False
        
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
End Sub

Private Sub txtOpQty_Change()
    txtValue = Val(txtRate) * Val(txtOpQty)
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtLocation.SetFocus
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtRate_Change()
    txtValue = Val(txtRate) * Val(txtOpQty)
End Sub

Private Function GetNextRcvNo(myDT As Date) As String
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT MAX(CAST (RIGHT(RcvNo,3) AS INT)) FROM RcvItemsSimple WHERE CONVERT(CHAR(8),DT,1)='" & Format(DT, "MM/dd/yy") & "'", con, adOpenForwardOnly, adLockReadOnly
        
        GetNextRcvNo = "RCV-" & Format(DT, "yyMMdd-") & Format(Val(.Fields(0) & "") + 1, "000")
        
        .Close
    End With
    Set rs = Nothing
    
End Function
