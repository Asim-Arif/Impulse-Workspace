VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmIPOLotsInspection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inspection Processes"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   14445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   270
      TabIndex        =   22
      Top             =   7605
      Visible         =   0   'False
      Width           =   1470
   End
   Begin MSComCtl2.DTPicker DTInspection 
      Height          =   390
      Left            =   8700
      TabIndex        =   5
      Top             =   7170
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   52822019
      CurrentDate     =   40357
   End
   Begin VB.TextBox txtRemarks 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   2730
      TabIndex        =   4
      Top             =   7170
      Width           =   5940
   End
   Begin VB.TextBox txtDisposation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   270
      TabIndex        =   3
      Top             =   7170
      Width           =   2415
   End
   Begin VB.TextBox txtOrderNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   12360
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2010
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7290
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox txtLotQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   270
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2010
   End
   Begin VB.TextBox txtItemName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   2295
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1440
      Width           =   10035
   End
   Begin VB.TextBox txtItemCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   270
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2010
   End
   Begin VB.CommandButton cmdReceive 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11310
      TabIndex        =   6
      Top             =   7605
      Width           =   1470
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   12870
      TabIndex        =   7
      Top             =   7605
      Width           =   1470
   End
   Begin VB.TextBox txtLotBarcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   510
      Left            =   6000
      TabIndex        =   0
      Top             =   540
      Width           =   2010
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3960
      Left            =   270
      TabIndex        =   2
      Top             =   2775
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   6985
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Attributes Description"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "AQL"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Sample Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Reject On"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Test Spec. No."
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Act. Rej."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Date Inspected"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   7
      Left            =   8640
      TabIndex        =   21
      Top             =   6810
      Width           =   1650
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "General Remarks"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   6
      Left            =   2745
      TabIndex        =   20
      Top             =   6810
      Width           =   1920
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Material Disposation"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   4
      Left            =   270
      TabIndex        =   19
      Top             =   6810
      Width           =   2310
   End
   Begin MSForms.ComboBox cmbProcessNo 
      Height          =   390
      Left            =   2295
      TabIndex        =   1
      Top             =   2280
      Width           =   12075
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "21299;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Customer Order"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   5
      Left            =   12360
      TabIndex        =   17
      Top             =   1035
      Width           =   2010
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4740
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   2790
      Width           =   4035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   3
      Left            =   2280
      TabIndex        =   15
      Top             =   1905
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Lot Qty"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   2
      Left            =   255
      TabIndex        =   13
      Top             =   1905
      Width           =   2010
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   1
      Left            =   2295
      TabIndex        =   11
      Top             =   1035
      Width           =   1230
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Item Code"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   0
      Left            =   270
      TabIndex        =   9
      Top             =   1035
      Width           =   2010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IPO Lot"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   6390
      TabIndex        =   8
      Top             =   135
      Width           =   1020
   End
End
Attribute VB_Name = "frmIPOLotsInspection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iColNo As Integer

Private Sub cmbProcessNo_Change()
    Call cmbProcessNo_Click
End Sub

Private Sub cmbProcessNo_Click()
    
    If cmbProcessNo.MatchFound = False Then Exit Sub
    On Error GoTo err
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM ProcessInspectionParameters WHERE ProcessID=" & cmbProcessNo.List(cmbProcessNo.ListIndex, 1) & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EntryID & "'", !parameterName & "")
            ITM.ListSubItems.add
            ITM.ListSubItems.add
            ITM.ListSubItems.add
            ITM.ListSubItems.add , , cmbProcessNo.List(cmbProcessNo.ListIndex, 2)
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "OK"
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT * FROM ProcessInspection WHERE IPO_RefID=" & Val(txtLotBarcode) & " AND ProcessID=" & cmbProcessNo.List(cmbProcessNo.ListIndex, 1), con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            LV.Tag = Val(!EntryID & "")
            txtDisposation = !Disposation & ""
            txtRemarks = !Comments & ""
            DTInspection = !DT
            '------------------------------------
        Else
            LV.Tag = "0"
            txtDisposation = ""
            txtRemarks = ""
            DTInspection = Date
        End If
        .Close
        
        If Val(LV.Tag) > 0 Then
            .Open "SELECT * FROM ProcessInspectionDetail WHERE PI_RefID=" & Val(LV.Tag), con, adOpenForwardOnly, adLockReadOnly
            Do Until .EOF
                Set ITM = LV.ListItems(!PIP_RefID & "'")
                ITM.SubItems(1) = !AQL & ""
                ITM.SubItems(2) = !SampleSize & ""
                ITM.SubItems(3) = !RejectOn & ""
                ITM.SubItems(4) = !TestSpecificationNo & ""
                ITM.SubItems(5) = !ActualRejection & ""
                ITM.SubItems(6) = !Status & ""
                .MoveNEXT
            Loop
            .Close
        End If
    End With
    Set rs = Nothing
    If Val(LV.Tag) > 0 Then
        cmdPrint.Visible = True
    Else
        cmdPrint.Visible = False
    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    If Val(LV.Tag) <= 0 Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\QAIR.rpt")
    Dim f As New frmPrevRpt
    Me.Hide
    f.ShowReport "{ProcessInspection.EntryID}=" & Val(LV.Tag), rpt
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdReceive_Click()

    On Error GoTo err
    
    Dim i As Integer
    
    Dim myDT As Date
    myDT = Date
    Call StartTrans(con)
     
    Dim lEntryID As Long
    
    If Val(LV.Tag) > 0 Then
        con.Execute "DELETE FROM ProcessInspection WHERE EntryID=" & Val(LV.Tag)
    End If
    
    con.Execute "INSERT INTO ProcessInspection(IPO_RefID,ProcessID,Disposation,Comments,DT,UserName,MachineName) " & _
     "VALUES(" & txtLotBarcode.Text & "," & cmbProcessNo.List(cmbProcessNo.ListIndex, 1) & ",'" & txtDisposation & _
     "','" & txtRemarks & "','" & DTInspection & "','" & CurrentUserName & "','" & strComputerName & "')"
      
    lEntryID = GetSingleLongValue("MAX(EntryID)", "ProcessInspection", " WHERE MachineName='" & strComputerName & "'")
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            con.Execute "INSERT INTO ProcessInspectionDetail(PI_RefID,PIP_RefID,AQL,SampleSize,RejectOn," & _
             "TestSpecificationNo,ActualRejection,Status) VALUES(" & lEntryID & "," & Val(.key) & ",'" & _
             .SubItems(1) & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & "','" & _
             .SubItems(5) & "','" & .SubItems(6) & "')"
        End With
    Next
    con.CommitTrans
            
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub Form_Load()
    DTInspection = Date
End Sub

Private Sub LV_DblClick()
    If LV.ListItems.count = 0 Then Exit Sub
    If iColNo = 1 Then Exit Sub
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
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
        If x > Val(LV.ColumnHeaders(i).Left) And x < Val(LV.ColumnHeaders(i).Left) + Val(LV.ColumnHeaders(i).Width) Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Sub txtLotBarcode_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Dim rs As New ADODB.Recordset
        Dim ITM As ListItem
        Dim strStatus As String, iStatus As Integer
        Dim i As Integer
        Dim lForeColor As Long
        With rs
             
            .Open "SELECT * FROM VInternalProductionOrderDetail WHERE EntryID=" & Val(txtLotBarcode), con, adOpenForwardOnly, adLockReadOnly
            If .EOF Then
                MsgBox "Invalid Barcode.", vbInformation
                Exit Sub
            Else
                txtItemCode = !CompItemCode & ""
                txtItemCode.Tag = !CustomerPO & ""
                txtItemName = !ItemName & ""
                txtLotQty = Val(!LotQty & "")
                txtLotBarcode.Tag = Val(!EntryID & "")
                txtOrderNo = !CustomerPO & ""
                .Close
                                 
            End If
            .Open "SELECT ItemPic FROM Items WHERE ItemID='" & txtItemCode & "'", con, adOpenForwardOnly, adLockReadOnly
            If Not IsNull(![ItemPic]) Then
                Set Image1.DataSource = rs
                Image1.DataField = "ItemPic"
            End If
            .Close
        End With
    End If
    Set rs = Nothing
    
    Call AddToCombo(cmbProcessNo, "'{' + Code + '} ' + Description", "VItemProcesses", " WHERE IsExist='" & txtItemCode & "' AND InspectionProcess=1", , "ProcessID,Code")
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        LV.SelectedItem.SubItems(iColNo - 1) = txtEdit
        If LV.SelectedItem.Index < LV.ListItems.count Then
            LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
            LV_DblClick
            Exit Sub
        ElseIf iColNo < LV.ColumnHeaders.count Then
            iColNo = iColNo + 1
            LV.ListItems(1).Selected = True
            Call LV_DblClick
            Exit Sub
        End If
        
        txtEdit.Visible = False
        'Call txtLotBarcode_KeyPress(13)
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub
