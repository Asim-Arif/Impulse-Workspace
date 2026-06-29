VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSupplierReEvaluationAE 
   Caption         =   "Supplier Re-Evaluation"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10020
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
   ScaleHeight     =   8370
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEdit 
      Height          =   360
      Left            =   1650
      TabIndex        =   11
      Top             =   2130
      Visible         =   0   'False
      Width           =   1500
   End
   Begin MSComctlLib.ListView LV 
      Height          =   6435
      Left            =   60
      TabIndex        =   10
      Top             =   690
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   11351
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Requirements"
         Object.Width           =   9894
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Rating"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Comments"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6645
      TabIndex        =   9
      Top             =   7815
      Width           =   1605
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8325
      TabIndex        =   8
      Top             =   7815
      Width           =   1605
   End
   Begin VB.ComboBox cmbEvaluatedBy 
      Height          =   315
      Left            =   3630
      TabIndex        =   5
      Tag             =   "EvaluatedBy"
      Top             =   7380
      Width           =   2820
   End
   Begin MSComCtl2.DTPicker DTReview 
      Height          =   315
      Left            =   4950
      TabIndex        =   13
      Top             =   330
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   556
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
      Format          =   101908483
      CurrentDate     =   37384
   End
   Begin MSForms.Label Label1 
      Height          =   195
      Left            =   4935
      TabIndex        =   12
      Top             =   105
      Width           =   1080
      VariousPropertyBits=   276824091
      Caption         =   "Review Date"
      Size            =   "1905;344"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtApprovedBy 
      Height          =   315
      Left            =   7185
      TabIndex        =   7
      Tag             =   "ApprovedBy"
      Top             =   7365
      Width           =   2760
      VariousPropertyBits=   746604571
      Size            =   "4868;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtPreparedBy 
      Height          =   315
      Left            =   75
      TabIndex        =   6
      Tag             =   "PreparedBy"
      Top             =   7365
      Width           =   2820
      VariousPropertyBits=   746604571
      Size            =   "4974;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label lblAproved 
      Height          =   195
      Left            =   7170
      TabIndex        =   4
      Top             =   7140
      Width           =   1185
      VariousPropertyBits=   276824091
      Caption         =   " Approved By:"
      Size            =   "2090;344"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblEvalute 
      Height          =   195
      Left            =   3615
      TabIndex        =   3
      Top             =   7140
      Width           =   975
      VariousPropertyBits=   276824091
      Caption         =   "Review By :"
      Size            =   "1720;344"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblPrepaid 
      Height          =   270
      Left            =   90
      TabIndex        =   2
      Top             =   7140
      Width           =   1230
      VariousPropertyBits=   276824091
      Caption         =   " Prepared By:"
      Size            =   "2170;476"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblSelection 
      Height          =   195
      Left            =   45
      TabIndex        =   1
      Top             =   105
      Width           =   1995
      VariousPropertyBits=   276824091
      Caption         =   " Select Maker / Vendor:"
      Size            =   "3519;344"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbSupplier 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Tag             =   "AccNo"
      Top             =   330
      Width           =   4890
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "8625;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmSupplierReEvaluationAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iColNo As Integer, lEntryID As Long

Private Sub cmbSupplier_Change()
    Call cmbSupplier_Click
End Sub

Private Sub cmbSupplier_Click()
    If cmbSupplier.MatchFound Then
        'Call AddEvaluation(cmbSupplier.List(cmbSupplier.ListIndex, 1))
    Else
        Call ClearAll
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    
    On Error GoTo err
    
    If cmbSupplier.MatchFound = -False Then
        MsgBox "Please Select Supplier.", vbInformation
        Exit Sub
    End If
      
    Dim strSql As String, lmyEntryID As Long
    
    Call StartTrans(con)
    If lEntryID = 0 Then
        strSql = "INSERT INTO SupplierReEvaluation(AccNo,ReviewDT,ReviewBy,PreparedBy,ApprovedBy,UserName,MachineName) VALUES('" & _
         cmbSupplier.List(cmbSupplier.ListIndex, 1) & "','" & DTReview & "','" & cmbEvaluatedBy.Text & _
         "','" & txtPreparedBy.Text & "','" & txtApprovedBy & "','" & CurrentUserName & "','" & strComputerName & "')"
    Else
        strSql = "UPDATE SupplierReEvaluation SET AccNo='" & cmbSupplier.List(cmbSupplier.ListIndex, 1) & "',ReviewDT='" & _
         DTReview & "',ReviewBy='" & cmbEvaluatedBy.Text & "',PreparedBy='" & txtPreparedBy & "',ApprovedBy='" & _
         txtApprovedBy & "' WHERE EntryID=" & lEntryID
    End If
    
    con.Execute strSql
     
    If lEntryID = 0 Then
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "SupplierReEvaluation")
    Else
        lmyEntryID = lEntryID
    End If
    
    strSql = "UPDATE SupplierReEvaluation SET "
    
    Dim i As Integer
    For i = 2 To LV.ListItems.count
        With LV.ListItems(i)
            If .Bold = False Then
                strSql = strSql & .Tag & "=" & Val(.SubItems(1)) & "," & .Tag & "Comments='" & .SubItems(2) & "',"
            End If
        End With
    Next
    
    strSql = Left(strSql, Len(strSql) - 1)
    strSql = strSql & " WHERE EntryID=" & lmyEntryID
     
    con.Execute strSql
    con.CommitTrans
    
    MsgBox "Saved Successfully.", vbInformation
     
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Load()
 
    Call AddToCombo(cmbSupplier, "AccTitle", "VVendersAndMakers", , , "AccNo")
    Dim ITM As ListItem
    
    Set ITM = LV.ListItems.add(, , "Quality")
    ITM.Bold = True
    
    Set ITM = LV.ListItems.add(, , "A. Quality of Product (According to Specification)")
    ITM.Tag = "QualityOfProduction"
    Set ITM = LV.ListItems.add(, , "B. Price")
    ITM.Tag = "Price"
    
    Set ITM = LV.ListItems.add(, , "C. Packing")
    ITM.Tag = "Packing"
    
    Set ITM = LV.ListItems.add(, , "D. Rejection Rate < 10%")
    ITM.Tag = "RejectionRate"
    Set ITM = LV.ListItems.add(, , "E. Delivery in time (Must be > 50%)")
    ITM.Tag = "DeliveryInTime"
    Set ITM = LV.ListItems.add(, , "F. Re-Work Rate < 15")
    ITM.Tag = "ReWorkRate"
    Set ITM = LV.ListItems.add(, , "G. No. of Orders Completed > 80%")
    ITM.Tag = "NoOfOrdersCompleted"
    Set ITM = LV.ListItems.add(, , "H. Cancelled Orders < 20%")
    ITM.Tag = "CancelledOrders"
    Set ITM = LV.ListItems.add(, , "I. Response time")
    ITM.Tag = "ResponseTime"
    
    Set ITM = LV.ListItems.add(, , "Performance")
    ITM.Bold = True
    
    Set ITM = LV.ListItems.add(, , "A. Technical Ability")
    ITM.Tag = "TechnicalAbility"
    Set ITM = LV.ListItems.add(, , "B. Production Capacity")
    ITM.Tag = "ProductionCapacity"
    
    Set ITM = LV.ListItems.add(, , "Environment System")
    ITM.Bold = True
    
    Set ITM = LV.ListItems.add(, , "A. Environment Requirements")
    ITM.Tag = "EnvironmentRequirements"
    Set ITM = LV.ListItems.add(, , "B. Lighting, Ventilation and vehicles smoke free")
    ITM.Tag = "LightingEtc"
    
    iColNo = 2
    DTReview.Value = Date
End Sub


Private Sub ClearAll()

    
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.Bold Then Exit Sub
    
    Call SetEntry(2)
    
End Sub

Private Sub SetEntry(p_iColNo As Integer)

    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(p_iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(p_iColNo).Width
        .Text = LV.SelectedItem.SubItems(p_iColNo - 1)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(iColNo - 1) = txtEdit
        If iColNo = 2 Then
            iColNo = 3
        ElseIf iColNo = 3 Then
            iColNo = 2
        End If
        If iColNo = 2 Then
            If LV.SelectedItem.Index = LV.ListItems.count Then
                txtEdit.Visible = False
                Exit Sub
            Else
                Dim iNextIndex As Integer
                iNextIndex = LV.SelectedItem.Index + 1
                If LV.ListItems(iNextIndex).Bold Then
                    iNextIndex = iNextIndex + 1
                End If
                LV.ListItems(iNextIndex).Selected = True
                Call SetEntry(iColNo)
            End If
        Else
            iColNo = 3
            Call SetEntry(iColNo)
        End If
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Public Sub EditEntry(p_lEntryID As Long)

    lEntryID = p_lEntryID
         
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    With rs
        .Open "SELECT * FROM SupplierReEvaluation WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        For i = 0 To cmbSupplier.ListCount - 1
            If cmbSupplier.List(i, 1) = !AccNo & "" Then
                cmbSupplier.ListIndex = i
                Exit For
            End If
        Next
        
        txtPreparedBy = !PreparedBy & ""
        cmbEvaluatedBy = !ReviewBy & ""
        txtApprovedBy = !ApprovedBy & ""
         
        For i = 2 To LV.ListItems.count
            With LV.ListItems(i)
                If .Bold = False Then
                    .SubItems(1) = Val(rs.Fields(.Tag) & "")
                    .SubItems(2) = rs.Fields(.Tag & "Comments") & ""
                End If
            End With
        Next
        .Close
    End With
    Set rs = Nothing
    
    Me.Show 1
    
End Sub
