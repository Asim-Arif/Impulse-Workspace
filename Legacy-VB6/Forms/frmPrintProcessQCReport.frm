VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmPrintProcessQCReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QC"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11715
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassQty 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   450
      Left            =   2475
      TabIndex        =   17
      Top             =   1410
      Width           =   2190
   End
   Begin VB.PictureBox PicEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1725
      ScaleHeight     =   345
      ScaleWidth      =   7740
      TabIndex        =   11
      Top             =   2325
      Visible         =   0   'False
      Width           =   7770
      Begin VB.TextBox txtRemarks 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   4785
         TabIndex        =   15
         Top             =   -30
         Width           =   2000
      End
      Begin VB.TextBox txtPieces 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   3795
         TabIndex        =   14
         Top             =   -30
         Width           =   1000
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   390
         Left            =   -30
         TabIndex        =   12
         Top             =   -30
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   688
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   157810691
         CurrentDate     =   44205
      End
      Begin MSForms.ComboBox cmbAction 
         Height          =   390
         Left            =   6780
         TabIndex        =   16
         Top             =   -30
         Width           =   1005
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "1773;688"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbEmp 
         Height          =   390
         Left            =   1425
         TabIndex        =   13
         Top             =   -30
         Width           =   2400
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4233;688"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4680
      Left            =   105
      TabIndex        =   10
      Top             =   1905
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   8255
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Activities"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2593
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Responsible Person"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Piecies"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Remarks"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Action"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.TextBox txtLotQty 
      Alignment       =   2  'Center
      BackColor       =   &H00D7FFFF&
      ForeColor       =   &H00800000&
      Height          =   450
      Left            =   10380
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   705
      Width           =   1215
   End
   Begin VB.TextBox txtProcess 
      BackColor       =   &H00D7FFFF&
      ForeColor       =   &H00800000&
      Height          =   450
      Left            =   2475
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   705
      Width           =   7890
   End
   Begin VB.TextBox txtTrayNo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   450
      Left            =   105
      TabIndex        =   0
      Top             =   1410
      Width           =   2355
   End
   Begin VB.TextBox txtLotNo 
      Alignment       =   2  'Center
      BackColor       =   &H00D7FFFF&
      ForeColor       =   &H00800000&
      Height          =   450
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   705
      Width           =   2355
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9075
      TabIndex        =   2
      Top             =   6645
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10380
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6645
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pass Qty"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   2475
      TabIndex        =   18
      Top             =   1185
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lot Qty"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   10380
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   2475
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tray No."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Top             =   1185
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lot No."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   3
      Top             =   480
      Width           =   600
   End
End
Attribute VB_Name = "frmPrintProcessQCReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lVRD_EntryID As Long

Private Sub cmbAction_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode.Value = 13 Then
    
        If cmbEmp.MatchFound = False Then
            MsgBox "Please Select Employee.", vbInformation
            Exit Sub
        End If
        If cmbAction.MatchFound = False Then
            MsgBox "Please Select Action.", vbInformation
            Exit Sub
        End If
        
        LV.SelectedItem.SubItems(1) = Format(DT, "dd-MMM-yyyy")
        LV.SelectedItem.SubItems(2) = cmbEmp.Text
        LV.SelectedItem.ListSubItems(2).Tag = cmbEmp.List(cmbEmp.ListIndex, 1)
             
        LV.SelectedItem.SubItems(3) = Val(txtPieces)
        LV.SelectedItem.SubItems(4) = txtRemarks
        LV.SelectedItem.SubItems(5) = cmbAction.Text
        LV.SelectedItem.ListSubItems(5).Tag = cmbAction.ListIndex
        
        PicEdit.Visible = False
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub ShowMe(p_lVRD_EntryID As Long)

    lVRD_EntryID = p_lVRD_EntryID
    Dim lProcessID As Long
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT Description,LotNo,RcvdQty,ProcessID FROM VVendRcvdDetail WHERE EntryID=" & p_lVRD_EntryID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtLotNo = !LotNo & ""
            txtProcess = !Description & ""
            txtLotQty = Val(!RcvdQty & "")
            lProcessID = Val(!ProcessID & "")
        End If
        .Close
        
        .Open "SELECT TrayNo,PassQty FROM VendRcvdDetail_QC_Report WHERE VRD_RefID=" & lVRD_EntryID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtTrayNo = (!TrayNo & "")
            txtPassQty = Val(!PassQty & "")
        End If
        .Close
        
        .Open "SELECT * FROM VVendRcvdDetail_QC_Report_PIP WHERE VRD_RefID=" & lVRD_EntryID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then    'Not Feeded...
            .Close
            .Open "SELECT * FROM Process_Inspection_Points WHERE ProcessID=" & lProcessID & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
            Do Until .EOF
                Set ITM = LV.ListItems.add(, , !Point_Description & "")
                ITM.Tag = Val(!EntryID & "")
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
                ITM.ListSubItems.add , , ""
                .MoveNEXT
            Loop
            .Close
        Else
        
            Do Until .EOF
            
                Set ITM = LV.ListItems.add(, , !Point_Description & "")
                ITM.Tag = Val(!PIP_RefID & "")
                ITM.ListSubItems.add , , IIf(IsNull(!DT), "", Format(!DT, "dd-MMM-yyyy"))
                ITM.ListSubItems.add(, , !Name & "").Tag = !EmpId & ""
                ITM.ListSubItems.add , , Val(!Pieces & "")
                ITM.ListSubItems.add , , !Remarks & ""
                ITM.ListSubItems.add(, , IIf(Val(!Action & "") = 1, "Yes", "No")).Tag = Val(!Action & "")
                
                .MoveNEXT
            Loop
            .Close
        End If
    End With
    Me.Show 1
End Sub

Private Sub cmdPrint_Click()
    
    Dim lCount As Long, i As Integer
    lCount = GetSingleLongValue("COUNT(*)", "VendRcvdDetail_QC_Report", " WHERE VRD_RefID=" & lVRD_EntryID)
    If lCount > 0 Then
        con.Execute "UPDATE VendRcvdDetail_QC_Report SET TrayNo='" & txtTrayNo & "',PassQty=" & Val(txtPassQty) & " WHERE VRD_RefID=" & lVRD_EntryID
    Else
        con.Execute "INSERT INTO VendRcvdDetail_QC_Report(VRD_RefID,TrayNo,PassQty,UserName,MachineName) VALUES(" & _
         lVRD_EntryID & ",'" & txtTrayNo & "'," & Val(txtPassQty) & ",'" & CurrentUserName & "','" & strComputerName & "')"
    End If
    
    Dim strDT As String, iRowAffect As Integer
    
    For i = 1 To LV.ListItems.count
    
        With LV.ListItems(i)
            If .SubItems(1) = "" Then
                strDT = "NULL"
            Else
                strDT = "'" & Format(.SubItems(1), "dd-MMM-yyyy") & "'"
            End If
            con.Execute "UPDATE VendRcvdDetail_QC_Report_PIP SET DT=" & strDT & ",EmpID='" & .ListSubItems(2).Tag & _
             "',Pieces=" & Val(.SubItems(3)) & ",Remarks='" & .SubItems(4) & "',Action=" & Val(.ListSubItems(5).Tag) & " WHERE VRD_RefID=" & _
             lVRD_EntryID & " AND PIP_RefID=" & Val(.Tag), iRowAffect
            
            If iRowAffect = 0 Then
                con.Execute "INSERT INTO VendRcvdDetail_QC_Report_PIP(VRD_RefID,PIP_RefID,DT,EmpID,Pieces,Remarks,Action) " & _
                 "VALUES(" & lVRD_EntryID & "," & Val(.Tag) & "," & strDT & ",'" & .ListSubItems(2).Tag & "'," & Val(.SubItems(3)) & _
                 ",'" & .SubItems(4) & "'," & Val(.ListSubItems(5).Tag) & ")"
                 
            End If
        End With
    Next
    
    If LV.ListItems.count = 0 Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Maker_Receiving_QC_Report.rpt")
    
    Dim f As New frmPrevRpt
    Load f
    Me.Hide
    f.ShowReport "{VVendRcvdDetail.EntryID}=" & lVRD_EntryID, rpt
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    DT.Value = Date
    Call AddToCombo(cmbEmp, "EmpID+' - '+Name", "VEmp", " WHERE Active=1", False, "EmpID")
    cmbAction.AddItem "No"
    cmbAction.AddItem "Yes"
    
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If LV.SelectedItem.SubItems(1) <> "" Then
        DT.Value = LV.SelectedItem.SubItems(1)
    Else
        DT.Value = Date
    End If
    Dim i As Integer
    If LV.SelectedItem.SubItems(2) <> "" Then
        For i = 0 To cmbEmp.ListCount
            If cmbEmp.List(i, 1) = LV.SelectedItem.ListSubItems(2).Tag Then
                cmbEmp.ListIndex = i
                Exit For
            End If
        Next
    End If
    txtPieces = Val(LV.SelectedItem.SubItems(3))
    txtRemarks = LV.SelectedItem.SubItems(4)
    cmbAction.ListIndex = Val(LV.SelectedItem.ListSubItems(5).Tag)
     
    PicEdit.Move LV.Left + LV.ColumnHeaders(2).Left, LV.Top + LV.SelectedItem.Top
    PicEdit.Visible = True
    DT.SetFocus
    
End Sub
