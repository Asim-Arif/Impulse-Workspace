VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendRepairDetail 
   ClientHeight    =   10965
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   10965
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTotalRepairQty 
      Height          =   300
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   10095
      Width           =   1155
   End
   Begin MSComCtl2.DTPicker DTMonth 
      Height          =   300
      Left            =   6675
      TabIndex        =   9
      Top             =   915
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "MMMM, yyyy"
      Format          =   53805059
      CurrentDate     =   38334
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11520
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Repair Detail"
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
         Index           =   1
         Left            =   4440
         TabIndex        =   7
         Top             =   120
         Width           =   3315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Repair Detail"
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
         Index           =   2
         Left            =   4455
         TabIndex        =   8
         Top             =   135
         Width           =   3315
      End
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   1365
      TabIndex        =   2
      Top             =   915
      Width           =   2220
      _ExtentX        =   3916
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
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   3570
      TabIndex        =   3
      Top             =   915
      Width           =   3090
      _ExtentX        =   5450
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
   End
   Begin MSComctlLib.ListView LV 
      Height          =   8685
      Left            =   45
      TabIndex        =   10
      Top             =   1305
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   15319
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
         Text            =   "S. #"
         Object.Width           =   926
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Reciept ID"
         Object.Width           =   2221
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Item Code"
         Object.Width           =   1666
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Item"
         Object.Width           =   4923
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date"
         Object.Width           =   2249
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Qty"
         Object.Width           =   1613
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Rep. Qty"
         Object.Width           =   1575
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Rep. Date"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Process"
         Object.Width           =   2560
      EndProperty
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   8790
      TabIndex        =   5
      Top             =   10560
      Width           =   1425
      ForeColor       =   0
      Caption         =   "Print         "
      PicturePosition =   327683
      Size            =   "2514;635"
      Accelerator     =   80
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   10305
      TabIndex        =   4
      Top             =   10560
      Width           =   1425
      ForeColor       =   0
      Caption         =   " Close      "
      PicturePosition =   327683
      Size            =   "2514;635"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   1380
      TabIndex        =   1
      Top             =   645
      Width           =   8415
      BackColor       =   11517387
      Caption         =   "  Process                                Venders                                                     Month"
      Size            =   "14843;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   8490
      TabIndex        =   0
      Top             =   900
      Width           =   1305
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2302;635"
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmVendRepairDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub cmdPrint_Click()

    On Error GoTo err
    Call StartTrans(con)
    con.Execute "Delete From PrintVendRepairDetail"
    For i = 1 To LV.ListItems.Count
    
        With LV.ListItems(i)
        
            If .SubItems(7) = "N.A" Then
                con.Execute "Insert Into PrintVendRepairDetail(RecieptID,ItemCode,VendID,DT,Qty,RepairRcvd) Values('" & _
                .SubItems(1) & "','" & .SubItems(2) & "'," & cmbEmp.ID & ",'" & .SubItems(4) & "'," & Val(.SubItems(5)) & ",0)"
            Else
                con.Execute "Insert Into PrintVendRepairDetail(RecieptID,ItemCode,VendID,DT,Qty,RepairDT,RepairQty,RepairRcvd) Values('" & _
                .SubItems(1) & "','" & .SubItems(2) & "'," & cmbEmp.ID & ",'" & .SubItems(4) & "'," & Val(.SubItems(5)) & _
                ",'" & .SubItems(7) & "'," & Val(.SubItems(6)) & ",1)"
            End If
            
        End With
    Next
    
    con.CommitTrans
    Dim rpt As craxddrt.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendRepairDetail.rpt")
    rpt.FormulaFields(1).Text = "#" & DTMonth & "#"
    Dim F As New frmPrevRpt
    F.ShowReport " ", rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub cmdRefresh_Click()

    Me.MousePointer = vbHourglass
    Call RefreshLedger
    Me.MousePointer = vbNormal
    'Call Form_Resize
    
End Sub


Private Sub Form_Load()
    
    On Error GoTo err
    
    DTMonth = ServerDate
    
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1600
    
    cmbDept.ClearVals
    'cmbDept.AddItem "<All Processes>", "0"
    cmbDept.AddVals con, "Description", "Processes", "ProcessID", " Where ProcessID=7"
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub cmbdept_matched()
    
    On Error GoTo err
    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Contractors>", "0"
    'cmbEmp.AddItem "<All Contractors>", "0"
    If cmbDept.ID = "0" Then
        'cmbEmp.AddVals con, "'{' +  EmpID + '}  ' + Name", "Contractors", "EmpID"
        Call cmbEmp.AddVals(con, "VenderName", "VVenders", "VendID")
    Else
        Call cmbEmp.AddVals(con, "VenderName", "VVenders", "VendID", " Where PhaseID=" & cmbDept.ID & " OR Left(VendID1,3)='REP'")
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    
End Sub

Private Sub RefreshLedger()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim lRepairQty As Long
    Dim ITM As ListItem
    lRepairQty = 0
    With rs
        .Open "Select Processes.ProcessID,Processes.Description,T1.RecieptID,T1.itemCode,T1.RcvdQty,T2.RcvRepairQty,T1.DT,T2.RepairRcvDT,VItems1.CompleteItemName From (Select RecieptID,ItemCode,DT,RcvdQty From VVendRcvdDetail Where ProcessID=7 AND VendID=" & cmbEmp.ID & ") T1 " & _
                "LEFT OUTER JOIN (Select RecieptID,DT As RepairRcvDT,RcvdQty,RcvRepairQty From VVendRcvdDetail Where ProcessID=12 ) T2 ON " & _
                "T1.RecieptID=T2.RecieptID INNER JOIN (Select Max(ProcessID) As LastProcID,RecieptID From VVendIssdDetail Group By RecieptID) T3 " & _
                "ON T1.RecieptID=T3.RecieptID INNER JOIN VItems1 On T1.ItemCode = VItems1.ItemID INNER JOIN Processes ON T3.LastProcID=Processes.ProcessID " & _
                "Where Month(IsNull(T2.RepairRcvDT, t1.DT))=" & Month(DTMonth) & " And year(IsNull(T2.RepairRcvDT, t1.DT))=" & year(DTMonth) & " AND Left(T1.RecieptID,2)='HS'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Dim lColor As Long
        Do Until .EOF
            If IsNull(!RepairRcvDT) Then
                lColor = vbRed
            Else
                lColor = LV.ForeColor
            End If
            
            Set ITM = LV.ListItems.Add(, , LV.ListItems.Count + 1)
            ITM.ForeColor = lColor
            ITM.ListSubItems.Add , , !RecieptID & ""
            ITM.ListSubItems.Add , , !ItemCode & ""
            ITM.ListSubItems.Add , , !CompleteItemName & ""
            ITM.ListSubItems.Add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.Add , , Val(!RcvdQty & "")
            
            ITM.ListSubItems.Add , , Val(!RcvRepairQty & "")
            lRepairQty = lRepairQty + Val(!RcvRepairQty & "")
            If IsNull(!RepairRcvDT) Then
                ITM.ListSubItems.Add , , "N.A"
            Else
                ITM.ListSubItems.Add , , Format(!RepairRcvDT, "dd-MMM-yyyy")
            End If
            ITM.ListSubItems.Add , , !Description
            For i = 1 To ITM.ListSubItems.Count
                ITM.ListSubItems(i).ForeColor = lColor
            Next
            .MoveNext
        Loop
        .Close
        
    End With
    
    Set rs = Nothing
    txtTotalRepairQty = lRepairQty
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

