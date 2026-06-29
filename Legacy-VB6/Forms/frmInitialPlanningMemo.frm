VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmInitialPlanningMemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Initial Planning Memo"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11670
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Detail :"
      Height          =   6750
      Left            =   105
      TabIndex        =   9
      Top             =   1290
      Width           =   11445
      Begin ComboList.Usercontrol1 cmbEmp 
         Height          =   285
         Left            =   165
         TabIndex        =   24
         Top             =   4425
         Width           =   3435
         _ExtentX        =   6059
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
      Begin VB.CommandButton cmdAdd 
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
         Height          =   315
         Left            =   3640
         TabIndex        =   25
         Top             =   4395
         Width           =   525
      End
      Begin VB.TextBox txtGeneralRemarks 
         Height          =   1635
         Left            =   4215
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   4425
         Width           =   6975
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3300
         TabIndex        =   17
         Top             =   1005
         Visible         =   0   'False
         Width           =   2000
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   390
         Left            =   165
         TabIndex        =   15
         Top             =   435
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   688
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
         Format          =   118554627
         CurrentDate     =   40894
      End
      Begin MSComctlLib.ListView LV 
         Height          =   3345
         Left            =   165
         TabIndex        =   14
         Top             =   840
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   5900
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Stage"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Plan Qty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Remarks"
            Object.Width           =   11404
         EndProperty
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8220
         TabIndex        =   11
         Top             =   6135
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   9735
         TabIndex        =   10
         Top             =   6135
         Width           =   1455
      End
      Begin MSComctlLib.ListView LVUsers 
         Height          =   1320
         Left            =   165
         TabIndex        =   23
         Top             =   4755
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   2328
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User"
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Users"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   225
         TabIndex        =   22
         Top             =   4200
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "General Remarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   4215
         TabIndex        =   20
         Top             =   4200
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date of Plan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   16
         Top             =   225
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order Details :"
      Height          =   1095
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   11445
      Begin VB.TextBox txtDeliveryDT 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   9465
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtOrderNo 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   7755
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   6870
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   870
      End
      Begin VB.TextBox txtCustomer 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   5490
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1365
      End
      Begin VB.TextBox txtItemName 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtItemCode 
         Alignment       =   2  'Center
         BackColor       =   &H00F1FEFE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Delivery Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   9510
         TabIndex        =   19
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Order No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   7785
         TabIndex        =   13
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   6885
         TabIndex        =   7
         Top             =   270
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   5760
         TabIndex        =   5
         Top             =   255
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1380
         TabIndex        =   3
         Top             =   255
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Item Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   1
         Top             =   255
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmInitialPlanningMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lEntryID As Long, lFOI_ID As Long, iColNo As Integer
Dim bCancelled As Boolean

Public Function ViewInitialPlanningMemo(p_lFOI_ID As Long, lRunningStatus As Long, lPOStatus As Long, lSemiStatus As Long, lFinishStatus As Long, lForgingStatus As Long) As Boolean
    
    lFOI_ID = p_lFOI_ID
    
    Dim ITM As ListItem
    
    LV.ListItems.Clear
    
    Set ITM = LV.ListItems.add(, , "Running")
    ITM.ListSubItems.add(, , lRunningStatus).Tag = "RunningStatus"
    ITM.ListSubItems.add(, , "").Tag = "RunningPlanned"
    ITM.ListSubItems.add(, , "").Tag = "RunningRemarks"
    
    Set ITM = LV.ListItems.add(, , "PO")
    ITM.ListSubItems.add(, , lPOStatus).Tag = "POStatus"
    ITM.ListSubItems.add(, , "").Tag = "POPlanned"
    ITM.ListSubItems.add(, , "").Tag = "PORemarks"
    
    Set ITM = LV.ListItems.add(, , "Semi")
    ITM.ListSubItems.add(, , lSemiStatus).Tag = "SemiStatus"
    ITM.ListSubItems.add(, , "").Tag = "SemiPlanned"
    ITM.ListSubItems.add(, , "").Tag = "SemiRemarks"
    
    Set ITM = LV.ListItems.add(, , "Finish")
    ITM.ListSubItems.add(, , lFinishStatus).Tag = "FinishStatus"
    ITM.ListSubItems.add(, , "").Tag = "FinishPlanned"
    ITM.ListSubItems.add(, , "").Tag = "FinishRemarks"
    
    Set ITM = LV.ListItems.add(, , "Forging")
    ITM.ListSubItems.add(, , lForgingStatus).Tag = "ForgingStatus"
    ITM.ListSubItems.add(, , "").Tag = "ForgingPlanned"
    ITM.ListSubItems.add(, , "").Tag = "ForgingRemarks"
    
    Call ViewDetails(lFOI_ID)
    
    Me.Show 1
    ViewInitialPlanningMemo = Not bCancelled
    
End Function

Private Sub ViewDetails(p_lEntryID As Long)
    
    Dim ITM As ListItem
    Dim rs As New ADODB.Recordset, i As Integer
    With rs
        .Open "SELECT * FROM VFOrderItems WHERE ID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtItemCode = !ItemCode & ""
            txtItemName = !ItemName & ""
            txtQty = Val(!Qty & "")
            txtCustomer = !CustCode & ""
            txtOrderNo = !OrderNo & ""
            txtDeliveryDT = Format(!DeliveryDT, "dd-MMM-yyyy")
        End If
        .Close
        
        Dim j As Integer
        Dim lEntryID As Long
        .Open "SELECT * FROM FOrderItems_InitialPlanningMemo WHERE FOI_RefID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            cmdSave.Enabled = False
            lEntryID = Val(!EntryID & "")
            DT.Value = !DT
            txtGeneralRemarks = !GeneralRemarks & ""
            For i = 1 To LV.ListItems.count
                With LV.ListItems(i)
                    For j = 1 To .ListSubItems.count
                        .SubItems(j) = rs.Fields(.ListSubItems(j).Tag) & ""
                    Next
                End With
            Next
        Else
            cmdSave.Enabled = True
            lEntryID = 0
        End If
        .Close
        
        .Open "SELECT * FROM FOrderItems_InitialPlanningMemo_Users WHERE FOI_IPM_RefID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        LVUsers.ListItems.Clear
        Do Until .EOF
            LVUsers.ListItems.add , !UserName & "", !UserName & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
         
End Sub


Private Sub cmdAdd_Click()
    On Error GoTo err
    LVUsers.ListItems.add , cmbEmp.ID, cmbEmp.Text
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()

    bCancelled = True
    Unload Me
     
End Sub

Private Sub cmdSave_Click()
    
    On Error GoTo err
    
    If MsgBox("Are you sure to Save?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Call StartTrans(con)
    'Save...
    
    con.Execute "INSERT INTO FOrderItems_InitialPlanningMemo(FOI_RefID,DT,UserName,MachineName,GeneralRemarks) VALUES(" & _
     lFOI_ID & ",'" & DT.Value & "','" & CurrentUserName & "','" & strComputerName & "','" & txtGeneralRemarks & "')"
       
    Dim lLatestEntryID As Long
    lLatestEntryID = GetSingleLongValue("MAX(EntryID)", "FOrderItems_InitialPlanningMemo", " WHERE MachineName='" & strComputerName & "'")
    
    Dim strUpdateSQL As String
    strUpdateSQL = "UPDATE FOrderItems_InitialPlanningMemo SET "
    Dim i As Integer, j As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            For j = 1 To .ListSubItems.count
                strUpdateSQL = strUpdateSQL & .ListSubItems(j).Tag & "='" & .SubItems(j) & "',"
            Next
        End With
    Next
    
    strUpdateSQL = Left(strUpdateSQL, Len(strUpdateSQL) - 1)
    strUpdateSQL = strUpdateSQL & " WHERE EntryID=" & lLatestEntryID
    
    con.Execute strUpdateSQL
    
    
    
    con.Execute "DELETE FROM FOrderItems_InitialPlanningMemo_Users WHERE FOI_IPM_RefID=" & lLatestEntryID
    For i = 1 To LVUsers.ListItems.count
        con.Execute "INSERT INTO FOrderItems_InitialPlanningMemo_Users(FOI_IPM_RefID,UserName) VALUES(" & lLatestEntryID & ",'" & LVUsers.ListItems(i).Text & "')"
    Next
    
    con.CommitTrans
    bCancelled = False
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()

    bCancelled = True
    DT.Value = Date
    
    
    cmbEmp.ClearVals
    cmbEmp.ListHeight = 2300
    cmbEmp.AddVals con, "UserName", "Users", "UserName"
    
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    If cmdSave.Enabled = False Then Exit Sub
    
    If iColNo <> 4 Then iColNo = 3
    
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
        With LV.ColumnHeaders(i)
            If x > Val(.Left) And x < (Val(.Left) + Val(.Width)) Then
                iColNo = i
                Exit Sub
            End If
        End With
    Next
End Sub

Private Sub LVUsers_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 46 Then
'        con.Execute "DELETE"
'    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(iColNo - 1) = txtEdit
        If iColNo = 3 Then
            iColNo = 4
            Call LV_DblClick
        ElseIf LV.SelectedItem.Index < LV.ListItems.count Then
        
            LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
            iColNo = 3
            Call LV_DblClick
            
        Else
            txtEdit.Visible = False
        End If
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub


