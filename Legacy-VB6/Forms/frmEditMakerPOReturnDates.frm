VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmEditMakerPOReturnDates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Return Dates"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7005
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEditRemarks 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4215
      TabIndex        =   6
      Top             =   255
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   450
      Left            =   3675
      TabIndex        =   5
      Top             =   4590
      Width           =   1410
   End
   Begin MSComCtl2.DTPicker DTReturn 
      Height          =   315
      Left            =   2745
      TabIndex        =   4
      Top             =   2175
      Visible         =   0   'False
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   152961027
      CurrentDate     =   40637
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   450
      Left            =   5175
      TabIndex        =   3
      Top             =   4590
      Width           =   1410
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3840
      Left            =   375
      TabIndex        =   2
      Top             =   705
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   6773
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Rate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Remarks"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.TextBox txtMasterPONo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   375
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   330
      Width           =   2970
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Master PO No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   375
      TabIndex        =   0
      Top             =   90
      Width           =   1365
   End
End
Attribute VB_Name = "frmEditMakerPOReturnDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bReturnValue As Boolean
Dim lEntryID As Long, iColNo As Long

Public Function ShowMe(p_lEntryID As Long) As Boolean

    lEntryID = p_lEntryID
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT VendIssdDetail_ReturnDTs.*,VendIssdDetail.Rate FROM VendIssdDetail_ReturnDTs INNER JOIN VendIssdDetail ON VendIssdDetail.EntryID=VendIssdDetail_ReturnDTs.VIS_RefID WHERE VendIssdDetail.RefID=" & lEntryID & " ORDER BY VendIssdDetail_ReturnDTs.EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EntryID & "'", LV.ListItems.count + 1)
            ITM.Tag = Val(!VIS_RefID & "")
            ITM.ListSubItems.add(, , Format(!ReturnDT, "dd-MMM-yyyy")).Tag = Format(!ReturnDT, "dd-MMM-yyyy")
            ITM.ListSubItems.add(, , Val(!IssQty & "")).Tag = Val(!IssQty & "")
            ITM.ListSubItems.add(, , Val(!Rate & "")).Tag = Val(!Rate & "")
            ITM.ListSubItems.add(, , !Remarks & "").Tag = !Remarks & ""
            .MoveNEXT
        Loop
        
        .Close
    End With
    txtMasterPONo = GetSingleStringValue("MasterPONo", "VendIssued", " WHERE EntryID=" & lEntryID)
    
    Me.Show 1
    
    ShowMe = bReturnValue
    
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()

    Dim i As Integer
    Call StartTrans(con)
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            con.Execute "UPDATE VendIssdDetail_ReturnDTs SET ReturnDT='" & .SubItems(1) & "',IssQty=" & Val(.SubItems(2)) & ",Remarks='" & .SubItems(4) & "' WHERE EntryID=" & Val(.key)
            con.Execute "UPDATE VendIssdDetail SET Rate=" & Val(.SubItems(3)) & ",IssQty=" & Val(.SubItems(2)) & " WHERE EntryID=" & Val(.Tag)
            'If Return Date,Rate or Qty Changed...
            If (Format(.SubItems(1), "dd-MMM-yyyy") <> .ListSubItems(1).Tag) Or (Val(.SubItems(2)) <> Val(.ListSubItems(2).Tag)) Or (Val(.SubItems(3)) <> Val(.ListSubItems(3).Tag)) Then
                con.Execute "INSERT INTO VendIssdDetail_ReturnDTs_Revisions(VIS_RD_RefID,ReturnDT,UserName,MachineName,Remarks,Qty_Old,Qty_New,Rate_Old,Rate_New) " & _
                 "VALUES(" & Val(.key) & ",'" & .ListSubItems(1).Tag & "','" & CurrentUserName & "','" & strComputerName & "','" & .ListSubItems(4).Tag & _
                 "'," & Val(.ListSubItems(2).Tag) & "," & Val(.SubItems(2)) & "," & Val(.ListSubItems(3).Tag) & "," & Val(.SubItems(3)) & ")"
            End If
            
        End With
    Next
    con.CommitTrans
    bReturnValue = True
    
    Unload Me
    
End Sub

Private Sub DTReturn_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        LV.SelectedItem.SubItems(1) = Format(DTReturn.Value, "dd-MMM-yyyy")
        DTReturn.Visible = False
    ElseIf KeyCode = 27 Then
        DTReturn.Visible = False
    End If
    
End Sub

Private Sub DTReturn_LostFocus()
    DTReturn.Visible = False
End Sub

Private Sub Form_Load()
    bReturnValue = False
    DTReturn.Value = Date
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If iColNo = 5 Or iColNo = 4 Or iColNo = 3 Then  'Remarks,Rate,Qty
        With txtEditRemarks
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
            If iColNo = 5 Then
                .Text = LV.SelectedItem.SubItems(4)
            Else
                .Text = Val(LV.SelectedItem.SubItems(iColNo - 1))
            End If
            .Visible = True
            .SetFocus
        End With
    Else
        With DTReturn
            .Move LV.Left + LV.ColumnHeaders(2).Left, LV.Top + LV.SelectedItem.Top
            .Value = LV.SelectedItem.SubItems(1)
            .Visible = True
            .SetFocus
        End With
    End If
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        With LV.ColumnHeaders(i)
            If x > Val(.Left) And x < (Val(.Left) + Val(.Width)) Then
                iColNo = i
                Exit For
            End If
        End With
    Next
    
End Sub

Private Sub txtEditRemarks_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If iColNo = 5 Then
            LV.SelectedItem.SubItems(4) = txtEditRemarks.Text
        Else
            LV.SelectedItem.SubItems(iColNo - 1) = Val(txtEditRemarks.Text)
        End If
        txtEditRemarks.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEditRemarks.Visible = False
    End If
    
End Sub

Private Sub txtEditRemarks_LostFocus()
    txtEditRemarks.Visible = False
End Sub
