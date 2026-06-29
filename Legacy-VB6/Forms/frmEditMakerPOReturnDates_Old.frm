VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmEditMakerPOReturnDates_Old 
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
      Format          =   152895491
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
      NumItems        =   4
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
         SubItemIndex    =   3
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
Attribute VB_Name = "frmEditMakerPOReturnDates_Old"
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
        .Open "SELECT VendIssdDetail_ReturnDTs.* FROM VendIssdDetail_ReturnDTs INNER JOIN VendIssdDetail ON VendIssdDetail.EntryID=VendIssdDetail_ReturnDTs.VIS_RefID WHERE VendIssdDetail.RefID=" & lEntryID & " ORDER BY VendIssdDetail_ReturnDTs.EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EntryID & "'", LV.ListItems.count + 1)
            ITM.ListSubItems.add(, , Format(!ReturnDT, "dd-MMM-yyyy")).Tag = Format(!ReturnDT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , Val(!IssQty & "")
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
            con.Execute "UPDATE VendIssdDetail_ReturnDTs SET ReturnDT='" & .SubItems(1) & "',Remarks='" & .SubItems(3) & "' WHERE EntryID=" & Val(.key)
            
            'If Return Date Changed...
            If Format(.SubItems(1), "dd-MMM-yyyy") <> .ListSubItems(1).Tag Then
                con.Execute "INSERT INTO VendIssdDetail_ReturnDTs_Revisions(VIS_RD_RefID,ReturnDT,UserName,MachineName,Remarks) " & _
                 "VALUES(" & Val(.key) & ",'" & .ListSubItems(1).Tag & "','" & CurrentUserName & "','" & strComputerName & "','" & .ListSubItems(3).Tag & "')"
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
    
    If iColNo = 4 Then
        With txtEditRemarks
            .Move LV.Left + LV.ColumnHeaders(4).Left, LV.Top + LV.SelectedItem.Top
            .Text = LV.SelectedItem.SubItems(3)
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
        LV.SelectedItem.SubItems(3) = txtEditRemarks.Text
        txtEditRemarks.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEditRemarks.Visible = False
    End If
    
End Sub

Private Sub txtEditRemarks_LostFocus()
    txtEditRemarks.Visible = False
End Sub
