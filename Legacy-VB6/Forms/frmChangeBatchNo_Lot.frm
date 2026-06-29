VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmChangeBatchNo_Lot 
   Caption         =   "Change Batch No."
   ClientHeight    =   9135
   ClientLeft      =   1170
   ClientTop       =   450
   ClientWidth     =   13800
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
   ScaleWidth      =   13800
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
      Left            =   75
      TabIndex        =   7
      Top             =   0
      Width           =   13695
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Batch No"
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
         Left            =   5325
         TabIndex        =   8
         Top             =   135
         Width           =   2685
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Batch No"
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
         Left            =   5340
         TabIndex        =   9
         Top             =   150
         Width           =   2685
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8550
      Left            =   75
      TabIndex        =   0
      Top             =   555
      Width           =   13695
      Begin VB.TextBox txtEdit 
         Height          =   390
         Left            =   6225
         TabIndex        =   10
         Top             =   3315
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   435
         Left            =   9675
         TabIndex        =   6
         Top             =   7965
         Width           =   1275
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   435
         Left            =   11025
         TabIndex        =   5
         Top             =   7965
         Width           =   1275
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   435
         Left            =   12360
         TabIndex        =   4
         Top             =   7965
         Width           =   1275
      End
      Begin ComboList.Usercontrol1 cmbMaterial 
         Height          =   285
         Left            =   45
         TabIndex        =   2
         Top             =   360
         Width           =   13590
         _ExtentX        =   23971
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
      Begin MSComctlLib.ListView LV 
         Height          =   7245
         Left            =   45
         TabIndex        =   1
         Top             =   675
         Width           =   13590
         _ExtentX        =   23971
         _ExtentY        =   12779
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Lot No."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Batch #"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Mill Cert."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Remarks"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
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
         TabIndex        =   3
         Top             =   150
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmChangeBatchNo_Lot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bDontStartRefreshing As Boolean
Dim iColNo As Integer
Dim bSample As Boolean

Private Sub RefreshList()
    
    'If cmbVend.MatchFound = False Then Exit Sub
    If cmbMaterial.MatchFound = False Then Exit Sub
    
    If bDontStartRefreshing Then Exit Sub

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim strSQL As String
    Dim i As Integer
    Dim strCond As String
    
    
    If cmbMaterial.ID <> "0" Then
        strCond = " WHERE ItemID='" & cmbMaterial.ID & "'"
    End If
'    If cmbVend.ID <> "0" Then
'        strCond = strCond & " AND VendID='" & cmbVend.ID & "'"
'    End If
'
'    If bSample Then
'        strCond = strCond & " AND Sample=" & Abs(bSample)
'    End If
'
    With rs
        strSQL = "SELECT * FROM Lots_List " & strCond & " ORDER BY EntryID"
        .Open strSQL, con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
        
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, , !LotNo & "")
            ITM.Tag = !EntryID & ""
            
            ITM.ListSubItems.add , , !Batch_No & ""
            ITM.ListSubItems.add , , !Mill_Certificate_No & ""
            ITM.ListSubItems.add , , !Lot_Remarks & ""
            
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ToolTipText = ITM.ListSubItems(i).Text
            Next
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub

err:
   MsgBox err.Description
   
End Sub

Private Sub cmbMaterial_matched()
    Call RefreshList
End Sub



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    If cmbMaterial.MatchFound = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Item_Lots.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "{Lots_List.ItemID}='" & cmbMaterial.ID & "'", rpt
    
End Sub

Private Sub cmdSave_Click()
    
    If Saved Then
        Call RefreshList
    End If
    
End Sub



Private Sub Form_Load()
    
    bDontStartRefreshing = True
'    cmbVend.AddItem "<All Venders>", "0"
'    cmbVend.AddVals con, "AccTitle", "VVenders", "AccNo"
'    cmbVend.ID = "0"
    
    
     
'    Call AddToCombo(cmbStore, "StoreName", "Stores", " ORDER BY StoreName", , "EntryID")
    
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
    Next
    bDontStartRefreshing = False
    
    cmbMaterial.AddVals con, "'{' + ItemID + '} ' + ItemName", "Items", "ItemID", " ORDER BY ItemID"
    
End Sub

Private Sub LV_DblClick()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    If iColNo < 2 Then iColNo = 2
    
    txtEdit.Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
    txtEdit.Text = (LV.SelectedItem.SubItems(iColNo - 1))
    txtEdit.Visible = True
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(txtEdit.Text)
    txtEdit.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        With LV.ColumnHeaders(i)
            If X > Val(.Left) And X < (Val(.Left) + Val(.Width)) Then
                iColNo = i
                Exit For
            End If
        End With
    Next
    
End Sub

Private Function Saved() As Boolean

    Saved = False
    On Error GoTo err
    
    Call StartTrans(con)
    
    Dim i As Integer
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
        
            'If Val(.SubItems(7)) > 0 Then
                con.Execute "UPDATE Lots_List SET Batch_No='" & .SubItems(1) & "',Mill_Certificate_No='" & .SubItems(2) & "',Lot_Remarks='" & .SubItems(3) & "' WHERE EntryID=" & Val(.Tag)
'                con.Execute "INSERT INTO MaterialLocationwiseStatus(Rcvd_RefID,Shelf_RefID,QtyPlaced,UserName,MachineName,BatchNo,LotNo) VALUES(" & _
                 Val(.Tag) & "," & .ListSubItems(10).Tag & "," & Val(.SubItems(7)) & ",'" & CurrentUserName & "','" & strComputerName & "','" & (.SubItems(11)) & "','" & .SubItems(12) & "')"
                 
                
                
            'End If
            
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


Public Sub ShowMe(Optional p_bSample As Boolean = False)

    bSample = p_bSample
        
        
    'cmbMaterial.AddItem "<All Materials>", "0"
'    If bSample Then
'        cmbMaterial.AddVals con, "ItemName + '     {' + ItemID + '}'", "Items", "ItemID", " WHERE InActive=0 AND Sample=1 ORDER BY ItemName"
'    Else
'        cmbMaterial.AddVals con, "'{' + RMID1 + '} ' + RMName", "RM", "RMID1", " ORDER BY RMID1"
'    End If
'    cmbMaterial.ID = "0"
    cmbMaterial.AddVals con, "'{' + RMID1 + '} ' + RMName", "RM", "RMID1", " ORDER BY RMID1"
'    Call RefreshList
    
    Me.Show
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(iColNo - 1) = txtEdit
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub
