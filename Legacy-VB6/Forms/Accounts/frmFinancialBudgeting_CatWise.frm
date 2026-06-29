VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFinancialBudgeting_CatWise 
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   180
   ClientWidth     =   13905
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   13905
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtEdit 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   4305
      TabIndex        =   7
      Top             =   2730
      Visible         =   0   'False
      Width           =   850
   End
   Begin MSComctlLib.ListView LV 
      Height          =   6720
      Left            =   105
      TabIndex        =   5
      Top             =   1185
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   11853
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S#"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account #"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Account Title"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "Month_July"
         Text            =   "Jul"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "Month_August"
         Text            =   "Aug"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "Month_September"
         Text            =   "Sep"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "Month_October"
         Text            =   "Oct"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "Month_November"
         Text            =   "Nov"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Object.Tag             =   "Month_December"
         Text            =   "Dec"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Object.Tag             =   "Month_January"
         Text            =   "Jan"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Object.Tag             =   "Month_February"
         Text            =   "Feb"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Tag             =   "Month_March"
         Text            =   "Mar"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Object.Tag             =   "Month_April"
         Text            =   "Apr"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Object.Tag             =   "Month_May"
         Text            =   "May"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Object.Tag             =   "Month_June"
         Text            =   "Jun"
         Object.Width           =   1499
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13890
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Financial Budgeting"
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
         Index           =   8
         Left            =   75
         TabIndex        =   3
         Top             =   165
         Width           =   13770
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Financial Budgeting"
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
         Height          =   465
         Index           =   9
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   13755
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   480
      Left            =   11925
      TabIndex        =   1
      Top             =   7995
      Width           =   1920
   End
   Begin MSForms.ComboBox cmbFinYear 
      Height          =   315
      Left            =   1485
      TabIndex        =   6
      Top             =   825
      Width           =   3810
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "6720;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Financial Year :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   870
      Width           =   1260
   End
End
Attribute VB_Name = "frmFinancialBudgeting_CatWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iColNo As Integer

Private Sub cmbFinYear_Change()
    Call cmbFinYear_Click
End Sub

Private Sub cmbFinYear_Click()

    On Error GoTo err
    
    If cmbFinYear.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM Accounts WHERE Parent=1 AND Active=1 ORDER BY AccNo", con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
        
        Do Until .EOF
        
            Set ITM = LV.ListItems.add(, !AccNo & "", LV.ListItems.count + 1)
            ITM.ListSubItems.add , , !AccNo & ""
            ITM.ListSubItems.add , , !AccTitle & ""
            
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , "0"

            .MoveNext
            
        Loop
        .Close
        
        .Open "SELECT * FROM FinancialBudgeting WHERE FinancialYear='" & cmbFinYear.List(cmbFinYear.ListIndex, 1) & "'", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            On Error Resume Next
            Set ITM = LV.ListItems(!AccNo & "")
            If err.Number = 35601 Then GoTo GoToNextRecord
            
            ITM.SubItems(3) = Val(!Month_July & "")
            ITM.SubItems(4) = Val(!Month_August & "")
            ITM.SubItems(5) = Val(!Month_September & "")
            ITM.SubItems(6) = Val(!Month_October & "")
            ITM.SubItems(7) = Val(!Month_November & "")
            ITM.SubItems(8) = Val(!Month_December & "")
            ITM.SubItems(9) = Val(!Month_January & "")
            ITM.SubItems(10) = Val(!Month_February & "")
            ITM.SubItems(11) = Val(!Month_March & "")
            ITM.SubItems(12) = Val(!Month_April & "")
            ITM.SubItems(13) = Val(!Month_May & "")
            ITM.SubItems(14) = Val(!Month_June & "")
GoToNextRecord:
            .MoveNext
        Loop
        .Close
        
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call FillCmbs
End Sub

Private Sub FillCmbs()

    cmbFinYear.Clear
    Dim iStartYear As Integer
    
    iStartYear = 2010
    Dim i As Integer
    For i = iStartYear To year(Date)
        cmbFinYear.AddItem i & " - " & i + 1 & " [ " & Format(DateSerial(i, 7, 1), "dd-MMM-yyyy") & " to " & Format(DateSerial(i + 1, 6, 30), "dd-MMM-yyyy") & " ]"
        cmbFinYear.List(cmbFinYear.ListCount - 1, 1) = i & " - " & (i + 1)
    Next i
    
End Sub

Private Sub LV_DblClick()

    If iColNo < 4 Then iColNo = 4
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
        .Text = Val(LV.SelectedItem.SubItems(iColNo - 1))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 13 Then
        iColNo = 4
        Call LV_DblClick
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

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Dim iRecAffect As Integer, strFinYear As String
        
        strFinYear = cmbFinYear.List(cmbFinYear.ListIndex, 1)
        
        con.Execute "UPDATE FinancialBudgeting SET " & LV.ColumnHeaders(iColNo).Tag & "=" & Val(txtEdit) & _
         " WHERE FinancialYear='" & strFinYear & "' AND AccNo='" & LV.SelectedItem.SubItems(1) & "'", iRecAffect
        If iRecAffect = 0 Then
            con.Execute "INSERT INTO FinancialBudgeting(FinancialYear,AccNo," & LV.ColumnHeaders(iColNo).Tag & _
             ") VALUES('" & strFinYear & "','" & LV.SelectedItem.SubItems(1) & "'," & Val(txtEdit) & ")"
        End If
        LV.SelectedItem.SubItems(iColNo - 1) = Val(txtEdit)
        If iColNo < LV.ColumnHeaders.count Then 'If Not on Last Column...
            iColNo = iColNo + 1
            LV_DblClick
        Else    'If on Last Column
            If LV.SelectedItem.Index < LV.ListItems.count Then  'If Not On Last Item
                LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                iColNo = 4
                Call LV_DblClick
            Else
                txtEdit.Visible = False
            End If
        End If
        
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub
