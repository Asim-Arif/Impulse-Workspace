VERSION 5.00
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmWastageTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wastage Type"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7620
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
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComboList.Usercontrol1 cmbRepairs 
      Height          =   285
      Left            =   1410
      TabIndex        =   1
      Top             =   240
      Width           =   5310
      _ExtentX        =   9366
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
   Begin VB.OptionButton Opt 
      Caption         =   "Charge to Ask"
      Height          =   315
      Index           =   3
      Left            =   5430
      TabIndex        =   13
      Top             =   900
      Width           =   1935
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Charge to Originator"
      Height          =   315
      Index           =   2
      Left            =   3570
      TabIndex        =   12
      Top             =   900
      Width           =   1935
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Charge to Vender"
      Height          =   315
      Index           =   1
      Left            =   1890
      TabIndex        =   11
      Top             =   900
      Width           =   1635
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Charge to Store"
      Height          =   315
      Index           =   0
      Left            =   210
      TabIndex        =   10
      Tag             =   "0"
      Top             =   900
      Value           =   -1  'True
      Width           =   1635
   End
   Begin VB.Frame FraSave 
      Height          =   630
      Left            =   4290
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   3240
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   345
         Left            =   90
         TabIndex        =   8
         Top             =   180
         Width           =   1485
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   1620
         TabIndex        =   7
         Top             =   180
         Width           =   1485
      End
   End
   Begin VB.Frame FraAdd 
      Height          =   630
      Left            =   2790
      TabIndex        =   2
      Top             =   1920
      Width           =   4740
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   345
         Left            =   1635
         TabIndex        =   5
         Top             =   180
         Width           =   1485
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   345
         Left            =   3165
         TabIndex        =   4
         Top             =   180
         Width           =   1485
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   345
         Left            =   90
         TabIndex        =   3
         Top             =   180
         Width           =   1485
      End
   End
   Begin VB.TextBox txtRepairType 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1410
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   5310
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Wastage Type :"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   270
      Width           =   1155
   End
End
Attribute VB_Name = "frmWastageTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbRepairs_matched()
    
    Dim i As Integer, ITM As ListItem
'    For i = 1 To LV.ListItems.count
'        LV.ListItems(i).Checked = False
'    Next
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM WastageTypes WHERE EntryID=" & cmbRepairs.ID, con, adOpenForwardOnly, adLockReadOnly
'        LV.ListItems.Clear
'        Do Until .EOF
'            Set ITM = LV.ListItems.add(, !ProcessID & "'", !Description & "")
'            ITM.ListSubItems.add , , !code & ""
'            ITM.Checked = True
'            .MoveNEXT
'        Loop
        Opt(Val(!WastageType)).Value = True
        .Close
        
    End With
    Set rs = Nothing
    
    
    
End Sub

Private Sub cmdAdd_Click()
'    Call FillProcesses
    Call ChangeState(True)
End Sub

Private Sub ChangeState(Optional bAdd As Boolean = False, Optional bEdit As Boolean = False)

    Dim bAddEdit As Boolean
    bAddEdit = bAdd Or bEdit
    
    FraAdd.Visible = Not bAddEdit
    FraSave.Visible = bAddEdit
    cmbRepairs.Visible = Not bAddEdit
    
    txtRepairType.Visible = bAddEdit
    
    Dim i As Integer
    If bAdd Then
'        For i = 1 To LV.ListItems.count
'            LV.ListItems(i).Checked = False
'        Next
        txtRepairType = ""
        txtRepairType.Tag = ""
    ElseIf bEdit Then
        txtRepairType = cmbRepairs.Text
        txtRepairType.Tag = cmbRepairs.ID
    End If
    
End Sub

Private Sub cmdCancel_Click()
    If cmbRepairs.MatchFound Then
        Call cmbRepairs_matched
    End If
    Call ChangeState
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    
    If cmbRepairs.MatchFound = False Then Exit Sub
'    Call FillProcesses
'    Dim rs As New ADODB.Recordset
'    With rs
'        .Open "SELECT * FROM RepairTypeProcesses WHERE Repair_RefID=" & cmbRepairs.ID, con, adOpenForwardOnly, adLockReadOnly
'        Do Until .EOF
'            LV.ListItems(!ProcessID & "'").Checked = True
'            .MoveNEXT
'        Loop
'        .Close
'    End With
'    Set rs = Nothing
    
    Call ChangeState(False, True)
    
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
    Dim lEntryID As Long
    Call StartTrans(con)
    If Val(txtRepairType.Tag) > 0 Then
        con.Execute "UPDATE WastageTypes SET WastageName='" & txtRepairType & "',WastageType=" & Opt(0).Tag & " WHERE EntryID=" & Val(txtRepairType.Tag)
        lEntryID = Val(txtRepairType.Tag)
    Else
        con.Execute "INSERT INTO WastageTypes(WastageName,WastageType) VALUES('" & txtRepairType & "'," & Opt(0).Tag & ")"
        lEntryID = GetSingleLongValue("MAX(EntryID)", "RepairTypes")
    End If
    
    Dim iSelCount As Long, i As Integer
'    For i = 1 To LV.ListItems.count
'        If LV.ListItems(i).Checked Then
'            iSelCount = iSelCount + 1
'        End If
'    Next
    
    con.CommitTrans
        
    Call FillRepairTypes
    cmbRepairs.ID = lEntryID
    
    Call ChangeState
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    
'    Call FillProcesses
    
    Call FillRepairTypes
    
End Sub

Private Sub FillProcesses()
    
    Exit Sub
'    Dim rs As New ADODB.Recordset, ITM As ListItem
'    With rs
'        .Open "SELECT * FROM Processes", con, adOpenForwardOnly, adLockReadOnly
'        LV.ListItems.Clear
'        Do Until .EOF
'            Set ITM = LV.ListItems.add(, !ProcessID & "'", !Description & "")
'            ITM.ListSubItems.add , , !code & ""
'            .MoveNEXT
'        Loop
'        .Close
'    End With
'    Set rs = Nothing
    
End Sub

Private Sub FillRepairTypes()

    cmbRepairs.ClearVals
    cmbRepairs.AddVals con, "WastageName", "WastageTypes", "EntryID", " ORDER BY WastageName"
        
End Sub

Private Sub Opt_Click(Index As Integer)
    Opt(0).Tag = Index
End Sub
