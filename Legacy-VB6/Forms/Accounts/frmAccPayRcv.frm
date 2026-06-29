VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAccPayRcv 
   ClientHeight    =   8625
   ClientLeft      =   165
   ClientTop       =   75
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   480
      Left            =   0
      TabIndex        =   9
      Top             =   8100
      Width           =   5475
      Begin VB.ComboBox cmbfield 
         Height          =   315
         ItemData        =   "frmAccPayRcv.frx":0000
         Left            =   60
         List            =   "frmAccPayRcv.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   135
         Width           =   1875
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Enabled         =   0   'False
         Height          =   300
         Left            =   4395
         TabIndex        =   11
         Top             =   135
         Width           =   1020
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1965
         TabIndex        =   10
         Top             =   135
         Width           =   2430
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox PicEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1215
      ScaleHeight     =   225
      ScaleWidth      =   5805
      TabIndex        =   5
      Top             =   1695
      Visible         =   0   'False
      Width           =   5835
      Begin VB.ComboBox cmbCrDr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmAccPayRcv.frx":0060
         Left            =   5250
         List            =   "frmAccPayRcv.frx":006A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   -45
         Width           =   600
      End
      Begin VB.TextBox txtAccTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "RptTime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   -15
         MaxLength       =   150
         TabIndex        =   0
         Top             =   0
         Width           =   3510
      End
      Begin VB.TextBox txtBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "RptTime"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3540
         MaxLength       =   12
         TabIndex        =   1
         Top             =   0
         Width           =   1680
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1635
      Top             =   3255
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccPayRcv.frx":0076
            Key             =   "ACC"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccPayRcv.frx":04CA
            Key             =   "PARIN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccPayRcv.frx":05DE
            Key             =   "PAR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccPayRcv.frx":06F2
            Key             =   "ACCIN"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   7695
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   13573
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
      BorderStyle     =   1
      Appearance      =   1
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
         Text            =   "Account #"
         Object.Width           =   4895
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account Title"
         Object.Width           =   10477
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Opening Bal."
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Current Bal."
         Object.Width           =   3175
      EndProperty
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   390
      Left            =   10230
      TabIndex        =   8
      Top             =   8190
      Width           =   1560
      ForeColor       =   64
      Caption         =   "Close            "
      PicturePosition =   327683
      Size            =   "2752;688"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Index           =   3
      Left            =   9210
      TabIndex        =   7
      Top             =   7845
      Width           =   645
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E7EBEF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10170
      TabIndex        =   6
      Top             =   7845
      Width           =   1605
   End
   Begin MSForms.CommandButton cmdOpt 
      Height          =   390
      Left            =   8640
      TabIndex        =   4
      Top             =   8190
      Width           =   1560
      Caption         =   "Options         "
      PicturePosition =   327683
      Size            =   "2752;688"
      Picture         =   "frmAccPayRcv.frx":0B46
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Begin VB.Menu mnurefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "&Show Inactive Accounts"
      End
      Begin VB.Menu mnuZero 
         Caption         =   "Show Accounts With Zero Balance"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLedger 
         Caption         =   "View &Ledger For Last 30 Days."
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActive 
         Caption         =   "Make This Account Inactive"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAction 
         Caption         =   "Make &Journal Entry"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print The Accounts"
      End
   End
End
Attribute VB_Name = "frmAccPayRcv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PermissionDenied As Boolean

Dim Rec As New ADODB.Recordset
Dim rsAccounts As New ADODB.Recordset
Dim i As Integer
Dim Bal As Double
Dim AccList As String
Dim SubAccList As String
Dim TotalBalance As Double
Dim TotalOpenBalance As Double
Dim ShowAll As Boolean

Private Sub Command2_Click()
    
    Dim AccNo As String
    Dim AccTitle As String
    Dim Bal As Double

    For i = 1 To LV.ListItems.count
        AccNo = LV.ListItems(i).Text & ""
        AccTitle = LV.ListItems(i).ListSubItems(1).Text
        Bal = Val(LV.ListItems(i).ListSubItems(2).Tag)
        con.Execute "Delete * From TempChaccount"
        con.Execute "Insert into TempChAccount Values('" & AccNo & "','" & AccTitle & "'," & Bal & ")"
    Next i

End Sub

Private Sub RefreshList()

    Screen.MousePointer = vbHourglass
    LV.Sorted = False

    Dim PrevIndex As Double

    If LV.ListItems.count > 0 Then PrevIndex = LV.SelectedItem.Index

With rsAccounts
    If .RecordCount = 0 Then Exit Sub
    
    LV.Visible = False
    LV.ListItems.Clear
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    .MoveFirst
    For i = 0 To .RecordCount - 1
        TotalBalance = 0
        TotalOpenBalance = 0
        AccList = ""
        
        If ![Parent] Then
            Call CalculateBalance(.Fields("AccNo"))
        Else
            TotalBalance = Val(![Balance] & "")
            TotalOpenBalance = Val(![OpenBal] & "")
        End If
        
        Call AddToList(i, .Fields("SubaccOf") & "", .Fields("AccNo"), .Fields("AccTitle"), TotalOpenBalance, TotalBalance, .Fields("AccNo"), ![Parent], ![Active])
        .MoveNEXT
    Next i
End With
   LV.Visible = True
   LV.Sorted = True
If PrevIndex > 0 And PrevIndex <= LV.ListItems.count Then LV.ListItems(PrevIndex).Selected = True
   Screen.MousePointer = vbDefault
If LV.ListItems.count = 0 Then Exit Sub
   lblTotal = LV.ListItems("[" & Me.Tag).SubItems(3)
   'Me.Caption = Lv.ListItems("[" & Me.Tag).SubItems(1)

End Sub

Private Sub cmdFind_Click()

    Dim StartIdx As Integer

    If LV.SelectedItem Is Nothing Then
       StartIdx = 1
    Else
       If LV.SelectedItem.Index = LV.ListItems.count Then
          StartIdx = 1
       Else
          StartIdx = LV.SelectedItem.Index
       End If
    End If

    Dim fldValue As String




    For i = StartIdx To LV.ListItems.count
       
       
       If cmbfield.ListIndex = 0 Then
          fldValue = Trim(Replace(LV.ListItems(i), "»", ""))
       Else
          fldValue = Trim(Replace(LV.ListItems(i).SubItems(cmbfield.ListIndex), "»", ""))
       End If
    
       If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
          LV.ListItems(i).Selected = True
          LV.ListItems(i).EnsureVisible
          If i <> StartIdx Then Exit For
       End If
    Next
    
    If i = LV.ListItems.count + 1 Then
        MsgBox "No Mathes Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If

    

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    'Me.Move LEFTMARGIN, TOPMARGIN, FORMWIDTH, FORMHEIGHT

    LV.Move 50, 50, Me.ScaleWidth - 100, Me.ScaleHeight - (cmdClose.Height + 100)
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, LV.Top + LV.Height + 50
    cmdOpt.Move cmdClose.Left - cmdOpt.Width - 50, cmdClose.Top
    Frame1.Move LV.Left, LV.Top + LV.Height - 25
    
End Sub

Private Sub mnuPrint_Click()

    Dim Inactive As Boolean
    Screen.MousePointer = vbHourglass
    con.BeginTrans
    con.Execute "Delete From TempAccounts"
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .SmallIcon = "PARIN" Or .SmallIcon = "ACCIN" Then Inactive = True Else Inactive = False
            con.Execute "Insert into TempAccounts Values('" & .Text & "','" & .SubItems(1) & _
                        "','" & Left(.Text, 2) & "','" & .ListSubItems(2) & "','" & .ListSubItems(3) & "'," & Abs(.Bold) & "," & Abs(Inactive) & ")"
        End With
    Next
    
    con.CommitTrans
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Accounts.rpt")
    
    rpt.FormulaFields.GetItemByName("Company").Text = "'" & strCompany & "'"
    rpt.FormulaFields.GetItemByName("Caption").Text = "'" & Me.Caption & "'"
    rpt.FormulaFields.GetItemByName("PrintBalance").Text = True
    rpt.FormulaFields.GetItemByName("Address").Text = "'" & COMPANYADDRESS & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuShowAll_Click()
   mnuShowAll.Checked = Not mnuShowAll.Checked
   ShowAll = mnuShowAll.Checked
   Call mnuRefresh_Click
End Sub
Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdOpt_Click()
With cmdOpt
 Call CheckMenu
 Me.PopUpMenu mnupop, 255, .Left, .Top + .Height
End With
End Sub

Private Sub Form_Activate()

'Check For User Access
    If PermissionDenied Then
        Me.Hide
        MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
        Unload Me
        Exit Sub
    End If

    Dim Condition As String
    If ShowAll Then
        Condition = ""
    Else
        Condition = " AND Active=1 "
    End If

    If rsAccounts.State <> 0 Then Exit Sub
    rsAccounts.Open "Select * from Accounts Where (Accno='" & Me.Tag & "' OR SubAccOf Like '" & Me.Tag & "%') " & Condition & " Order By Accno", con, adOpenForwardOnly
  
   
    Call RefreshList
   
End Sub

Private Sub Form_Load()
    
    mnupop.Visible = False
    'PermissionDenied = Not UserHasAccess("AccountsPayableReceivable")
    Set LV.SmallIcons = ImageList1
    cmbfield.Clear
    For i = 1 To LV.ColumnHeaders.count
        cmbfield.AddItem LV.ColumnHeaders(i).Text
    Next

End Sub
Private Sub AddToList(ByVal No As Integer, Parent As String, code As String, AccTitle As String, OpenBal, Bal As Double, key As String, ParentAccount As Boolean, Active As Boolean)

If (Not mnuZero.Checked) And (Bal = 0) And (Not ParentAccount) Then Exit Sub


Dim LColor As Long
Dim NumSpaces As Integer
Dim TheTag As String

' Format Opening Balance

Dim Openbalance As String

   Openbalance = Format(OpenBal, "#,##0.00 DR;#,##0.00 CR;#,##0.00      ")


'Format current Balance


Dim CurrBalance As String

      CurrBalance = Format(Bal, "#,##0.00 DR;#,##0.00 CR;#,##0.00      ")



If Not Active Then
   LColor = &H808080
   TheTag = "Inactive"
Else
   LColor = &H800000
   TheTag = "Active"
End If

'For Image Selection
Dim img As String
If ParentAccount Then
   If Active Then
     img = "PAR"
   Else
     img = "PARIN"
   End If
Else
   If Active Then
     img = "ACC"
   Else
     img = "ACCIN"
   End If
End If
   
   
   
   

Dim i As Integer




For i = 1 To LV.ListItems.count
    If Parent = Replace(LV.ListItems(i).key, "[", "") Then
    
        NumSpaces = Len(LV.ListItems(i).ListSubItems(1).Text) - Len(Trim(LV.ListItems(i).ListSubItems(1).Text))
         
         
        

        
        LV.ListItems.add(i + 1, "[" & key, code, , img).Bold = ParentAccount
        LV.ListItems(i + 1).Tag = TheTag
        
        LV.ListItems(i + 1).ForeColor = LColor
        LV.ListItems(i + 1).ListSubItems.add , , Space$(5 + NumSpaces) & "»  " & AccTitle
        LV.ListItems(i + 1).ListSubItems(1).ForeColor = LColor
        LV.ListItems(i + 1).ListSubItems(1).Bold = ParentAccount
        LV.ListItems(i + 1).ListSubItems.add , , Openbalance
        LV.ListItems(i + 1).ListSubItems(2).Tag = OpenBal
        LV.ListItems(i + 1).ListSubItems(2).ForeColor = LColor
        LV.ListItems(i + 1).ListSubItems(2).Bold = ParentAccount
        LV.ListItems(i + 1).ListSubItems.add , , CurrBalance
        LV.ListItems(i + 1).ListSubItems(3).ForeColor = LColor
        LV.ListItems(i + 1).ListSubItems(3).Tag = Bal
        LV.ListItems(i + 1).ListSubItems(3).Bold = ParentAccount
        Exit Sub
    End If
Next i


No = LV.ListItems.add(, "[" & key, code, , img).Index - 1

LV.ListItems(No + 1).Tag = TheTag
LV.ListItems(No + 1).Bold = True
LV.ListItems(No + 1).ForeColor = LColor
LV.ListItems(No + 1).ListSubItems.add , , AccTitle
LV.ListItems(No + 1).ListSubItems(1).ForeColor = LColor
LV.ListItems(No + 1).ListSubItems(1).Bold = ParentAccount

LV.ListItems(No + 1).ListSubItems.add , , Openbalance
LV.ListItems(No + 1).ListSubItems(2).Bold = ParentAccount
LV.ListItems(No + 1).ListSubItems(2).ForeColor = LColor
LV.ListItems(No + 1).ListSubItems(2).Tag = OpenBal

LV.ListItems(No + 1).ListSubItems.add , , CurrBalance
LV.ListItems(No + 1).ListSubItems(3).Bold = ParentAccount
LV.ListItems(No + 1).ListSubItems(3).ForeColor = LColor
LV.ListItems(No + 1).ListSubItems(3).Tag = Bal

End Sub
Private Sub CreateLists(PRec As ADODB.Recordset, AccNo As String)

'If Len(AccList) > 0 Then AccList = AccList & "'" & AccNo & "',"
AccList = AccList & "'" & AccNo & "',"
SubAccList = ""
Dim i As Integer
With PRec
    For i = 0 To .RecordCount - 1
        AccList = AccList & "'" & .Fields("AccNo") & "',"
        SubAccList = SubAccList & "'" & .Fields("AccNo") & "',"
        .MoveNEXT
    Next i
    
End With
If Right$(SubAccList, 1) = "," Then SubAccList = Left$(SubAccList, Len(SubAccList) - 1)

End Sub

Private Sub CalculateBalance(AccNo As String)

Dim i As Integer
With Rec

    If .State = 1 Then .Close
    .CursorLocation = adUseClient
    .Open "Select AccNo from Accounts Where SubAccOf='" & AccNo & "'", con, adOpenForwardOnly, adLockReadOnly
    Do While .RecordCount > 0
        Call CreateLists(Rec, AccNo)
        If .State = 1 Then .Close
    .CursorLocation = adUseClient
    'If Right$(SubAccList, 1) = "," Then
    '    Useable = Left$(SubAccList, Len(SubAccList) - 1)
    'Else
    '    Useable = SubAccList
    'End If
    .Open "Select AccNo from Accounts Where SubAccOf in (" & SubAccList & ")", con, adOpenForwardOnly, adLockReadOnly
    Loop
    If Len(AccList) > 0 Then
        AccList = Left$(AccList, Len(AccList) - 1)
    Else
        AccList = "'" & AccNo & "'"
    End If
    If .State = 1 Then .Close
    .CursorLocation = adUseClient
    .Open "Select Sum(Balance),sum(OpenBal) from Accounts Where AccNo in (" & AccList & ")", con, adOpenForwardOnly, adLockReadOnly
    TotalBalance = Val(.Fields(0) & "")
    TotalOpenBalance = Val(.Fields(1) & "")
    
End With
    
End Sub

Private Sub Form_Unload(cancel As Integer)

    Set rsAccounts = Nothing

End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If LV.ListItems.count = 0 Then Exit Sub
        Call CheckMenu
        Me.PopUpMenu mnupop, , (LV.Left + (LV.Width / 2)), (LV.Top + LV.SelectedItem.Top)
    End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Or LV.ListItems.count = 0 Then Exit Sub
    Call CheckMenu
    Me.PopUpMenu mnupop
 
End Sub


Private Sub mnuAction_Click()

    If LV.SelectedItem.Tag = "Inactive" Then
       MsgBox "No Transaction Can Be Made In An Inactive Account.", vbInformation
       Exit Sub
    End If
    
    Call CheckMenu
    Dim f As Form
     If mnuAction.Tag = "24" Then
       Set f = New frmCashPV
     ElseIf mnuAction.Tag = "15" Then
       Set f = New frmCashRV
     End If
     
    With f
       Call FindInCombo(.cmbAccNo, LV.SelectedItem)
       .Show
    End With
End Sub

Private Sub mnuLedger_Click()
Dim f As New frmLedger
With f
   .Show
   Call FindInCombo(.cmbAccNo, LV.SelectedItem)
   Call .cmdRefresh_Click
End With
End Sub



Private Sub mnuRefresh_Click()
   Screen.MousePointer = vbHourglass
   If rsAccounts.State = 1 Then rsAccounts.Close
   Set rsAccounts = Nothing
   Call Form_Activate
   Screen.MousePointer = vbDefault
End Sub




Private Sub CheckMenu()

   If OutOfFinancialYear Then
      mnuAction.Enabled = False
      mnuActive.Enabled = False
      Exit Sub
   End If


If Len(Me.Tag) < 2 Then Exit Sub
   
   Dim Acc As Integer
   Acc = Left(Me.Tag, 2)
   mnuLedger.Enabled = True = Not (LV.SelectedItem.SmallIcon = "PAR" Or LV.SelectedItem.SmallIcon = "PARIN")
     If Acc = "24" Then
      mnuAction.Caption = "Enter &Payment Voucher"
      mnuAction.Tag = Acc
     Else
       mnuAction.Caption = "Enter &Receipt Voucher"
       mnuAction.Tag = Acc
     End If
     
     
     
     
     If LV.SelectedItem.Tag = "Inactive" Then
       mnuActive.Caption = "Make This Account A&ctive."
     Else
       mnuActive.Caption = "Make This Account Ina&ctive."
     End If
     mnuActive.Tag = LV.SelectedItem.Tag
     
     
     
     mnuAction.Enabled = LV.ListItems.count > 1
     mnuLedger.Enabled = LV.ListItems.count > 1
   
End Sub

Private Sub mnuActive_Click()



   If mnuActive.Tag = "Active" Then
      If Not MakeAccountInActive Then Exit Sub
   ElseIf mnuActive.Tag = "Inactive" Then
      If Not MakeAccountActive Then Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
'   Sleep (500)
   Call mnuRefresh_Click
   Screen.MousePointer = vbDefault
End Sub


Private Function MakeAccountInActive() As Boolean
MakeAccountInActive = False
 With LV.SelectedItem
   
   If .ListSubItems(3).Tag <> 0 Then
        MsgBox "This Account Cannot Be Made Inactive." & vbNewLine & _
           "To Make An Account Inactive Balance Must Be Zero.", vbInformation
         Exit Function
   ElseIf .SmallIcon = "PAR" Then
       If MsgBox("This Is A Master Account, All Sub Accounts Will Also Become Inactive." & vbNewLine & _
           "Are You Sure to Make '" & Trim(Replace(LV.SelectedItem.SubItems(1), "»", "")) & "' Inactive", vbYesNo + vbQuestion) = vbNo Then Exit Function
   Else
      If MsgBox("Are You Sure to Make '" & Trim(Replace(LV.SelectedItem.SubItems(1), "»", "")) & "' Inactive", vbYesNo + vbQuestion) = vbNo Then Exit Function
   End If
   
   
   
   
   con.Execute "Update Accounts Set Active=0 Where Accno='" & Trim(.Text) & _
               "' OR SubAccOf Like '" & Trim(.Text) & "'"
  
   MakeAccountInActive = True
End With

End Function


Private Function MakeAccountActive() As Boolean

MakeAccountActive = False
Dim AlsoChild As String

 With LV.SelectedItem
   
   If .SmallIcon = "PARIN" Then
       If MsgBox("This Is A Parent Account, Do You Want All Child Accounts To Become Active." _
          , vbYesNo + vbQuestion) = vbYes Then
          AlsoChild = " OR SubAccOf Like '" & Trim(.Text) & "'"
       Else
           AlsoChild = ""
       End If
   End If
   
   
   
   
   con.Execute "Update Accounts Set Active=1 Where Accno='" & Trim(.Text) & "'" & AlsoChild
  
MakeAccountActive = True
End With




End Function



Private Function UpdateAccount() As Boolean
 If txtAccTitle.Text = "" Then
   MsgBox "Invalid Account Title", vbInformation
   Exit Function
 End If
 
Dim Diff As Double

Screen.MousePointer = vbHourglass
Diff = Val(txtBal.Tag) - Val(cmbCrDr.Tag)
   con.Execute "Update Accounts Set Acctitle='" & txtAccTitle & "',OpenBal=" & Val(cmbCrDr.Tag) & _
               ", Balance=Balance+" & Diff & " Where AccNo='" & LV.SelectedItem.Text & "'"
               
   'Adjust The Ledger
   Call adjustLedger(LV.SelectedItem.Text, YearStartDate, 0, Diff, True)
   
   'Sleep (500)
   Call mnuRefresh_Click
Screen.MousePointer = vbDefault
   
End Function


Private Sub mnuZero_Click()
   
   mnuZero.Checked = Not mnuZero.Checked
   Call mnuRefresh_Click
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = Len(Trim(txtFind)) > 0
End Sub
