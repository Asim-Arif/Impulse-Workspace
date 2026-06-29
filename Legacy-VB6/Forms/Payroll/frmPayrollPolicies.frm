VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmPayrollPolicies 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Policies"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
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
   ScaleHeight     =   5385
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbYesNo 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3330
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   345
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1485
      TabIndex        =   1
      Top             =   2055
      Visible         =   0   'False
      Width           =   1600
   End
   Begin MSComctlLib.ListView LVCat 
      Height          =   4575
      Left            =   225
      TabIndex        =   0
      Top             =   675
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15199215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Setting"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2822
      EndProperty
   End
   Begin MSForms.ComboBox cmbCatagory 
      Height          =   330
      Left            =   1365
      TabIndex        =   4
      Top             =   330
      Width           =   6690
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "11800;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Segoe UI"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Catagory # :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   72
      Left            =   225
      TabIndex        =   2
      Top             =   360
      Width           =   1125
   End
End
Attribute VB_Name = "frmPayrollPolicies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCatagory_Change()
    Call cmbCatagory_Click
End Sub

Private Sub cmbCatagory_Click()

    If cmbCatagory.MatchFound = False Then Exit Sub
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    LVCat.ListItems.Clear
    
    With rs
        .Open "SELECT * FROM EmployeeAttendanceCatagories WHERE EntryID=" & cmbCatagory.List(cmbCatagory.ListIndex, 1), con, adOpenForwardOnly, adLockReadOnly
        
        Set ITM = LVCat.ListItems.add(, , "Lunch Scanning")
        ITM.ListSubItems.add(, , IIf(!LunchTimeScanning, "Yes", "No")).Tag = Abs(!LunchTimeScanning)
        
        Set ITM = LVCat.ListItems.add(, , "Time Deduction")
        ITM.ListSubItems.add(, , IIf(!LateDeduction, "Yes", "No")).Tag = Abs(!LateDeduction)
         
        Set ITM = LVCat.ListItems.add(, , "Time Deduction After")
        ITM.ListSubItems.add(, , !LateDeductionAfter & "").Tag = Val(!LateDeductionAfter & "")
        
        Set ITM = LVCat.ListItems.add(, , "Gate Pass Required")
        ITM.ListSubItems.add(, , IIf(!GatePassRequired, "Yes", "No")).Tag = Abs(!GatePassRequired)
        
        Set ITM = LVCat.ListItems.add(, , "Leaves Recorded")
        ITM.ListSubItems.add(, , IIf(!LeavesRecorded, "Yes", "No")).Tag = Abs(!LeavesRecorded)
        
        Set ITM = LVCat.ListItems.add(, , "Over Time Paid")
        ITM.ListSubItems.add(, , IIf(!OverTimePaid, "Yes", "No")).Tag = Val(!OverTimePaid)
        
        Set ITM = LVCat.ListItems.add(, , "Sandwich Leaves Deduction")
        ITM.ListSubItems.add(, , IIf(!DeductionSpecialAbsents, "Yes", "No")).Tag = Abs(!DeductionSpecialAbsents)
        
        Set ITM = LVCat.ListItems.add(, , "Over Time Rate")
        ITM.ListSubItems.add(, , Val(!OTRate & "")).Tag = Val(!OTRate & "")
        
        Set ITM = LVCat.ListItems.add(, , "Deduction Percentage")
        ITM.ListSubItems.add(, , Val(!DeductionPercentage & "")).Tag = Val(!DeductionPercentage & "")
        
        Set ITM = LVCat.ListItems.add(, , "Saturday Early Out (Mins.)")
        ITM.ListSubItems.add(, , Val(!Saturday_Early_Out_Mins & "")).Tag = Val(!Saturday_Early_Out_Mins & "")
        
        Set ITM = LVCat.ListItems.add(, , "Remarks")
        ITM.ListSubItems.add(, , !Remarks & "").Tag = !Remarks & ""
        
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbYesNo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        LVCat.SelectedItem.SubItems(1) = cmbYesNo.Text
        LVCat.SelectedItem.ListSubItems(1).Tag = cmbYesNo.ListIndex
        
        con.Execute "UPDATE EmployeeAttendanceCatagories SET LunchTimeScanning=" & Val(LVCat.ListItems(1).ListSubItems(1).Tag) & _
         ",LateDeduction=" & Val(LVCat.ListItems(2).ListSubItems(1).Tag) & ",GatePassRequired=" & Val(LVCat.ListItems(4).ListSubItems(1).Tag) & _
         ",LeavesRecorded=" & Val(LVCat.ListItems(5).ListSubItems(1).Tag) & ",OverTimePaid=" & Val(LVCat.ListItems(6).ListSubItems(1).Tag) & _
         ",DeductionSpecialAbsents=" & Val(LVCat.ListItems(7).ListSubItems(1).Tag) & " WHERE EntryID=" & cmbCatagory.List(cmbCatagory.ListIndex, 1)
         
        cmbYesNo.Visible = False
    ElseIf KeyAscii = 27 Then
        cmbYesNo.Visible = False
    End If
    
End Sub

Private Sub cmbYesNo_LostFocus()
    cmbYesNo.Visible = False
End Sub

Private Sub Form_Load()
    
    Call AddToCombo(cmbCatagory, "CAST(EntryID AS VARCHAR(5)) + ' ('  + ISNULL(Remarks,'') + ')'", "EmployeeAttendanceCatagories", "", , "EntryID")
    cmbYesNo.AddItem "No"
    cmbYesNo.AddItem "Yes"
    
End Sub

Private Sub LVCat_DblClick()

    If LVCat.ListItems.count = 0 Then Exit Sub
    Select Case LVCat.SelectedItem.Index
        Case 1, 2, 4, 5, 6, 7
            With cmbYesNo
                .Move LVCat.Left + LVCat.ColumnHeaders(2).Left, LVCat.Top + LVCat.SelectedItem.Top
                .ListIndex = Val(LVCat.SelectedItem.ListSubItems(1).Tag)
                .Visible = True
                .SetFocus
            End With
        Case Else
            With txtEdit
                .Move LVCat.Left + LVCat.ColumnHeaders(2).Left, LVCat.Top + LVCat.SelectedItem.Top
                If LVCat.SelectedItem.Index = 11 Then
                    .Text = LVCat.SelectedItem.SubItems(1)
                Else
                    .Text = Val(LVCat.SelectedItem.ListSubItems(1).Tag)
                End If
                .Visible = True
                .SetFocus
            End With
    End Select
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If LVCat.SelectedItem.Index = 3 Or LVCat.SelectedItem.Index = 11 Then
            LVCat.SelectedItem.SubItems(1) = txtEdit
            LVCat.SelectedItem.ListSubItems(1).Tag = txtEdit
        Else
            LVCat.SelectedItem.SubItems(1) = Val(txtEdit)
            LVCat.SelectedItem.ListSubItems(1).Tag = Val(txtEdit)
        End If
        
        con.Execute "UPDATE EmployeeAttendanceCatagories SET LateDeductionAfter='" & LVCat.ListItems(3).ListSubItems(1).Tag & _
         "',OTRate=" & Val(LVCat.ListItems(8).ListSubItems(1).Tag) & ",DeductionPercentage=" & Val(LVCat.ListItems(9).ListSubItems(1).Tag) & _
         ",Saturday_Early_Out_Mins=" & Val(LVCat.ListItems(10).ListSubItems(1).Tag) & ",Remarks='" & LVCat.ListItems(11).SubItems(1) & "' WHERE EntryID=" & cmbCatagory.List(cmbCatagory.ListIndex, 1)
        
        txtEdit.Visible = False
        
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub
