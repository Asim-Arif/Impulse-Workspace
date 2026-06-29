VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMonthlyAccured 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Accured / PrePaid Expenses"
   ClientHeight    =   6960
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   11130
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2505
      ScaleHeight     =   225
      ScaleWidth      =   3570
      TabIndex        =   12
      Top             =   1785
      Visible         =   0   'False
      Width           =   3600
      Begin VB.TextBox txts 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "RptTime"
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   2730
         MaxLength       =   150
         TabIndex        =   3
         Top             =   15
         Width           =   690
      End
      Begin VB.TextBox txts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "RptTime"
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
         Height          =   225
         Index           =   2
         Left            =   1860
         MaxLength       =   150
         TabIndex        =   2
         Top             =   15
         Width           =   690
      End
      Begin VB.TextBox txts 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "RptTime"
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   780
         MaxLength       =   150
         TabIndex        =   1
         Top             =   15
         Width           =   975
      End
      Begin VB.TextBox txts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "RptTime"
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
         Height          =   225
         Index           =   0
         Left            =   -15
         MaxLength       =   150
         TabIndex        =   0
         Top             =   15
         Width           =   690
      End
   End
   Begin VB.Frame fTop 
      Height          =   645
      Left            =   60
      TabIndex        =   9
      Top             =   -60
      Width           =   10995
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Accured / PrePaid Expenses"
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
         Index           =   7
         Left            =   60
         TabIndex        =   10
         Top             =   165
         Width           =   10125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Accured / PrePaid Expenses"
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
         Index           =   8
         Left            =   75
         TabIndex        =   11
         Top             =   180
         Width           =   10110
      End
   End
   Begin MSComctlLib.ListView LvExps 
      Height          =   2460
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3540
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   4339
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Accured Expense Title"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   1940
      EndProperty
   End
   Begin MSComctlLib.ListView LvSals 
      Height          =   2865
      Left            =   60
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   5054
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Emp. Type"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Employee Name"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Salary"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Absent"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Deduction"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "O.Time Hrs"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "OT. Amount"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Net Salary"
         Object.Width           =   1940
      EndProperty
   End
   Begin MSComctlLib.ListView LvPrePymts 
      Height          =   2460
      Left            =   5580
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3525
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   4339
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PrePaid Expense Title"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   1940
      EndProperty
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   375
      Left            =   7860
      TabIndex        =   7
      Top             =   6240
      Width           =   1425
      Caption         =   "Save     "
      PicturePosition =   327683
      Size            =   "2514;661"
      Picture         =   "frmMonthlyAccured.frx":0000
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   9330
      TabIndex        =   8
      Top             =   6240
      Width           =   1425
      Caption         =   " Cancel       "
      PicturePosition =   327683
      Size            =   "2514;661"
      Picture         =   "frmMonthlyAccured.frx":0112
      Accelerator     =   67
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
      Visible         =   0   'False
      Begin VB.Menu mnuSelect 
         Caption         =   "Select All"
      End
   End
End
Attribute VB_Name = "frmMonthlyAccured"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PermissionDenied As Boolean
Private Sub GetDueSalEmps()
    Dim AccRec As New ADODB.Recordset
    
    Dim itm As ListItem

    With AccRec
    
         .Open "Select * From VEmps Where AcrAccno Not In (Select distinct AccNo From Vouchers Where Month(VDATE)=" & Month(Date) & ") And Active=1 Order By EmpName", Con, adOpenForwardOnly, adLockReadOnly
         
         LvSals.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set itm = LvSals.ListItems.Add(, ![AccNo], ![EmpType] & "")
            
            'Accured Salary AcccNo
            
            itm.Tag = ![AcrAccNo] & ""
            
            'Accured Over Time Account No. in TAG
            
            itm.ListSubItems.Add(, ![OTAcc] & "'", ![EmpName]).Tag = ![AcrOTAcc] & ""
            
            'Over Time Percent Of Real Time Hour Salary (1.5 OR 2.0)
            
            itm.ListSubItems.Add(, , Format(Val(![Salary] & ""), "0.00")).Tag = Val(![OverTime] & "")
            itm.SubItems(3) = "0"
            itm.SubItems(4) = "0.00"
            itm.SubItems(5) = "0"
            itm.SubItems(6) = "0.00"
            itm.SubItems(7) = itm.SubItems(2)
            itm.Checked = True
           .MoveNext
        Next
        
        .Close
   End With
    Set AccRec = Nothing
End Sub

Private Sub GetDueExpenses()
    Dim AccRec As New ADODB.Recordset
    
    Dim itm As ListItem

    With AccRec
    
         .Open "Select ExpAccNo,ExpAccTitle,AcrExpAccNo,Amount,PrePaid From VExpenseAccounts Where AcrExpAccno Not In (Select distinct AccNo From Vouchers Where Month(VDATE)=" & Month(Date) & ")  and Amount >0 Order By ExpAccTitle", Con, adOpenForwardOnly, adLockReadOnly
         
         LvExps.ListItems.Clear
         LvPrePymts.ListItems.Clear
         
        For i = 0 To .RecordCount - 1
            If ![PrePaid] Then
               Set itm = LvPrePymts.ListItems.Add(, ![ExpAccNo], ![ExpAcTitle] & "")
            Else
               Set itm = LvExps.ListItems.Add(, ![ExpAccNo], ![ExpAccTitle] & "")
            End If
            
            itm.Tag = ![AcrExpAccNo] & ""
            itm.SubItems(1) = Format(Val(![Amount] & ""), "0.00")
           
            itm.Checked = True
           
           .MoveNext
           
        Next
        .Close
   End With
    Set AccRec = Nothing
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdSave_Click()


If AnyItemSelected("SAL") Then Call AddAcrSalVoucher
If AnyItemSelected("EXP") Then Call AddAcrExpVoucher
If AnyItemSelected("PRE") > 0 Then Call AddPrepaidExpVoucher

Unload Me
   
End Sub

Private Sub Form_Activate()
'Check For User Access
If PermissionDenied Then
   Me.Hide
   MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
   Unload Me
   Exit Sub
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
End Sub

Private Sub Form_Load()


'Employees Options Included Or Not

LvSals.Visible = EmployeesEnabled
If EmployeesEnabled Then Call GetDueSalEmps

''''''''''''''''''''''''''''''''''''''

   PermissionDenied = Not UserHasAccess("AccuredPrepaidVoucher")
   Call GetDueExpenses
   Me.KeyPreview = True
End Sub



Private Function AddAcrExpVoucher() As Boolean

Dim rsexp As New ADODB.Recordset

Dim VchrNo As String, INextSno As Double, Bal As Double
Dim ExpAmt As Double

      
    VchrNo = getNextVchrNo(Date, "JV")
    INextSno = getNextSno(Date)
Con.BeginTrans
With LvExps.ListItems

For i = 1 To .Count
      
If Not .ITEM(i).Checked Then GoTo EndOfLoop

      ''''''''''''' Debit Entry To Expense Account '''''''''''''
      
         Bal = getBalance(.ITEM(i).key, Date)
         ExpAmt = Val(.ITEM(i).SubItems(1) & "")
         Bal = Bal + ExpAmt
        
         Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Date & "','" & _
         VchrNo & "','" & .ITEM(i).key & "','Expense For " & Format(Date, "MMMM yyyy") & "'," & _
         ExpAmt & ",0," & Bal & ",'',0)"
         
         Call adjustLedger(.ITEM(i).key, Date, INextSno, ExpAmt)
         INextSno = INextSno + 1
         
      ''''''''''''' Credit Entry To Accrued Expense Account '''''''''''''
         
         
         Bal = getBalance(.ITEM(i).Tag, Date)
         ExpAmt = Val(Val(.ITEM(i).SubItems(1) & "") & "")
         Bal = Bal - ExpAmt
        
         Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Date & "','" & _
         VchrNo & "','" & .ITEM(i).Tag & "','Due For " & Format(Date, "MMMM yyyy") & "',0," & _
         ExpAmt & "," & Bal & ",'',0)"
         
         Call adjustLedger(.ITEM(i).Tag, Date, INextSno, -ExpAmt)
         INextSno = INextSno + 1
         

EndOfLoop:
    Next i
      


End With
Con.CommitTrans
AddAcrExpVoucher = True
Exit Function
ERR:
MsgBox "Accured Expenses Are Not Added Due To Error:" & ERR.Number & vbNewLine & ERR.Description
AddAcrExpVoucher = False
Con.RollbackTrans
End Function





Private Function AddPrepaidExpVoucher() As Boolean



Dim VchrNo As String, INextSno As Double, Bal As Double
Dim ExpAmt As Double


With LvPrePymts.ListItems

    VchrNo = getNextVchrNo(Date, "JV")
    INextSno = getNextSno(Date)
Con.BeginTrans

For i = 1 To .Count
      
If Not .ITEM(i).Checked Then GoTo EndOfLoop
      ''''''''''''' Debit Entry To Expense Account '''''''''''''
      
         Bal = getBalance(.ITEM(i).key, Date)
         'ExpAmt = IIf(Val(![Balance] & "") > Val(Val(.ITEM(i).SubItems(1) & "") & ""), Val(Val(.ITEM(i).SubItems(1) & "") & ""), Val(![Balance] & ""))
         Bal = Bal + ExpAmt
        
         Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Date & "','" & _
         VchrNo & "','" & .ITEM(i).key & "','Prepaid Adjustment For " & Format(Date, "MMMM yyyy") & "'," & _
         ExpAmt & ",0," & Bal & ",'',0)"
         
         Call adjustLedger(.ITEM(i).key, Date, INextSno, ExpAmt)
         INextSno = INextSno + 1
         
      ''''''''''''' Credit Entry To Prepaid Expense Account '''''''''''''
         
         
         Bal = getBalance(.ITEM(i).Tag, Date)
         
         Bal = Bal - ExpAmt
        
         Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Date & "','" & _
         VchrNo & "','" & .ITEM(i).Tag & "','Prepaid Adjustment For " & Format(Date, "MMMM yyyy") & "',0," & _
         ExpAmt & "," & Bal & ",'',0)"
         
         Call adjustLedger(.ITEM(i).Tag, Date, INextSno, -ExpAmt)
         INextSno = INextSno + 1
         
         
         
EndOfLoop:
    Next i
      


End With
Con.CommitTrans
AddPrepaidExpVoucher = True
Exit Function
ERR:
MsgBox "Accured Expenses Are Not Added Due To Error:" & ERR.Number & vbNewLine & ERR.Description
AddPrepaidExpVoucher = False
Con.RollbackTrans
End Function



Private Function AddAcrSalVoucher() As Boolean

Dim VchrNo As String, INextSno As Double, Bal As Double
Dim Salary As Double

With LvSals.ListItems
      
    VchrNo = getNextVchrNo(Date, "JV")
    INextSno = getNextSno(Date)
Con.BeginTrans

For i = 1 To .Count

If Not .ITEM(i).Checked Then GoTo EndOfLoop

        ''''''''''''' Debit Entry To Salary Account '''''''''''''
        Bal = getBalance(.ITEM(i).key, Date)
        Salary = Val(Val(.ITEM(i).SubItems(7) & "") & "") - Val(.ITEM(i).SubItems(6) & "")
        Bal = Bal + Salary
        
        Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Date & "','" & _
         VchrNo & "','" & .ITEM(i).key & "','Salary For " & Format(Date, "MMMM yyyy") & "'," & _
         Salary & ",0," & Bal & ",'',0)"
         
         Call adjustLedger(.ITEM(i).key, Date, INextSno, Salary)
         INextSno = INextSno + 1
         
         ''''''''''''' Credit Entry To Accrued Salary Account '''''''''''''
         
         
         Bal = getBalance(.ITEM(i).Tag, Date)
         'Salary Payable Excluding Over time
         Salary = Val(.ITEM(i).SubItems(7) & "") - Val(.ITEM(i).SubItems(6) & "")
         Bal = Bal - Salary
        
         Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Date & "','" & _
         VchrNo & "','" & .ITEM(i).Tag & "','Salary For " & Format(Date, "MMMM yyyy") & "',0," & _
         Salary & "," & Bal & ",'',0)"
         
         Call adjustLedger(.ITEM(i).Tag, Date, INextSno, -Salary)
         INextSno = INextSno + 1
         
         
'''''''''''''''''''''''''''' Accured Ovr Time If Present '''''''''''''''''''''
        ' Check For Over Time is Applicable Or Not
            If .ITEM(i).ListSubItems(1).Tag = "" Or Val(.ITEM(i).SubItems(6) & "") = 0 Then GoTo EndOfLoop
            
         ''''''''''''' Debit Entry To Over Time Account '''''''''''''
        Dim OverTimeAcc As String, OverTime As Double
        OverTimeAcc = Replace(.ITEM(i).ListSubItems(1).key, "'", "")
         
        Bal = getBalance(OverTimeAcc, Date)
        OverTime = Val(.ITEM(i).SubItems(6) & "")
        Bal = Bal + OverTime
        
        Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Date & "','" & _
         VchrNo & "','" & OverTimeAcc & "','Over Time For " & Format(Date, "MMMM yyyy") & "'," & _
         OverTime & ",0," & Bal & ",'',0)"
         
         Call adjustLedger(OverTimeAcc, Date, INextSno, OverTime)
         INextSno = INextSno + 1
         
         ''''''''''''' Credit Entry To Accrued Over Time Account '''''''''''''
         Dim AcrOverTimeAcc As String
         AcrOverTimeAcc = .ITEM(i).ListSubItems(1).Tag
         
         Bal = getBalance(AcrOverTimeAcc, Date)
         Bal = Bal - OverTime
        
         Con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Date & "','" & _
         VchrNo & "','" & AcrOverTimeAcc & "','Over Time For " & Format(Date, "MMMM yyyy") & "',0," & _
         OverTime & "," & Bal & ",'',0)"
         
         Call adjustLedger(AcrOverTimeAcc, Date, INextSno, -OverTimes)
         INextSno = INextSno + 1
         
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

         
EndOfLoop:
    Next i
      


End With
Con.CommitTrans
AddAcrSalVoucher = True
Exit Function
ERR:
MsgBox "Error:" & ERR.Number & vbNewLine & ERR.Description
AddAcrSalVoucher = False
Con.RollbackTrans
End Function


Private Function AnyItemSelected(LVString As String) As Boolean
Dim v As ListView
Select Case LVString
   Case "SAL"
      Set v = LvSals
   Case "EXP"
      Set v = LvExps
   Case "PRE"
      Set v = LvPrePymts
   End Select
   
   For i = 1 To v.ListItems.Count
      If v.ListItems(i).Checked Then Exit For
   Next
   
   If i <= v.ListItems.Count And i <> 0 Then
      AnyItemSelected = True
   Else
      AnyItemSelected = False
   End If
   
   Set v = Nothing
   
End Function

Private Sub Form_Resize()
   On Error Resume Next
   fTop.Move (Me.ScaleWidth - fTop.Width) / 2
   If LvSals.Visible Then
      LvSals.Move 50, fTop.Top + fTop.Height + 50, Me.ScaleWidth - 100, Me.ScaleHeight * 2 / 5
   Else
      LvSals.Move 50, fTop.Top + fTop.Height + 50, Me.ScaleWidth - 100, 10
   End If
   
   LvExps.Move 50, LvSals.Top + LvSals.Height + 50, Me.ScaleWidth / 2 - 75, Me.ScaleHeight - (LvSals.Height + LvSals.Top + 600)
   LvPrePymts.Move LvExps.Left + LvExps.Width + 50, LvExps.Top, LvExps.Width, LvExps.Height
   
   cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 50, Me.ScaleHeight - cmdCancel.Height - 100
   cmdSave.Move cmdCancel.Left - cmdCancel.Width - 50, Me.ScaleHeight - cmdSave.Height - 100
   
End Sub

Private Sub LvSals_DblClick()
Dim TheWidth As Long, HalfWidth As Long
With LvSals
    If .ListItems.Count = 0 Then Exit Sub
    If Not .SelectedItem.Checked Then .SelectedItem.Checked = True
    
    TheWidth = 0
    HalfWidth = 0
    
    For i = 0 To 3
        'For Both Absents And Over Time
        TheWidth = TheWidth + .ColumnHeaders(i + 4).Width
        'For Only Absents
        If i < 2 Then HalfWidth = HalfWidth + .ColumnHeaders(i + 4).Width
        
        txts(i).Move .ColumnHeaders(i + 4).Left - .ColumnHeaders(4).Left, 0, .ColumnHeaders(i + 4).Width - 20
        txts(i) = .SelectedItem.SubItems(i + 3)
    Next
    If .SelectedItem.ListSubItems(1).Tag = "" Then
        PicEdit.Move .Left + .ColumnHeaders(4).Left, .Top + .SelectedItem.Top, HalfWidth
        txts(2).Visible = False
        txts(3).Visible = False
    Else
        PicEdit.Move .Left + .ColumnHeaders(4).Left, .Top + .SelectedItem.Top, TheWidth
        txts(2).Visible = True
        txts(3).Visible = True
    End If
    
    
    PicEdit.Visible = True
    txts(0).SetFocus
End With

End Sub

Private Sub LvSals_GotFocus()
PicEdit.Visible = False
End Sub

Private Sub txts_Change(Index As Integer)
Dim DaySalary As Double, OverTimePerHour As Double

With LvSals.SelectedItem
    DaySalary = Format(Val(.SubItems(2)) / 30, "0.00")
    
    OverTimePerHour = Format(DaySalary / 8 * Val(.ListSubItems(2).Tag), "0.00")
    
    Select Case Index
        Case 0
               txts(1) = Val(txts(0)) * DaySalary
        Case 2
                txts(3) = Val(txts(2)) * OverTimePerHour
    End Select
End With
End Sub



Private Sub txts_GotFocus(Index As Integer)
    txts(Index).SelStart = 0
    txts(Index).SelLength = Len(txts(Index))
End Sub

Private Sub txts_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 And Not txts(2).Visible Then
        GoTo Save
    ElseIf Index = 2 Then
        GoTo Save
    End If
Else
    KeyAscii = OnlyNumber(KeyAscii)
End If
    
Exit Sub

Save:
    
    If txts(0) >= 30 Then
        MsgBox "Invalid No. of Absents", vbInformation
        Cancel = True
        Exit Sub
    End If
    
    
    With LvSals.SelectedItem
                
        .SubItems(3) = txts(0)
        .SubItems(4) = txts(1)
        .SubItems(5) = txts(2)
        .SubItems(6) = txts(3)
        'Net Payable Equals to Salary + OverTime - Absent Deduction
        .SubItems(7) = Val(.SubItems(2)) - Val(.SubItems(4)) + Val(.SubItems(6))
        PicEdit.Visible = False
    End With
    
    
    
End Sub

