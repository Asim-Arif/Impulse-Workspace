VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmOtherBenefits 
   BorderStyle     =   0  'None
   ClientHeight    =   7665
   ClientLeft      =   15
   ClientTop       =   210
   ClientWidth     =   12690
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
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleMode       =   0  'User
   ScaleWidth      =   11366.61
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   45
      TabIndex        =   10
      Top             =   6315
      Visible         =   0   'False
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   609
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
      CustomFormat    =   "dddd dd-MMM-yyyy HH:mm:ss tt"
      Format          =   71041027
      CurrentDate     =   37708
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   2580
      TabIndex        =   9
      Top             =   390
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   285
      Left            =   5490
      TabIndex        =   3
      Top             =   390
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   503
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
      CustomFormat    =   "dddd dd-MMM-yyyy"
      Format          =   71041027
      CurrentDate     =   37384
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   390
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.Frame FLV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   45
      TabIndex        =   0
      Top             =   765
      Width           =   12540
      Begin VB.PictureBox Pic 
         Height          =   285
         Left            =   2085
         ScaleHeight     =   225
         ScaleWidth      =   6480
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   6540
         Begin VB.TextBox txtDesc 
            Height          =   300
            Left            =   1500
            TabIndex        =   7
            Top             =   -30
            Width           =   5010
         End
         Begin VB.TextBox txtAmt 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   -15
            TabIndex        =   6
            Top             =   -30
            Width           =   1500
         End
      End
      Begin Crystal.CrystalReport cr1 
         Left            =   90
         Top             =   420
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
      Begin MSComctlLib.ListView LV 
         Height          =   6105
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   12390
         _ExtentX        =   21855
         _ExtentY        =   10769
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Amount"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   8811
         EndProperty
      End
      Begin MSForms.CommandButton cmdOpt 
         Height          =   405
         Left            =   8985
         TabIndex        =   11
         Top             =   6375
         Width           =   1680
         Caption         =   "Options         "
         PicturePosition =   327683
         Size            =   "2963;714"
         Accelerator     =   79
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCancel 
         Height          =   405
         Left            =   10725
         TabIndex        =   4
         Top             =   6375
         Width           =   1620
         Caption         =   "Close        "
         PicturePosition =   327683
         Size            =   "2857;714"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2070
      Top             =   11790
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherBenefits.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherBenefits.frx":005E
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   60
      TabIndex        =   8
      Top             =   150
      Width           =   8565
      BackColor       =   11517387
      VariousPropertyBits=   8388627
      Caption         =   "  Department                               Employee :                                              Date :"
      Size            =   "15108;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Begin VB.Menu mnuNew 
         Caption         =   "New Entry"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update This Entry"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete This Entry"
      End
      Begin VB.Menu lndelete 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Benefits Sheet"
      End
   End
End
Attribute VB_Name = "frmOtherBenefits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit As Boolean

Private Function Saved() As Boolean
'    On Error Resume Next
    Dim StartSno As Double, LastSno As Double
    Dim SQL As String
    StartSno = GetMax("EntryID", "Advances")
    
    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            If .SubItems(3) <> "" Then
                SQL = "ISERT INTO EmpFine (EmpID,FineDate,Type," & _
                 "Amount, Description) Values('" & .Text & "',#" & DT & _
                 "#,0,'" & .SubItems(4) & "'," & .SubItems(3) & _
                 "," & .SubItems(3) & ",0)"
                con.Execute SQL, a
            End If
        End With
    Next i
    
    LastSno = GetMax("EntryID", "Advances")
    
    
    
   SQL = ""
   Saved = True
      
    Exit Function
err:
    MsgBox err.Description
End Function




Private Sub cmbdept_matched()
cmbEmp.ClearVals

If cmbDept.ID = 0 Then
    cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where Active=1"
Else
    cmbEmp.AddVals con, "Name", "Employees", "EmpID", "WHERE DeptID='" & cmbDept.ID & "' and Active=1"
End If
End Sub

Private Sub cmbEmp_matched()
    Call FillList
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
        Unload Me
        NewAdvShort.Show
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub cmdOpt_Click()

PopUpMenu mnuPopup

End Sub

Private Sub DT_Change()
FillList
End Sub

Private Sub Form_Load()
    
    mnuPopup.Visible = False
    
    DT = Date
    
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1600
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where Active=1"
    cmbDept.Text = "<All Departments>"
    
End Sub
Private Sub FillList()
      
    If Not cmbEmp.MatchFound Then Exit Sub
    
    Dim SQL As String, LItem As ListItem
    Dim rs As New ADODB.Recordset
    
    SQL = "Select * from EmpFine Where EmpID='" & cmbEmp.ID & "' AND MONTH(FineDate)=" & DT.Month _
        & " AND YEAR(FineDate)=" & DT.year & " AND Type='Other'"
    theName = "Name"
    
    LV.ListItems.Clear
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            If IsNull(![EmpId]) Then
                theKey = ![DeptID] & ""
                IfBold = True
            Else
                theKey = ![EmpId] & ""
                IfBold = False
            End If
            
            Set LItem = LV.ListItems.add(, , Format(!FineDate, "ddd dd-mmm-yyyy"))
            LItem.Tag = !FineID
            LItem.ListSubItems.add , , !Amount
            LItem.ListSubItems.add , , !Description
            
            .MoveNext
        Loop
        .Close
    End With
    SQL = ""
    theKey = ""
    theName = ""
    
    Set rs = Nothing
    
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.Count = 0 Then Exit Sub
    With LV.SelectedItem
        If .SubItems(1) = "" Then
            Pic.Visible = False
            edit = False
            Exit Sub
        End If
        
        Pic.Move LV.ColumnHeaders(2).Left + LV.Left, LV.Top + LV.SelectedItem.Top, LV.Width - LV.ColumnHeaders(1).Width ', .Height
        
        txtAmt = .SubItems(1)
        txtDesc = IIf(.SubItems(2) <> "", .SubItems(2), "")
        txtAmt.SelStart = 0
        txtAmt.SelLength = Len(txtAmt)
        Pic.Visible = True
        edit = True
        txtAmt.SetFocus
    End With
    
End Sub

Private Sub LV_GotFocus()
Pic.Visible = False
edit = False
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    LV_DblClick
ElseIf KeyCode = 46 Then
    mnuDelete_Click
End If

End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbRightButton Then
    PopUpMenu mnuPopup
End If

End Sub

Private Sub mnuDelete_Click()
Dim Mess
    If LV.ListItems.Count > 0 Then
        Mess = MsgBox("Are You Sure You Want To Delete This Entry !", vbInformation + vbYesNo)
    If Mess = vbYes Then
        con.Execute "DELETE FROM EmpFine WHERE FineID=" & LV.SelectedItem.Tag & ""
        Call FillList
    End If
End If

End Sub

Private Sub mnuNew_Click()

If LV.ListItems.Count > 0 Then

    Pic.Move LV.ColumnHeaders(2).Left + LV.Left, LV.Top + LV.ListItems(LV.ListItems.Count).Top + 200, LV.Width - LV.ColumnHeaders(1).Width
    txtAmt = ""
    txtDesc = ""
Else
    Pic.Move LV.ColumnHeaders(2).Left + LV.Left, LV.Top + 300, LV.Width - LV.ColumnHeaders(1).Width
    
End If

Pic.Visible = True
txtAmt.SetFocus

End Sub

Private Sub mnuPrint_Click()

    Cond = "{VSalaryLedger.EmpID}='" & cmbEmp.ID & "'"
    
'With cr1
'    .ReportFileName = RptPath & "\EmpOtherBenefits.Rpt"
'    .Formulas(0) = "FineMonth=" & dt.Month & ""
'    .Formulas(1) = "Company='" & company & "'"
'    .Connect = ConnectStr
'    .SelectionFormula = Cond
'    .Action = 1
'    .PageZoomNext
'End With

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpOtherBenefits.rpt")
Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    End If
Next

Me.Hide
f.ShowReport Cond & " ", rpt


Screen.MousePointer = vbDefault

End Sub

Private Sub mnuUpdate_Click()
LV_DblClick
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    
    If KeyAscii = 13 Then
        txtDesc.SetFocus
        Exit Sub
    ElseIf KeyAscii = 27 Then
        txtAmt = ""
        Pic.Visible = False
        edit = False
        Exit Sub
    End If
    
    KeyAscii = OnlyNumber(KeyAscii)

    Exit Sub
err:
    MsgBox err.Description
    Screen.MousePointer = vbDefault
End Sub
Private Sub txtDesc_KeyPress(KeyAscii As Integer)
Dim ITM As ListItem
With LV.SelectedItem
    If Val(txtAmt) = 0 Then
        MsgBox "Amount Could Not Be Zero !", vbInformation
        txtAmt.SetFocus
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        
        If Not cmbEmp.MatchFound Then
            MsgBox "Please Select The Employee !", vbInformation
            Exit Sub
        End If
        
        If LV.ListItems.Count = 0 Then
            Set ITM = LV.ListItems.add(, , " ")
            ITM.ListSubItems.add , , " "
            ITM.ListSubItems.add , , " "
        End If
        
        LV.SelectedItem.SubItems(1) = Format(txtAmt, "0.00")
        LV.SelectedItem.SubItems(2) = txtDesc
        If edit = True Then
            con.Execute "UPDATE EmpFine SET Amount=" & txtAmt & ", Description='" & txtDesc & "' WHERE FineID=" & .Tag & ""
            edit = False
        Else
            con.Execute "INSERT INTO EmpFine(EmpID,FineDate,Type,Amount,Description) VALUES('" & _
                        cmbEmp.ID & "','" & DT.value & "','Other'," & Val(txtAmt) & ",'" & txtDesc & "')"
        End If
            Call FillList
        txtAmt = ""
        txtDesc = ""
        Pic.Visible = False
        edit = False
        
    ElseIf KeyAscii = 27 Then
        txtAmt = ""
        txtDesc = ""
        Pic.Visible = False
        edit = False
    End If
End With
End Sub
