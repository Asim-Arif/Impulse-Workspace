VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#2.0#0"; "ComboList.ocx"
Begin VB.Form NewAdvSalary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance Sheet..."
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleMode       =   0  'User
   ScaleWidth      =   10345.49
   WindowState     =   2  'Maximized
   Begin ComboList.UserControl1 cmbDept 
      Height          =   330
      Left            =   570
      TabIndex        =   2
      Top             =   495
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
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
      Height          =   6045
      Left            =   15
      TabIndex        =   0
      Top             =   795
      Width           =   11385
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1545
         ScaleHeight     =   285
         ScaleWidth      =   4800
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   4830
         Begin VB.TextBox txtDesc 
            Height          =   300
            Left            =   1200
            TabIndex        =   11
            Top             =   0
            Width           =   3600
         End
         Begin VB.TextBox txtAmt 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   1200
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
         Height          =   5265
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   9287
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "F Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Description"
            Object.Width           =   6350
         EndProperty
      End
      Begin MSForms.CommandButton cmdOk 
         Height          =   405
         Left            =   7950
         TabIndex        =   8
         Top             =   5535
         Width           =   1620
         Caption         =   " Save & Close"
         PicturePosition =   327683
         Size            =   "2857;714"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdNext 
         Height          =   405
         Left            =   6270
         TabIndex        =   7
         Top             =   5535
         Width           =   1620
         Caption         =   "Save & New"
         PicturePosition =   327683
         Size            =   "2857;714"
         Picture         =   "NewAdvSalary.frx":0000
         Accelerator     =   78
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
         Left            =   9630
         TabIndex        =   6
         Top             =   5535
         Width           =   1620
         Caption         =   "Cancel        "
         PicturePosition =   327683
         Size            =   "2857;714"
         Picture         =   "NewAdvSalary.frx":0112
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
      Left            =   0
      Top             =   0
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
            Picture         =   "NewAdvSalary.frx":0224
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NewAdvSalary.frx":0676
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   2850
      TabIndex        =   5
      Top             =   525
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   55115779
      CurrentDate     =   37384
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
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
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      Top             =   315
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Department"
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
      Height          =   195
      Index           =   0
      Left            =   645
      TabIndex        =   3
      Top             =   285
      Width           =   1035
   End
End
Attribute VB_Name = "NewAdvSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbDept_matched()
    Call FillList
End Sub

Private Function Saved() As Boolean
    On Error Resume Next
    Dim Sql As String
    
    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            If .SubItems(3) <> "" Then
                Sql = "Insert into Advances (EmpID,DT,Type," & _
                 "Description,Amount,DAmount,AmountCleared) " & _
                 "Values('" & .Text & "',#" & DTPicker1 & _
                 "#,0,'" & .SubItems(4) & "'," & .SubItems(3) & _
                 "," & .SubItems(3) & ",0)"
                Con.Execute Sql, a
                MsgBox a
            End If
        End With
    Next i
    Saved = True
    Sql = ""
    Exit Function
err:
    MsgBox err.Description
End Function




Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
        Unload Me
        NewAdvSalary.Show
    End If
End Sub

Private Sub cmdOk_Click()
    If Saved Then Unload Me
End Sub

Private Sub Form_Load()

    DtTo = Date
    DtFrom = DateAdd("M", -1, Date)
    cmbDept.ListHeight = 1600
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals Con, "Name", "Departments", "DeptID"
    
    'Set LV.SmallIcons = ImageList1
    
End Sub
Private Sub FillList()
      
    If Not cmbDept.MatchFound Then Exit Sub
    
    Dim Sql As String
    Dim LItem As ListItem
    Dim rs As New ADODB.Recordset
    Dim theKey As String
    Dim IfBold As Boolean
    Dim theName As String
    
    If cmbDept.ID = "0" Then
        Sql = "Select * from VEmp"
        theName = "Employees.Name"
    Else
        Sql = "Select * from Employees Where DeptID='" & cmbDept.ID & "'"
        theName = "Name"
    End If
    LV.ListItems.Clear
    With rs
        .Open Sql, Con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            If IsNull(![EmpID]) Then
                theKey = ![DeptID] & ""
                IfBold = True
            Else
                theKey = ![EmpID] & ""
                IfBold = False
            End If
            
            Set LItem = LV.ListItems.Add(, theKey, theKey)
            LItem.Bold = IfBold
            LItem.SubItems(1) = .Fields(theName) & ""
            LItem.SubItems(2) = ![FName] & ""
            LItem.SubItems(3) = ""
            LItem.SubItems(4) = ""
            'LItem.SubItems(3) = ![] & ""
            .MoveNext
        Loop
        .Close
    End With
    Sql = ""
    theKey = ""
    theName = ""
    
    Set rs = Nothing
    
End Sub



Private Sub Form_Resize()

On Error Resume Next
Dim TheHeight As Long


 FAcc.Move (Me.ScaleWidth - FAcc.Width) / 2
 FLV.Move (Me.ScaleWidth - FLV.Width) / 2
'cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 100
'cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top

'FDrag.Move 50, Me.ScaleHeight * 3 / 5, Me.ScaleWidth - 100
 
 'If LVItems.Visible Then
 '  TheHeight = FDrag.Top
 '  FDrag.Visible = True
 'Else
 '  TheHeight = Me.ScaleHeight - cmdClose.Height - 200
 '  FDrag.Visible = False
 'End If
 
 'LV.Move 50, FAcc.Top + FAcc.Height, Me.ScaleWidth - 100, TheHeight - (FAcc.Top + FAcc.Height)
 'LVItems.Move 50, FDrag.Top + FDrag.Height, Me.ScaleWidth - 100, Me.ScaleHeight - (FAcc.Height + cmdClose.Height + LV.Height + 250)
 'FDrag.ZOrder 1
 
End Sub
Private Sub LV_DblClick()
    With LV.SelectedItem
        If .SubItems(1) = "" Then
            Pic.Visible = False
            Exit Sub
        End If
        Pic.Move LV.ColumnHeaders(4).Left + LV.Left, LV.Top + LV.SelectedItem.Top
        txtAmt.Text = .SubItems(3)
        txtDesc.Text = IIf(txtDesc = "", "Short Term Advance", .SubItems(4))
        txtAmt.SelStart = 0
        txtAmt.SelLength = Len(txtAmt)
        Pic.Visible = True
        txtAmt.SetFocus
    End With
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'LV.SelectedItem.SubItems(3) = txtAmt
        txtDesc.SetFocus
    ElseIf KeyAscii = 27 Then
        txtAmt = ""
        Pic.Visible = False
    End If
    
End Sub
Private Sub txtDesc_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(3) = txtAmt
        LV.SelectedItem.SubItems(4) = txtDesc
        txtAmt = ""
        txtDesc = ""
        Pic.Visible = False
    ElseIf KeyAscii = 27 Then
        txtAmt = ""
        txtDesc = ""
        Pic.Visible = False
    End If
    
End Sub
