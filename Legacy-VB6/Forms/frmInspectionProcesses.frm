VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmInspectionProcesses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inspection Processes"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7770
   ScaleMode       =   0  'User
   ScaleWidth      =   6758.159
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   0
      Top             =   375
      Width           =   1200
   End
   Begin VB.TextBox txtUrduFields 
      Alignment       =   1  'Right Justify
      DataSource      =   "datPrimaryRS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   915
      Width           =   6780
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   1845
      Top             =   6705
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   705
      Top             =   6570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInspectionProcesses.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInspectionProcesses.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FLV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6390
      Left            =   60
      TabIndex        =   5
      Top             =   1320
      Width           =   7440
      Begin VB.TextBox txtUrduFields 
         Alignment       =   1  'Right Justify
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   300
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2430
         Visible         =   0   'False
         Width           =   6780
      End
      Begin VB.CheckBox chkShowUrdu 
         Caption         =   "Show Urdu Description"
         Height          =   270
         Left            =   90
         TabIndex        =   12
         Top             =   5940
         Width           =   2430
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   495
         TabIndex        =   11
         Top             =   705
         Visible         =   0   'False
         Width           =   5535
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5670
         Left            =   90
         TabIndex        =   8
         Top             =   210
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   10001
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   15199215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Process"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   2985
         TabIndex        =   7
         Top             =   5940
         Visible         =   0   'False
         Width           =   1425
         ForeColor       =   0
         Caption         =   "Print         "
         PicturePosition =   327683
         Size            =   "2514;635"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   360
         Left            =   5955
         TabIndex        =   4
         Top             =   5940
         Width           =   1425
         ForeColor       =   0
         Caption         =   " Close     "
         PicturePosition =   327683
         Size            =   "2514;635"
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
   Begin VB.TextBox txtDesc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1335
      TabIndex        =   1
      Top             =   375
      Width           =   5535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code"
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
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   10
      Top             =   150
      Width           =   375
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÑÇÓíÓ ˜Ç äÇã "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Index           =   19
      Left            =   5685
      TabIndex        =   9
      Top             =   645
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Process Description"
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
      Height          =   195
      Index           =   1
      Left            =   1335
      TabIndex        =   6
      Top             =   150
      Width           =   1395
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   960
      Left            =   6885
      TabIndex        =   3
      Top             =   360
      Width           =   615
      Caption         =   " Add"
      PicturePosition =   65543
      Size            =   "1085;1693"
      Accelerator     =   65
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmInspectionProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iColNo As Integer

Private Sub cmbAuthReqE_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode.Value = 13 Then
    
        con.Execute "UPDATE Processes Set SNo=" & txtSNo & ",Description='" & txtDescription & _
         "',Supervisor='" & txtSupervisor & "',Operation=" & cmbOperationE.ListIndex & ",AuthRequired=" & _
         cmbAuthReqE.ListIndex & " Where ProcessID=" & Val(LV.SelectedItem.key), iAffect
         
        If iAffect = 1 Then
            PicEdit.Visible = False
            Call FillList
        End If
        
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub cmbOperationE_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim iAffect As Integer
    If KeyCode.Value = 13 Then
        cmbAuthReqE.SetFocus
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
End Sub

Private Sub chkShowUrdu_Click()

    Dim i As Integer
    For i = 1 To LV.ListItems.count
        If chkShowUrdu.Value = vbChecked Then
            LV.ListItems(i).ListSubItems(2) = LV.ListItems(i).ListSubItems(1).Tag
        Else
            LV.ListItems(i).ListSubItems(2) = LV.ListItems(i).ListSubItems(2).Tag
        End If
    Next
    If chkShowUrdu.Value = vbChecked Then
        LV.ColumnHeaders(3).Text = "ÑÇÓíÓ ˜Ç äÇã "
    Else
        LV.ColumnHeaders(3).Text = "Process"
    End If
    
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    If txtCode = "" Then
        MsgBox "Please Enter Code.", vbInformation
        Exit Sub
    ElseIf txtDesc = "" Then
        MsgBox "Please Enter Process Name", vbInformation
        Exit Sub
    End If
    
    LV.Sorted = False
    
    Dim Affect As Long
   
    con.Execute "INSERT INTO InspectionProcesses(Code,ProcessName,ProcessNameUrdu) VALUES('" & _
     txtCode & "','" & txtDesc & "',N'" & txtUrduFields(0) & "')", Affect
    
    If Affect = 1 Then
        Call FillList
        txtDesc = ""
    End If
    
    LV.Sorted = True
    txtCode.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

On Error GoTo err
Screen.MousePointer = vbHourglass
Dim DateFrom As Date
Dim DateTo As Date
DateFrom = DateSerial(cmbYear, 1, 1)
DateTo = DateSerial(cmbYear, 12, 31)
With cr1
    .ReportFileName = RptPRPath & "\Holidays.rpt"
    .DataFiles(0) = DatabasePath
    
    .SelectionFormula = "Year({Holidays.DT})=" & cmbYear
      
    .Formulas(0) = "ForYear='Year-" & cmbYear & "'"
    
    .Action = 1
    .PageZoomNext
End With

Screen.MousePointer = vbDefault
Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillList()

    
    Dim rsList As New ADODB.Recordset
    Dim ITM As ListItem
    Dim SQL As String
    Dim strOperation As String
    
    SQL = "SELECT * FROM InspectionProcesses ORDER BY EntryID"

    With rsList
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.add(, ![EntryID] & "'", LV.ListItems.count + 1)
            ITM.ListSubItems.add(, , !code & "").Tag = !ProcessNameUrdu & ""
            ITM.ListSubItems.add(, , !ProcessName & "").Tag = !ProcessName & ""
           
            .MoveNEXT
        Next
        .Close
    End With
    Set rsList = Nothing
    SQL = ""
    
    Call chkShowUrdu_Click
End Sub

Private Sub txtDedE_KeyPress(KeyAscii As Integer)
    On Error GoTo err
    Dim theKey As String
    Dim Affect As Long
    theKey = LV.SelectedItem.key
    theKey = Right(theKey, Len(theKey) - 1)
    
    If KeyAscii = 13 Then
        con.Execute "Update TaxRanges Set SalFrom=" & _
         txtFromE & ",SalTo=" & txtToE & ",PerDeduct=" & _
         txtPerE & ",Reduction=" & txtDedE & _
         " Where SalFrom=" & theKey, Affect
        If Affect = 1 Then
            With LV.SelectedItem
                .key = "'" & txtFromE
                .Text = txtFromE
                .SubItems(1) = txtToE
                .SubItems(2) = txtPerE
                .SubItems(3) = txtDedE
            End With
        End If
        Pic.Visible = False
    ElseIf KeyAscii = 27 Then
        txtFromE = ""
        txtToE = ""
        txtPerE = ""
        txtDedE = ""
        Pic.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub Form_Load()
    
    Call FillList
End Sub

Private Sub Form_Resize()
    Me.Top = 1700 'MainForm.Toolbar1.Top + MainForm.Toolbar1.Height + 100
End Sub

Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
End Sub


Private Sub LV_DblClick()

    If iColNo = 0 Then Exit Sub
    
    Dim bShowUrdu As Boolean
    
    If iColNo = 3 And chkShowUrdu.Value = vbChecked Then
        With txtUrduFields(1)
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
            .Text = LV.SelectedItem.SubItems(iColNo - 1)
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    Else
        With txtEdit
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
            .Text = LV.SelectedItem.SubItems(iColNo - 1)
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    End If
     
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    PicEdit.Visible = False
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'For Now
    Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    On Error Resume Next
    If KeyCode = 46 Then
        If MsgBox("Do You Want To Delete This Component?", vbQuestion + vbYesNo) = vbYes Then
            Dim Affect As Integer
            con.Execute "Delete from Components Where CompID=" & Val(LV.SelectedItem.key), Affect
            If Affect = 1 Then
                LV.ListItems.Remove (LV.SelectedItem.Index)
            End If
        End If
    End If
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub txtDescription_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    
    If KeyCode.Value = 13 Then
        With txtSupervisor
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
    
End Sub

Private Sub txtSNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    If KeyCode.Value = 13 Then
        With txtDescription
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtSupervisor_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    Dim iAffect As Integer
    If KeyCode.Value = 13 Then
        cmbOperationE.SetFocus
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > LV.ColumnHeaders(i).Left And x < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtEdit = "" Then
            MsgBox "Invalid Data.", vbInformation
            Exit Sub
        End If
        If iColNo = 2 Then
            con.Execute "UPDATE InspectionProcesses SET Code='" & txtEdit.Text & "' WHERE EntryID=" & Val(LV.SelectedItem.key)
        ElseIf iColNo = 3 Then
            con.Execute "UPDATE InspectionProcesses SET ProcessName='" & txtEdit.Text & "' WHERE EntryID=" & Val(LV.SelectedItem.key)
        End If
        
        FillList
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub txtUrduFields_GotFocus(Index As Integer)
    Call SetUrduLanguage
End Sub

Private Sub txtUrduFields_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index = 1 Then
        If KeyAscii = 13 Then
            If txtUrduFields(1) = "" Then
                MsgBox "Invalid Data.", vbInformation
                Exit Sub
            End If
            con.Execute "UPDATE InspectionProcesses SET ProcessNameUrdu=N'" & txtUrduFields(1).Text & "' WHERE EntryID=" & Val(LV.SelectedItem.key)
            Call FillList
            txtUrduFields(1).Visible = False
        ElseIf KeyAscii = 27 Then
            txtUrduFields(1).Visible = False
        End If

    End If
    
End Sub

Private Sub txtUrduFields_LostFocus(Index As Integer)
    Call SetEnglishLanguage
    If Index = 1 Then
        txtUrduFields(1).Visible = False
    End If
End Sub
