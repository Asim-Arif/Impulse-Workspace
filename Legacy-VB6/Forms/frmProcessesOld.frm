VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmProcessesOld 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Processes"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7620
   ScaleMode       =   0  'User
   ScaleWidth      =   6758.159
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr1 
      Left            =   1770
      Top             =   7065
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
   Begin VB.TextBox Text1 
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
      Left            =   60
      TabIndex        =   1
      Top             =   885
      Width           =   3225
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   330
      Top             =   6900
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
            Picture         =   "frmProcessesOld.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcessesOld.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FLV 
      Height          =   6390
      Left            =   60
      TabIndex        =   4
      Top             =   1185
      Width           =   7440
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   105
         ScaleHeight     =   255
         ScaleWidth      =   5295
         TabIndex        =   7
         Top             =   555
         Visible         =   0   'False
         Width           =   5325
         Begin MSForms.ComboBox cmbAuthReqE 
            Height          =   300
            Left            =   3840
            TabIndex        =   18
            Top             =   -30
            Width           =   1455
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "2566;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbOperationE 
            Height          =   300
            Left            =   1830
            TabIndex        =   15
            Top             =   -30
            Width           =   2055
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "3625;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   178
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSupervisor 
            Height          =   300
            Left            =   1215
            TabIndex        =   10
            Top             =   -15
            Width           =   645
            VariousPropertyBits=   679495707
            Size            =   "1138;529"
            SpecialEffect   =   3
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtDescription 
            Height          =   300
            Left            =   600
            TabIndex        =   9
            Top             =   -15
            Width           =   645
            VariousPropertyBits=   679495707
            Size            =   "1138;529"
            SpecialEffect   =   3
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSNo 
            Height          =   300
            Left            =   -15
            TabIndex        =   8
            Top             =   -15
            Width           =   645
            VariousPropertyBits=   679495707
            Size            =   "1138;529"
            SpecialEffect   =   3
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5670
         Left            =   90
         TabIndex        =   11
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
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Component Description"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Supervisor"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Operation"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Auth. Req."
            Object.Width           =   2293
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   2985
         TabIndex        =   6
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
         TabIndex        =   3
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
      Left            =   60
      TabIndex        =   0
      Top             =   375
      Width           =   6780
   End
   Begin MSForms.ComboBox cmbAuthReq 
      Height          =   300
      Left            =   5385
      TabIndex        =   17
      Top             =   885
      Width           =   1451
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2566;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Authorization Req."
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
      Index           =   3
      Left            =   5415
      TabIndex        =   16
      Top             =   690
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Operation"
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
      Index           =   2
      Left            =   3330
      TabIndex        =   14
      Top             =   675
      Width           =   720
   End
   Begin MSForms.ComboBox cmbOperation 
      Height          =   300
      Left            =   3315
      TabIndex        =   13
      Top             =   885
      Width           =   2055
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3625;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Supervisor"
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
      Left            =   90
      TabIndex        =   12
      Top             =   675
      Width           =   765
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
      Left            =   90
      TabIndex        =   5
      Top             =   150
      Width           =   1395
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   825
      Left            =   6885
      TabIndex        =   2
      Top             =   360
      Width           =   615
      Caption         =   " Add"
      PicturePosition =   65543
      Size            =   "1085;1455"
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
Attribute VB_Name = "frmProcessesOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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

Private Sub cmdAdd_Click()

    On Error GoTo err
    If txtDesc = "" Then
        MsgBox "Please Enter Process Name", vbInformation
        Exit Sub
    ElseIf cmbOperation.MatchFound = False Then
        MsgBox "Please Select Operation.", vbInformation
        Exit Sub
    ElseIf cmbAuthReq.MatchFound = False Then
        MsgBox "Please Select Authorization.", vbInformation
        Exit Sub
    End If
    
    LV.Sorted = False
    
    Dim Affect As Long
    
    Dim MaxSNo As Integer
    MaxSNo = Val(con.Execute("SELECT Max(SNo) FROM Processes").Fields(0) & "")
    MaxSNo = MaxSNo + 1
    
    con.Execute "INSERT INTO Processes (SNo,Description,Supervisor,Operation,AuthRequired) VALUES(" & _
     MaxSNo & ",'" & txtDesc & "','" & Text1 & "'," & cmbOperation.ListIndex & _
     "," & cmbAuthReq.ListIndex & ")", Affect
     
    If Affect = 1 Then
        Call FillList
        txtDesc = ""
    End If
    
    LV.Sorted = True
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
    
    SQL = "SELECT * FROM Processes ORDER BY Sno"

    With rsList
        .CursorLocation = adUseClient
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.add(, ![ProcessID] & "'", ![Sno])
            ITM.SubItems(1) = ![Description] & ""
            ITM.SubItems(2) = ![Supervisor] & ""
            strOperation = ""
            If Val(!Operation & "") = 0 Then
                strOperation = "Employee"
            ElseIf Val(!Operation & "") = 1 Then
                strOperation = "Contractor"
            ElseIf Val(!Operation & "") = 2 Then
                strOperation = "Both"
            End If
            ITM.SubItems(3) = strOperation
            If IsNull(!AuthRequired) Then
                ITM.SubItems(4) = "No"
            Else
                ITM.SubItems(4) = cmbAuthReq.List(Abs(!AuthRequired))
            End If
            .MoveNEXT
        Next
        .Close
    End With
    Set rsList = Nothing
    SQL = ""
    
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
    With cmbOperation
        .AddItem "Employee"
        .AddItem "Contractor"
        .AddItem "Both"
    End With
    With cmbOperationE
        .AddItem "Employee"
        .AddItem "Contractor"
        .AddItem "Both"
    End With
    With cmbAuthReq
        .AddItem "No"
        .AddItem "Yes"
    End With
    With cmbAuthReqE
        .AddItem "No"
        .AddItem "Yes"
    End With
    Call FillList
End Sub

Private Sub Form_Resize()
    Me.Top = 1700 'MainForm.Toolbar1.Top + MainForm.Toolbar1.Height + 100
End Sub

Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Exit Sub
    With LV
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
                ColumnHeader.Icon = 2
            Else
                .SortOrder = lvwAscending
                ColumnHeader.Icon = 1
            End If
        Else
            LV.ColumnHeaderIcons = Nothing
            Set LV.ColumnHeaderIcons = ImageList1
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
            ColumnHeader.Icon = 1
        End If
        .Sorted = True
    End With
End Sub

Private Sub LV_DblClick()

    'Edit This Record...
    If LV.ListItems.count = 0 Then Exit Sub
    '-------------------------------------------------------
    PicEdit.Move LV.Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(1).Width + LV.ColumnHeaders(2).Width + LV.ColumnHeaders(3).Width + LV.ColumnHeaders(4).Width + LV.ColumnHeaders(5).Width
    
    
    txtSNo.Move 0, 0, LV.ColumnHeaders(1).Width
    txtSNo.Text = LV.SelectedItem
    
    txtDescription.Move txtSNo.Width - 10, 0, LV.ColumnHeaders(2).Width
    txtDescription.Text = LV.SelectedItem.SubItems(1)
    
    txtSupervisor.Move txtDescription.Left + txtDescription.Width - 20, 0, LV.ColumnHeaders(3).Width
    txtSupervisor.Text = LV.SelectedItem.SubItems(2)
    
    cmbOperationE.Move txtSupervisor.Left + txtSupervisor.Width - 20, 0, LV.ColumnHeaders(4).Width
    cmbOperationE.Text = LV.SelectedItem.SubItems(3)
    
    cmbAuthReqE.Move txtSupervisor.Left + txtSupervisor.Width + cmbOperationE.Width - 20, 0, LV.ColumnHeaders(5).Width
    cmbAuthReqE.ListIndex = IIf(LV.SelectedItem.SubItems(4) = "Yes", 1, 0)
    
    PicEdit.Visible = True
    
    With txtSNo
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PicEdit.Visible = False
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
