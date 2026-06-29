VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmFDAProcesses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FDA Processes"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   LinkTopic       =   "Form19"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9120
   ScaleMode       =   0  'User
   ScaleWidth      =   8007.679
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMaxDays 
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
      Left            =   2040
      TabIndex        =   15
      Top             =   900
      Width           =   1890
   End
   Begin VB.TextBox txtMinDays 
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
      Left            =   120
      TabIndex        =   14
      Top             =   900
      Width           =   1890
   End
   Begin VB.TextBox txtMinWaste 
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
      Left            =   3960
      TabIndex        =   16
      Top             =   900
      Width           =   1890
   End
   Begin VB.TextBox txtMaxWaste 
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
      Left            =   5880
      TabIndex        =   17
      Top             =   900
      Width           =   1890
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
      Left            =   4860
      TabIndex        =   13
      Top             =   270
      Width           =   2925
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
            Picture         =   "frmFDAProcesses.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFDAProcesses.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FLV 
      Height          =   7800
      Left            =   60
      TabIndex        =   0
      Top             =   1275
      Width           =   8835
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   105
         ScaleHeight     =   255
         ScaleWidth      =   5670
         TabIndex        =   4
         Top             =   645
         Visible         =   0   'False
         Width           =   5700
         Begin MSForms.TextBox txtEMaxDays 
            Height          =   300
            Left            =   2445
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
         Begin MSForms.TextBox txtEMaxWaste 
            Height          =   300
            Left            =   3690
            TabIndex        =   11
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
         Begin MSForms.TextBox txtEMinDays 
            Height          =   300
            Left            =   1830
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
         Begin MSForms.TextBox txtEMinWaste 
            Height          =   300
            Left            =   3075
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
         Begin MSForms.TextBox txtSupervisor 
            Height          =   300
            Left            =   1215
            TabIndex        =   7
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
            TabIndex        =   6
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
            TabIndex        =   5
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
         Height          =   7080
         Left            =   90
         TabIndex        =   21
         Top             =   210
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   12488
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Process Description"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Supervisor"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Min. Days"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Max. Days"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Min. Wast."
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Max. Wast."
            Object.Width           =   1764
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   5850
         TabIndex        =   3
         Top             =   7350
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
         Left            =   7305
         TabIndex        =   1
         Top             =   7350
         Width           =   1425
         ForeColor       =   0
         Caption         =   " Close     "
         PicturePosition =   327683
         Size            =   "2514;635"
         Picture         =   "frmFDAProcesses.frx":0378
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
      TabIndex        =   12
      Top             =   270
      Width           =   4755
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   1830
      Top             =   7200
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Max. Days"
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
      Index           =   5
      Left            =   2085
      TabIndex        =   24
      Top             =   690
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Min. Days"
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
      Index           =   4
      Left            =   165
      TabIndex        =   23
      Top             =   690
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Min. DHR Wastage (%)"
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
      Left            =   3975
      TabIndex        =   22
      Top             =   690
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Max. DHR Wastage (%)"
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
      Left            =   5880
      TabIndex        =   20
      Top             =   690
      Width           =   1740
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
      Left            =   4920
      TabIndex        =   19
      Top             =   45
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
      TabIndex        =   2
      Top             =   60
      Width           =   1395
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   975
      Left            =   7830
      TabIndex        =   18
      Top             =   255
      Width           =   1050
      Caption         =   " Add"
      PicturePosition =   65543
      Size            =   "1852;1720"
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
Attribute VB_Name = "frmFDAProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAdd_Click()

    On Error GoTo err
    
    If txtDesc = "" Then
        MsgBox "Can't Add Null Values.", vbInformation
        Exit Sub
    End If
    
    LV.Sorted = False
    
    Dim Affect As Long
    
    Dim MaxSNo As Integer
    MaxSNo = Val(con.Execute("Select Max(SNo) From FDAProcesses").Fields(0) & "")
    MaxSNo = MaxSNo + 1
    
    con.Execute "Insert Into FDAProcesses (SNo," & _
     "Description,Supervisor,MinDays,MaxDays,MinWastePer,MaxWastePer) Values(" & _
     MaxSNo & ",'" & txtDesc & "','" & Text1 & "'," & _
     txtMinDays & "," & txtMaxDays & "," & txtMinWaste & "," & _
     txtMaxWaste & ")", Affect
     
    If Affect = 1 Then
        Call FillList
        txtDesc = ""
    End If
    
    'LV.Sorted = True
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
With Cr1
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
    
    SQL = "select * from FDAProcesses Order By Sno"

    With rsList
        .CursorLocation = adUseClient
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.Add(, ![ProcessID] & "'", ![Sno])
            ITM.SubItems(1) = ![Description] & ""
            ITM.SubItems(2) = ![Supervisor] & ""
            ITM.SubItems(3) = ![MinDays] & ""
            ITM.SubItems(4) = ![MaxDays] & ""
            ITM.SubItems(5) = ![MinWastePer] & ""
            ITM.SubItems(6) = ![MaxWastePer] & ""
            .MoveNext
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
    theKey = LV.SelectedItem.Key
    theKey = Right(theKey, Len(theKey) - 1)
    
    If KeyAscii = 13 Then
        con.Execute "Update TaxRanges Set SalFrom=" & _
         txtFromE & ",SalTo=" & txtToE & ",PerDeduct=" & _
         txtPerE & ",Reduction=" & txtDedE & _
         " Where SalFrom=" & theKey, Affect
        If Affect = 1 Then
            With LV.SelectedItem
                .Key = "'" & txtFromE
                .Text = txtFromE
                .SubItems(1) = txtToE
                .SubItems(2) = txtPerE
                .SubItems(3) = txtDedE
            End With
        End If
        PIC.Visible = False
    ElseIf KeyAscii = 27 Then
        txtFromE = ""
        txtToE = ""
        txtPerE = ""
        txtDedE = ""
        PIC.Visible = False
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

Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
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
    If LV.ListItems.Count = 0 Then Exit Sub
    '-------------------------------------------------------
    PicEdit.Move LV.Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(1).Width + LV.ColumnHeaders(2).Width + LV.ColumnHeaders(3).Width + LV.ColumnHeaders(4).Width + LV.ColumnHeaders(5).Width + LV.ColumnHeaders(6).Width + LV.ColumnHeaders(7).Width
    
    
    txtSNo.Move 0, 0, LV.ColumnHeaders(1).Width
    txtSNo.Text = LV.SelectedItem
    
    txtDescription.Move txtSNo.Width - 10, 0, LV.ColumnHeaders(2).Width
    txtDescription.Text = LV.SelectedItem.SubItems(1)
    
    txtSupervisor.Move txtDescription.Left + txtDescription.Width, 0, LV.ColumnHeaders(3).Width
    txtSupervisor.Text = LV.SelectedItem.SubItems(2)
    
    txtEMinDays.Move txtSupervisor.Left + txtSupervisor.Width + 10, 0, LV.ColumnHeaders(4).Width + 20
    txtEMinDays.Text = LV.SelectedItem.SubItems(3)
    
    txtEMaxDays.Move txtEMinDays.Left + txtEMinDays.Width - 20, 0, LV.ColumnHeaders(5).Width + 20
    txtEMaxDays.Text = LV.SelectedItem.SubItems(4)
    
    txtEMinWaste.Move txtEMaxDays.Left + txtEMaxDays.Width - 20, 0, LV.ColumnHeaders(6).Width + 20
    txtEMinWaste.Text = LV.SelectedItem.SubItems(5)
    
    txtEMaxWaste.Move txtEMinWaste.Left + txtEMinWaste.Width - 20, 0, LV.ColumnHeaders(7).Width + 20
    txtEMaxWaste.Text = LV.SelectedItem.SubItems(6)
    
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
    If LV.ListItems.Count = 0 Then Exit Sub
    On Error Resume Next
    If KeyCode = 46 Then
        If MsgBox("Do You Want To Delete This Component?", vbQuestion + vbYesNo) = vbYes Then
            Dim Affect As Integer
            con.Execute "Delete from Components Where CompID=" & Val(LV.SelectedItem.Key), Affect
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
    
    
    If KeyCode.value = 13 Then
        With txtSupervisor
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyCode.value = 27 Then
        PicEdit.Visible = False
    End If
    
    
End Sub




Private Sub txtEMaxWaste_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    Dim iAffect As Integer
    If KeyCode.value = 13 Then
        con.Execute "Update FDAProcesses Set SNo=" & txtSNo & _
         ",Description='" & txtDescription & _
         "',Supervisor='" & txtSupervisor & _
         "',MinDays=" & Val(txtEMinDays) & _
         ",MaxDays=" & Val(txtEMaxDays) & _
         ",MinWastePer=" & Val(txtEMinWaste) & _
         ",MaxWastePer=" & Val(txtEMaxWaste) & _
         " Where ProcessID=" & _
         Val(LV.SelectedItem.Key), iAffect
        If iAffect = 1 Then
            'LV.SelectedItem.Text = txtSNo
            'LV.SelectedItem.SubItems(1) = txtDescription
            PicEdit.Visible = False
            Call FillList
        End If
    ElseIf KeyCode.value = 27 Then
        PicEdit.Visible = False
    End If
End Sub

Private Sub txtSNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    If KeyCode.value = 13 Then
        With txtDescription
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyCode.value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtSupervisor_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    
    If KeyCode.value = 13 Then
        With txtEMinDays
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyCode.value = 27 Then
        PicEdit.Visible = False
    End If
End Sub

Private Sub txtEMinDays_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    
    If KeyCode.value = 13 Then
        With txtEMaxDays
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyCode.value = 27 Then
        PicEdit.Visible = False
    End If
End Sub

Private Sub txtEMaxDays_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    
    If KeyCode.value = 13 Then
        With txtEMinWaste
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyCode.value = 27 Then
        PicEdit.Visible = False
    End If
End Sub

Private Sub txtEMinWaste_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    
    If KeyCode.value = 13 Then
        With txtEMaxWaste
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyCode.value = 27 Then
        PicEdit.Visible = False
    End If
End Sub

