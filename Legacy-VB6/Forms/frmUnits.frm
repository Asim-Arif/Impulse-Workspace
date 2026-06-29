VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmUnits 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Processes"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7620
   ScaleMode       =   0  'User
   ScaleWidth      =   5428.024
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUnitDesc 
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
      Left            =   1185
      TabIndex        =   10
      Top             =   375
      Width           =   4155
   End
   Begin VB.Frame FLV 
      Height          =   6870
      Left            =   60
      TabIndex        =   0
      Top             =   705
      Width           =   5955
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   105
         ScaleHeight     =   255
         ScaleWidth      =   4140
         TabIndex        =   6
         Top             =   645
         Visible         =   0   'False
         Width           =   4170
         Begin MSForms.TextBox txtDescription 
            Height          =   300
            Left            =   600
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
         Begin MSForms.TextBox txtName 
            Height          =   300
            Left            =   -15
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
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6150
         Left            =   90
         TabIndex        =   5
         Top             =   210
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   10848
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Unit Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Unit Description"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   2925
         TabIndex        =   4
         Top             =   6420
         Visible         =   0   'False
         Width           =   1425
         ForeColor       =   0
         Caption         =   "Print         "
         PicturePosition =   327683
         Size            =   "2514;635"
         Picture         =   "frmUnits.frx":0000
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
         Left            =   4440
         TabIndex        =   1
         Top             =   6420
         Width           =   1425
         ForeColor       =   0
         Caption         =   " Close     "
         PicturePosition =   327683
         Size            =   "2514;635"
         Picture         =   "frmUnits.frx":0112
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
   Begin VB.TextBox txtUnit 
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
      TabIndex        =   9
      Top             =   375
      Width           =   1095
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   6345
      Top             =   240
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
      Left            =   360
      Top             =   5445
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
            Picture         =   "frmUnits.frx":0224
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnits.frx":03E0
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Unit Description"
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
      Left            =   1245
      TabIndex        =   11
      Top             =   180
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Unit Name"
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
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   735
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   600
      Left            =   5370
      TabIndex        =   2
      Top             =   90
      Width           =   615
      Caption         =   " Add"
      PicturePosition =   65543
      Size            =   "1085;1058"
      Picture         =   "frmUnits.frx":059C
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
Attribute VB_Name = "frmUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

    On Error GoTo ERR
    
    LV.Sorted = False
    
    Dim iAffect As Integer
    
    con.Execute "Insert into Units (UnitName,UnitDesc)" & _
     " Values('" & txtUnit & "','" & txtUnitDesc & "')", iAffect
     
    If iAffect = 1 Then
        Call FillList
        txtUnit = ""
        txtUnitDesc = ""
        txtUnit.SetFocus
    End If
    
    LV.Sorted = True
    Exit Sub
ERR:
    MsgBox ERR.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

On Error GoTo ERR
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
ERR:
    MsgBox ERR.Description
End Sub

Private Sub FillList()

    
    Dim rsList As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Sql As String
    
    Sql = "select * from Units Order By UnitName"

    With rsList
        .CursorLocation = adUseClient
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.Add(, ![UnitID] & "'", ![UnitName] & "")
            ITM.SubItems(1) = ![UnitDesc] & ""
            .MoveNext
        Next
        .Close
    End With
    
    Set rsList = Nothing
    Sql = ""
    
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
    PicEdit.Move LV.Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(1).Width + LV.ColumnHeaders(2).Width
    
    
    txtName.Move 0, 0, LV.ColumnHeaders(1).Width
    txtName.Text = LV.SelectedItem
    
    txtDescription.Move txtName.Width - 10, 0, LV.ColumnHeaders(2).Width
    txtDescription.Text = LV.SelectedItem.SubItems(1)
    
    PicEdit.Visible = True
    
    With txtName
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
            con.Execute "Delete from Components Where CompID=" & Val(LV.SelectedItem.key), Affect
            If Affect = 1 Then
                LV.ListItems.Remove (LV.SelectedItem.Index)
            End If
        End If
    End If
    Exit Sub
ERR:
    MsgBox ERR.Description
End Sub


Private Sub txtDescription_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    Dim iAffect As Integer
    If KeyCode.value = 13 Then
        con.Execute "Update Units Set UnitName='" & txtName & _
         "',UnitDesc='" & txtDescription & "' Where UnitID=" & _
         Val(LV.SelectedItem.key), iAffect
        If iAffect = 1 Then
            'LV.SelectedItem.Text = txtName
            'LV.SelectedItem.SubItems(1) = txtDescription
            PicEdit.Visible = False
            Call FillList
        End If
    ElseIf KeyCode.value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
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
