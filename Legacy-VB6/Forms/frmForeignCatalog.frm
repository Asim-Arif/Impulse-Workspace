VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmForeignCatalog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Foreign Catalogue..."
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
         Begin MSForms.TextBox txtCatal 
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
         Width           =   5790
         _ExtentX        =   10213
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
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Catalog"
            Object.Width           =   8819
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
         Picture         =   "frmForeignCatalog.frx":0000
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
      TabIndex        =   8
      Top             =   375
      Width           =   5295
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
      Top             =   7305
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
            Picture         =   "frmForeignCatalog.frx":0112
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForeignCatalog.frx":02CE
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Catalog"
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
      Top             =   150
      Width           =   555
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   600
      Left            =   5400
      TabIndex        =   2
      Top             =   90
      Width           =   615
      Caption         =   " Add"
      PicturePosition =   65543
      Size            =   "1085;1058"
      Picture         =   "frmForeignCatalog.frx":048A
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
Attribute VB_Name = "frmForeignCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAdd_Click()

    On Error GoTo ERR
    
    If txtDesc = "" Then
        MsgBox "Can't Add Blank Catalog Name.", vbInformation
        Exit Sub
    End If
    
    LV.Sorted = False
    
    Dim Affect As Long
    
    con.Execute "insert into ForeignCatalog(CatalogName)" & _
     " Values('" & txtDesc & "')", Affect
     
    If Affect = 1 Then
        Call FillList
        txtDesc = ""
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
    
    Sql = "select * from ForeignCatalog Order By CatalogName"

    With rsList
        .CursorLocation = adUseClient
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Call LV.ListItems.Add(, ![CatalID] & "'", ![CatalogName])
            'ITM.SubItems(1) = ![Description]
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
    PicEdit.Move LV.Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(1).Width
    
    
    txtCatal.Move 0, 0, LV.ColumnHeaders(1).Width
    txtCatal.Text = LV.SelectedItem
    
    'txtDescription.Move txtCatal.Width - 10, 0, LV.ColumnHeaders(2).Width
    'txtDescription.Text = LV.SelectedItem.SubItems(1)
    
    PicEdit.Visible = True
    
    With txtCatal
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



Private Sub txtCatal_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    On Error GoTo ERR
    Dim iAffect As Integer
    If KeyCode.value = 13 Then
        con.Execute "Update ForeignCatalog Set CatalogName='" & txtCatal & "' Where CatalID=" & Val(LV.SelectedItem.key), iAffect
        If iAffect = 1 Then
            PicEdit.Visible = False
            Call FillList
        End If
    ElseIf KeyCode.value = 27 Then
        PicEdit.Visible = False
    End If
    
    Exit Sub
ERR:
    MsgBox ERR.Description
    
End Sub
