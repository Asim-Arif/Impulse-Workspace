VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vBalIml6.ocx"
Begin VB.Form frmADUsers 
   BorderStyle     =   0  'None
   Caption         =   "Employee Ledger..."
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11550
   LinkTopic       =   "Form19"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleMode       =   0  'User
   ScaleWidth      =   10345.49
   ShowInTaskbar   =   0   'False
   Begin vbalIml6.vbalImageList ilsIcons 
      Left            =   10500
      Top             =   15
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   70
      IconSizeY       =   14
      Size            =   8296
      Images          =   "frmADUsers.frx":0000
      Version         =   131072
      KeyCount        =   2
      Keys            =   "ÿ"
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   11490
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Domain Users"
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
         Index           =   1
         Left            =   4410
         TabIndex        =   2
         Top             =   120
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Domain Users"
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
         Height          =   390
         Index           =   2
         Left            =   4425
         TabIndex        =   3
         Top             =   135
         Width           =   2145
      End
   End
   Begin VB.Frame FLV 
      Height          =   6900
      Left            =   60
      TabIndex        =   0
      Top             =   525
      Width           =   11385
      Begin vbAcceleratorSGrid6.vbalGrid UsersGrid 
         Height          =   6150
         Left            =   60
         TabIndex        =   6
         Top             =   150
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   10848
         GridLines       =   -1  'True
         NoVerticalGridLines=   -1  'True
         GridLineMode    =   1
         BackgroundPictureHeight=   392
         BackgroundPictureWidth=   356
         GridLineColor   =   15263976
         GridFillLineColor=   15263976
         GroupingAreaBackColor=   14737632
         AlternateRowBackColor=   -2147483648
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderButtons   =   0   'False
         HeaderDragReorderColumns=   0   'False
         HeaderHotTrack  =   0   'False
         DisableIcons    =   -1  'True
         DrawFocusRectangle=   0   'False
         SelectionOutline=   -1  'True
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   480
         Left            =   5850
         TabIndex        =   4
         Top             =   6330
         Width           =   5400
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   2460
            TabIndex        =   7
            Top             =   105
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Update"
            PicturePosition =   327683
            Size            =   "2514;635"
            Accelerator     =   85
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
            Left            =   3930
            TabIndex        =   5
            Top             =   105
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Close      "
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
   End
End
Attribute VB_Name = "frmADUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalPresents As Double, TotalAbsents As Double, TotalLeaves As Double
Dim SQL As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
    Dim StrSQL As String
    
    Call StartTrans(con)
    
    With UsersGrid
        For i = 1 To .Rows
            StrSQL = ""
            For j = 2 To .Columns
                StrSQL = StrSQL & .ColumnKey(j) & "=" & .CellExtraIcon(i, j) & ","
            Next j
            con.Execute "Update UserMainSettings Set " & Left(StrSQL, Len(StrSQL) - 1) & " Where UserName='" & .CellText(i, 1) & "'"
        Next i
    End With
    con.CommitTrans
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Load()

    Call ConfigureGrid
    Call FillGrid

End Sub
Private Sub Form_Resize()


    On Error Resume Next

    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2

    FLV.Move 50, 525, Me.ScaleWidth - 100, Me.ScaleHeight - 700
 
    UsersGrid.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 
 
    fButs1.Move UsersGrid.Left + UsersGrid.Width - fButs1.Width, UsersGrid.Top + UsersGrid.Height
 
 
End Sub



Private Sub ConfigureGrid()

    On Error GoTo err
    
    With UsersGrid
        
        .Editable = False
        .ImageList = ilsIcons
        
        .DefaultRowHeight = 18
        .RowMode = False
        
        .AddColumn "Name", "Name", ecgHdrTextALignCentre, , 80, , True
        .AddColumn "Company", "Company", ecgHdrTextALignCentre, , 70, , True
        .AddColumn "Customers", "Customers", ecgHdrTextALignCentre, , 70, , True
        .AddColumn "Cost", "Cost", ecgHdrTextALignCentre, , 70, , True
        .AddColumn "Inventory", "Inventory", ecgHdrTextALignCentre, , 70, , True
        .AddColumn "Imports", "Imports", ecgHdrTextALignCentre, , 70, , True
        .AddColumn "Production", "Production", ecgHdrTextALignCentre, , 70, , True
        .AddColumn "Payroll", "Payroll", ecgHdrTextALignCentre, , 70, , True
        .AddColumn "Accounting", "Accounting", ecgHdrTextALignCentre, , 70, , True
        .AddColumn "Export", "Export", ecgHdrTextALignCentre, , 70, , True
        .AddColumn "Management", "Management", ecgHdrTextALignCentre, , 70, , True
        
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub FillGrid()
    
    On Error GoTo err
    
    With UsersGrid
        .Clear
        .StretchLastColumnToFit = False
        .Redraw = False
        .Visible = False
    End With
    Dim rs As New ADODB.Recordset
    
    rs.Open "Select * From UserMainSettings", con, adOpenForwardOnly, adLockReadOnly
        
        
    
    '-/-/-/-/-/--/-/---/-/-/--/-
        
    Dim sFnt As New StdFont
    sFnt.Name = "Tahoma"
    sFnt.Size = 8
    
    Dim iRow As Integer
    
    '-/-/-/-/-/--/-/---/-/-/--/-
    Dim TheDomain As IADsDomain
    
    Dim USer As IADsUser
    
    Dim strDomain As String
    Dim iMousePointer As Integer

    'Accept the Domain name
    strDomain = "Vision"

    'Use the WinNT Directory Services
    strDomain = "WinNT://" & strDomain

    'Initialize the User Interface
    
    iMousePointer = Me.MousePointer
    Me.MousePointer = vbHourglass

    'Create the Domain object
     
    Set TheDomain = GetObject(strDomain)
    
    'Search for Computers in the Domain
    TheDomain.Filter = Array("User")
     
    For Each USer In TheDomain
    
        If USer.AccountDisabled = False And USer.IsAccountLocked = False And Right(USer.Name, 7) <> "_SERVER" Then
            
            iRow = UsersGrid.Rows + 1
            
            UsersGrid.CellDetails iRow, 1, USer.Name
            rs.Filter = "UserName='" & USer.Name & "'"
            
            If rs.EOF Then
'                For i = 2 To UsersGrid.Columns
'                    UsersGrid.CellExtraIcon(iRow, i) = 0
'                Next
'                con.Execute "Insert Into UserMainSettings(UserName) Values('" & USer.Name & "')"
            Else
                On Error Resume Next
                For i = 2 To UsersGrid.Columns
                    UsersGrid.CellExtraIcon(iRow, i) = Abs(rs.Fields(UsersGrid.ColumnKey(i)))
                Next
                On Error GoTo 0
            End If
            
        End If
        
        
    Next USer
     
    
    With UsersGrid
        '.StretchLastColumnToFit = True
        .Redraw = True
        .Visible = True
    End With
    'Clean up
    Me.MousePointer = iMousePointer
    Set USer = Nothing
    Set TheDomain = Nothing
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub UsersGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, bDoDefault As Boolean)
    
    Dim lCol As Long, lRow As Long
    Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
    
    Dim lIconIndex As Long
    Dim i As Long
    '15
    
    UsersGrid.CellFromPoint x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY, lRow, lCol
    'UsersGrid.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
    If lRow = 0 Then Exit Sub
    If lCol > 1 Then
        If UsersGrid.CellExtraIcon(lRow, lCol) = 0 Then
           UsersGrid.CellExtraIcon(lRow, lCol) = 1
        Else
            UsersGrid.CellExtraIcon(lRow, lCol) = 0
        End If
    End If

End Sub

