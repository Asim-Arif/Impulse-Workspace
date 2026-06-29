VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmItemGroups 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Components..."
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7620
   ScaleMode       =   0  'User
   ScaleWidth      =   5454.896
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   10
      Top             =   375
      Width           =   4260
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      ClipControls    =   0   'False
      ForeColor       =   &H00800000&
      HasDC           =   0   'False
      Height          =   300
      Left            =   4380
      ScaleHeight     =   240
      ScaleWidth      =   915
      TabIndex        =   9
      Top             =   375
      Width           =   975
   End
   Begin VB.Frame FLV 
      Height          =   6870
      Left            =   60
      TabIndex        =   0
      Top             =   705
      Width           =   5985
      Begin VB.VScrollBar VScroll 
         Height          =   6120
         LargeChange     =   2
         Left            =   5610
         TabIndex        =   12
         Top             =   210
         Width           =   285
      End
      Begin VB.PictureBox PicColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   195
         Index           =   0
         Left            =   1755
         ScaleHeight     =   195
         ScaleWidth      =   915
         TabIndex        =   8
         Top             =   6480
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   825
         Top             =   6270
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   405
         ScaleHeight     =   255
         ScaleWidth      =   4140
         TabIndex        =   6
         Top             =   1995
         Visible         =   0   'False
         Width           =   4170
         Begin VB.TextBox txtDescription 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -15
            TabIndex        =   7
            Top             =   -15
            Width           =   3990
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6150
         Left            =   90
         TabIndex        =   2
         Top             =   195
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   10848
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Component Description"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Color"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   2985
         TabIndex        =   5
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
         Left            =   4470
         TabIndex        =   1
         Top             =   6420
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
      Top             =   7365
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
            Picture         =   "frmItemGroups.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemGroups.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Color"
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
      Left            =   4455
      TabIndex        =   11
      Top             =   150
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Component Description"
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
      TabIndex        =   4
      Top             =   150
      Width           =   1665
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   600
      Left            =   5400
      TabIndex        =   3
      Top             =   90
      Width           =   615
      Caption         =   " Add"
      PicturePosition =   65543
      Size            =   "1085;1058"
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
Attribute VB_Name = "frmItemGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strTable As String

Public Sub ShowForm(argTable As String, Optional argCap As String, Optional LabCap As String)
    On Error GoTo err
    
    Me.Caption = argCap
    Label1(1) = LabCap
    strTable = argTable
    LV.ColumnHeaders(2).Text = LabCap
    Call FillList
    Me.Show 1
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    
    LV.Sorted = False
    
    Dim Affect As Long
    
        
    con.Execute "insert into " & strTable & "(" & _
     "Description,GrpColor) Values('" & txtDesc & "'," & Picture1.BackColor & ")", Affect
    
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
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub FillList()

    
    Dim rsList As New ADODB.Recordset
    Dim ITM As ListItem
    Dim SQL As String
    
    SQL = "select * from " & strTable & " Order By Description" '& Components "
    For i = 1 To PicColor.UBound
        Unload PicColor(i)
    Next i
    With rsList
        .CursorLocation = adUseClient
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        For i = 0 To .RecordCount - 1
        
            Set ITM = LV.ListItems.add(, ![ID] & "'", i + 1)
            ITM.SubItems(1) = ![Description] & ""
            
            ITM.Tag = Val(!GrpColor & "")
            
            If i < 28 Then
                Load PicColor(i + 1)

                PicColor(i + 1).Move LV.Left + LV.ColumnHeaders(3).Left + 30, LV.Top + LV.ListItems(i + 1).Top + 30, LV.ColumnHeaders(3).Width - 10
                PicColor(i + 1).BackColor = Val(![GrpColor] & "")
                'PicColor(i + 1).ForeColor = vbYellow
                PicColor(i + 1).ZOrder

                PicColor(i + 1).Visible = True
            End If

            .MoveNEXT
        Next
        .Close
    End With
    Set rsList = Nothing
    SQL = ""
    
    With VScroll
        .Min = 1
        
        .Max = LV.ListItems.count - 27
        
        .Value = 1
    End With
        
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
    'Picture1.Print "Dbl Click To Change Color"
End Sub

Private Sub Form_Resize()
    Me.Top = 1700
End Sub


Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Exit Sub '
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

    If LV.ListItems.count = 0 Then Exit Sub
    
    PicEdit.Move LV.Left + LV.ColumnHeaders(2).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(2).Width
    PicEdit.Visible = True
    
    With txtDescription
        .Move 0, 0, PicEdit.Width
        .Text = LV.SelectedItem.SubItems(1)
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    

    If LV.ListItems.count = 0 Then Exit Sub
    
    
    On Error Resume Next
    If KeyCode = 46 Then
        Dim lCount As Long, lGroupID As Long
        lGroupID = Val(LV.SelectedItem.key)
        lCount = GetSingleLongValue("COUNT(*)", "RM", " WHERE GroupID=" & lGroupID)
        If lCount > 0 Then
            MsgBox "Can't Delete, Materials are linked.", vbInformation
            Exit Sub
        End If
        If MsgBox("Do You Want To Delete This Entry?", vbQuestion + vbYesNo) = vbYes Then
            Dim Affect As Integer
            con.Execute "DELETE FROM " & strTable & " WHERE ID=" & lGroupID, Affect
            If Affect = 1 Then
                LV.ListItems.Remove (LV.SelectedItem.Index)
            End If
        End If
    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub PicColor_DblClick(Index As Integer)
    
    
    On Error GoTo err
    Dim LColor As Long
    CD1.ShowColor
    If CD1.Color <> 0 Then
        LColor = CD1.Color
    End If
    
    
    Dim iAffect As Integer
    con.Execute "Update " & strTable & " Set GrpColor=" & LColor & " Where ID=" & Val(PicColor(Index).Tag), iAffect
    If iAffect = 1 Then
        PicColor(Index).BackColor = LColor
    End If
        
    Exit Sub
err:
    If err.Number = 32755 Then Exit Sub
    MsgBox err.Description

End Sub

Private Sub Picture1_Click()

    On Error GoTo err
    CD1.ShowColor
    If CD1.Color <> 0 Then
        Picture1.BackColor = CD1.Color
    End If
    
    Exit Sub
err:
    If err.Number = 32755 Then Exit Sub
    MsgBox err.Description

End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    If KeyAscii = 13 Then
        Dim iAffect As Integer
        con.Execute "Update " & strTable & " Set Description='" & txtDescription & "' Where ID=" & Val(LV.SelectedItem.key), iAffect
        If iAffect = 1 Then
            LV.SelectedItem.SubItems(1) = txtDescription
        End If
        PicEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub txtDescription_Validate(cancel As Boolean)
    If PicEdit.Visible Then PicEdit.Visible = False
End Sub

Private Sub VScroll_Change()
    On Error Resume Next
    Dim i As Integer
    LV.ListItems(VScroll.Value + 27).EnsureVisible
    LV.ListItems(VScroll.Value).EnsureVisible
    
'    For i = VScroll.value To VScroll.value + 27
'        PicColor(i - VScroll.value).BackColor = LV.ListItems(i).Tag
'    Next
    For i = 0 To 27
        PicColor(i + 1).BackColor = Val(LV.ListItems(VScroll.Value + i).Tag)
        PicColor(i + 1).Tag = Val(LV.ListItems(VScroll.Value + i).key)
    Next
End Sub

