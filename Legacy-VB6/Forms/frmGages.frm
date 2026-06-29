VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmGages 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Steel Types"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7620
   ScaleMode       =   0  'User
   ScaleWidth      =   5468.332
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbSteel 
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Top             =   300
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Appearance      =   0
   End
   Begin VB.Frame FLV 
      ForeColor       =   &H00800000&
      Height          =   6960
      Left            =   75
      TabIndex        =   0
      Top             =   615
      Width           =   5955
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5385
         Picture         =   "frmGages.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Remove This Entry From Voucher"
         Top             =   780
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   270
         ScaleHeight     =   255
         ScaleWidth      =   4140
         TabIndex        =   3
         Top             =   1350
         Visible         =   0   'False
         Width           =   4170
         Begin MSForms.ComboBox cmbEUnit 
            Height          =   300
            Left            =   585
            TabIndex        =   14
            Tag             =   "Unit"
            Top             =   -15
            Width           =   1125
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1984;529"
            ListWidth       =   3527
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtGage 
            Height          =   300
            Left            =   -30
            TabIndex        =   4
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
      Begin Crystal.CrystalReport cr1 
         Left            =   330
         Top             =   6720
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
         Left            =   810
         Top             =   6720
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
               Picture         =   "frmGages.frx":0532
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGages.frx":06EE
               Key             =   "Up"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5715
         Left            =   90
         TabIndex        =   12
         Top             =   750
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   10081
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Guage"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Unit"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSForms.TextBox txtGuage 
         Height          =   300
         Left            =   90
         TabIndex        =   13
         Top             =   375
         Width           =   3855
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6800;529"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbUnit 
         Height          =   300
         Left            =   3990
         TabIndex        =   11
         Tag             =   "Unit"
         Top             =   375
         Width           =   1125
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1984;529"
         ListWidth       =   3527
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Unit"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4005
         TabIndex        =   10
         Top             =   165
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Guage"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   165
         Width           =   465
      End
      Begin MSForms.CommandButton cmdAdd 
         Height          =   345
         Left            =   5160
         TabIndex        =   5
         Top             =   360
         Width           =   660
         Caption         =   " Add"
         PicturePosition =   65543
         Size            =   "1164;609"
         Accelerator     =   65
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   3000
         TabIndex        =   2
         Top             =   6510
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
         Left            =   4455
         TabIndex        =   1
         Top             =   6510
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Steel Types"
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
      Left            =   105
      TabIndex        =   7
      Top             =   90
      Width           =   990
   End
End
Attribute VB_Name = "frmGages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub EditGage(iGageID As String)
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select SteelID From SteelGages Where GageID='" & iGageID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            cmbSteel.ID = ![SteelID] & ""
        End If
        .Close
    End With
    Set rs = Nothing
    
    Me.Show 1
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbSteel_matched()
    Call FillList
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    
    If cmbSteel.MatchFound = False Then Exit Sub
    If cmbUnit.MatchFound = False Then Exit Sub
    LV.Sorted = False
    
    Dim Affect As Long
    
    Dim iNewID As Long
    iNewID = Val(con.Execute("Select Max(Cast(Substring(GageID,charindex('-',GageID)+1,50) as int)) From SteelGages Where SteelID=" & cmbSteel.ID).Fields(0).Value & "")
    iNewID = iNewID + 1
    Dim strNewID As String
    strNewID = cmbSteel.ID & "-" & iNewID
    con.Execute "insert into SteelGages(GageID,SteelID,Gage,GageUnit) " & _
     "Values('" & strNewID & "'," & cmbSteel.ID & ",'" & txtGuage & "','" & cmbUnit & "')", Affect
     
    If Affect = 1 Then
        Call FillList
        txtGuage = ""
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
'''''''''''''''
   Exit Sub
'''''''''''''''

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
    
    On Error GoTo err
    
    If cmbSteel.MatchFound = False Then Exit Sub
    
    Dim rsList As New ADODB.Recordset
    Dim ITM As ListItem
    Dim SQL As String
    
    SQL = "SELECT * FROM SteelGages WHERE SteelID=" & cmbSteel.ID & " Order By GageID"

    With rsList
        .CursorLocation = adUseClient
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.add(, ![GageID] & "'", ![Gage])
            ITM.SubItems(1) = ![GageUnit]
            .MoveNEXT
        Next
        .Close
    End With
    
    Set rsList = Nothing
    SQL = ""
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub Form_Load()
    cmbSteel.ListHeight = 2000
    Call cmbSteel.AddVals(con, "SteelType", "SteelTypes", "SteelID")
    Call AddToCombo(cmbUnit, "UnitName", "Units")
    Call AddToCombo(cmbEUnit, "UnitName", "Units")
End Sub

Private Sub Form_Resize()
    Me.Top = 1700 'MainForm.Toolbar1.Top + MainForm.Toolbar1.Height + 100
    cmbSteel.Height = 1800
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
    If LV.ListItems.count = 0 Then Exit Sub
    '-------------------------------------------------------
    PicEdit.Move LV.Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(1).Width + LV.ColumnHeaders(2).Width
    
    
    txtGage.Move 0, 0, LV.ColumnHeaders(1).Width
    txtGage.Text = LV.SelectedItem
    
    cmbEUnit.Move txtGage.Width - 10, 0, LV.ColumnHeaders(2).Width
    cmbEUnit.Text = LV.SelectedItem.SubItems(1)
    
    PicEdit.Visible = True
    
    With txtGage
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


Private Sub cmbEUnit_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    Dim iAffect As Integer
    If KeyCode.Value = 13 Then
        con.Execute "UPDATE SteelGages SET Gage='" & txtGage & _
         "',GageUnit='" & cmbEUnit & "' WHERE GageID='" & _
         LV.SelectedItem.key, iAffect
        If iAffect = 1 Then
            'LV.SelectedItem.Text = txtGage
            'LV.SelectedItem.SubItems(1) = cmbEUnit
            PicEdit.Visible = False
            Call FillList
        End If
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtGage_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    If KeyCode.Value = 13 Then
        With cmbEUnit
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub


