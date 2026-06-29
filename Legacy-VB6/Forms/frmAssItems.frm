VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendAssItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assigned Items"
   ClientHeight    =   6690
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   10095
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6525
      TabIndex        =   13
      Top             =   795
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cmbRate1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5925
      TabIndex        =   12
      Text            =   "cmbRate"
      Top             =   1785
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Frame fupdate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   390
      Left            =   90
      TabIndex        =   8
      Top             =   6090
      Visible         =   0   'False
      Width           =   5250
      Begin MSForms.CommandButton cmdUpdate 
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2160
         Caption         =   "Update Changes"
         PicturePosition =   327683
         Size            =   "3810;661"
         Accelerator     =   79
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClearAll 
         Height          =   375
         Left            =   3735
         TabIndex        =   10
         Top             =   0
         Width           =   1500
         ForeColor       =   0
         Caption         =   "Unassign All"
         PicturePosition =   327683
         Size            =   "2646;661"
         Picture         =   "frmAssItems.frx":0000
         Accelerator     =   108
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCheckAll 
         Height          =   375
         Left            =   2190
         TabIndex        =   9
         Top             =   0
         Width           =   1500
         ForeColor       =   0
         Caption         =   " Assign All"
         PicturePosition =   327683
         Size            =   "2646;661"
         Picture         =   "frmAssItems.frx":015A
         Accelerator     =   104
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin ComboList.Usercontrol1 cmbCont 
      Height          =   285
      Left            =   3555
      TabIndex        =   5
      Top             =   825
      Width           =   2220
      _ExtentX        =   3916
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
         Name            =   "MS Sans Serif"
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
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   825
      Width           =   2220
      _ExtentX        =   3916
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
      Appearance      =   0
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   0
      Top             =   0
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   570
      Top             =   0
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
            Picture         =   "frmAssItems.frx":02B4
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssItems.frx":0470
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssItems.frx":062C
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAssItems.frx":0A80
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned Items"
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
         Index           =   0
         Left            =   4260
         TabIndex        =   1
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned Items"
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
         Index           =   1
         Left            =   4275
         TabIndex        =   2
         Top             =   135
         Width           =   2295
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4785
      Left            =   105
      TabIndex        =   3
      Top             =   1215
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   8440
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
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
         Text            =   "Item ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   5997
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Rate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Plant"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Stamp"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Snaffing"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   4455
      BackColor       =   11517387
      Caption         =   "  Vendor                                     Nick Name"
      Size            =   "7858;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   375
      Left            =   8370
      TabIndex        =   4
      Top             =   6180
      Width           =   1665
      ForeColor       =   64
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmAssItems.frx":0ED4
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
Attribute VB_Name = "frmVendAssItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Add As Boolean, Edit As Boolean

Private Sub cmbCont_matched()
    Call RefreshList
End Sub

Private Sub cmbdept_matched()
   cmbCont.ClearVals
   Dim Cond As String
   Cond = IIf(cmbDept.ID = "0", "", "Where DeptID='" & cmbDept.ID & "'")
   Call cmbCont.AddVals(con, "Name", "Contractors", "EmpID", Cond)
   
End Sub

Private Sub cmbItem_matched()
On Error GoTo err
    Dim rs As New ADODB.Recordset
    cmbRate.Clear
    With rs
        .Open "select Rate1,Rate2,Rate3 from Items Where ItemID='" & cmbItem.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbRate.AddItem .Fields(0) & ""
        cmbRate.AddItem .Fields(1) & ""
        cmbRate.AddItem .Fields(2) & ""
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdAssign_Click()
    If cmbCont.MatchFound = False Then Exit Sub
    If cmbItem.MatchFound = False Then Exit Sub
    'If cmbRate.MatchFound = False Then Exit Sub
    
    Dim Saved As Integer
    con.Execute "Insert into AssignedItems(EmpID,ItemID,Rate) Values('" & _
                 cmbCont.ID & "','" & cmbItem.ID & "'," & cmbRate & ")", Saved
    If Saved = 1 Then
        Call RefreshList
    End If
End Sub



Private Sub cmdCheckAll_Click()

On Error Resume Next
For i = 1 To LV.ListItems.Count
   If Not LV.ListItems(i).Checked Then
      If LV.ListItems(i).Tag <> "" Then
         LV.ListItems(i).SubItems(3) = LV.ListItems(i).Tag
      Else
         Dim Rate As String
         Rate = Split(LV.ListItems(i).ListSubItems(3).Tag, ",")(0)
         LV.ListItems(i).SubItems(3) = Rate
      End If
      LV.ListItems(i).Checked = True
    End If
Next

End Sub

Private Sub cmdClearAll_Click()
For i = 1 To LV.ListItems.Count
   If LV.ListItems(i).Checked Then
         LV.ListItems(i).Tag = LV.ListItems(i).SubItems(3)
         LV.ListItems(i).SubItems(3) = ""
         LV.ListItems(i).Checked = False
    End If
Next
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Public Sub ShowInactive(m As Menu)

    Dim Cond As String
    If m.Checked Then
        Cond = " Where Active"
    Else
        Cond = ""
    End If
    m.Checked = Not m.Checked
    Call RefreshList(Cond)
    
End Sub

Private Sub cmdUpdate_Click()

   con.BeginTrans
   
   con.Execute "Delete From AssignedITems Where EMPID='" & cmbCont.ID & "'"
   For i = 1 To LV.ListItems.Count
      If LV.ListItems(i).Checked Then con.Execute "insert into AssignedItems(EMPID,ITemID,Rate) Values('" & cmbCont.ID & "','" & LV.ListItems(i) & "'," & Val(LV.ListItems(i).SubItems(3)) & ")"
   Next
   
   con.CommitTrans
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MainForm.mnuShowDept.Checked = False
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

LV.Sorted = True
If LV.SortKey = ColumnHeader.Index - 1 Then
   If LV.SortOrder = lvwAscending Then
      LV.SortOrder = lvwDescending
      ColumnHeader.Icon = "Up"
   Else
      LV.SortOrder = lvwAscending
      ColumnHeader.Icon = "Down"
   End If
Else
 LV.ColumnHeaderIcons = Nothing
 Set LV.ColumnHeaderIcons = ImageList2
 LV.SortKey = ColumnHeader.Index - 1
 LV.SortOrder = lvwAscending
 ColumnHeader.Icon = "Down"
End If

On Error Resume Next
LV.SelectedItem.EnsureVisible

End Sub

Private Sub LV_DblClick()

If LV.ListItems.Count = 0 Or cmdUpdate.Visible = False Then Exit Sub
If cmbCont.MatchFound = False Then Exit Sub

If Not LV.SelectedItem.Checked Then
   LV.SelectedItem.Checked = True
   Call LV_ItemCheck(LV.SelectedItem)
   Exit Sub
End If

txtRate.Move LV.Left + LV.ColumnHeaders(4).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(4).Width
txtRate.Text = LV.SelectedItem.SubItems(3)
txtRate.Visible = True
txtRate.SetFocus

End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    cmbCont.ListHeight = 1600
    cmbDept.ListHeight = 1600
    cmbCont.ClearVals
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "ContDepartments", "DeptID"
    cmbDept.Text = "<All Departments>"
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select *  from Items Where Active", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            'Check For Show All Option
            'ActiveCondition = IIf(mnuShowAll.Checked, "Where Left(ID,3)='CUS'", " Where not InActive and Left(ID,3)='CUS'")
            Set ITM = LV.ListItems.Add(, ![ITemID] & "'", ![ITemID] & "")
            ITM.ListSubItems.Add , , ![ItemName] & ""
            ITM.ListSubItems.Add , , ![Description] & ""
            ITM.ListSubItems.Add , , ![Rate1] & ""
            ITM.ListSubItems.Add , , ![PlantRate] & ""
            ITM.ListSubItems.Add , , ![StampRate] & ""
            ITM.ListSubItems.Add , , ![SnaffingRate] & ""
            'Store Rate To Show n Edit
            'ITM.ListSubItems.Add().Tag = Val(![Rate1] & "") & "," & Val(![PlantRate] & "") & "," & Val(![StampRate] & Val(![SnaffingRate]) & "")
            
            .MoveNext
        Loop
        .Close
        End With
        Set rs = Nothing
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
 Dim TitleBottom As Long

 TitleBottom = cmbCont.Top + cmbCont.Height 'fTitle.Top + fTitle.Height
 
 LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100)
 
 cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
 fupdate.Move LV.Left, LV.Top + LV.Height + 50

 
 
End Sub

Public Sub RefreshList(Optional Cond)

On Error GoTo err
    'Dim rs As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim PrevIndex As Double
    Dim TheColor As Long
    Dim ActiveCondition As String
    
    If IsMissing(Cond) Then
        Cond = " Where Active"
    End If
    
    For i = 1 To LV.ListItems.Count
      LV.ListItems(i).Checked = False
      LV.ListItems(i).SubItems(3) = ""
      'LV.ListItems(i).Tag = Split(LV.ListItems(i).ListSubItems(3).Tag, ",")(0)
    Next
    
    With rs
         .Open "Select ITemID,Rate  from AssignedItems Where EmpID='" & cmbCont.ID & "'", con, adOpenForwardOnly, adLockReadOnly
         Do Until .EOF
            LV.ListItems(![ITemID] & "'").Checked = True
            LV.ListItems(![ITemID] & "'").SubItems(3) = Format(![Rate] & "", "0.000")
            LV.ListItems(![ITemID] & "'").Tag = LV.ListItems(![ITemID] & "'").SubItems(3)
            .MoveNext
         Loop
         .Close
    End With
    
    Set rs = Nothing
    
    fupdate.Visible = True
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Public Sub showAssigned(ContID As String, ContName As String)
    cmbCont.Text = ContName
    Me.Show
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If Item.Checked Then
        Item.SubItems(3) = Format(Item.Tag, "0.000")
    Else
        Item.Tag = Item.SubItems(3)
        Item.SubItems(3) = ""
    End If

End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      LV.SelectedItem.SubItems(3) = Format(txtRate, "0.000")
   ElseIf KeyAscii = 27 Then
      txtRate.Visible = False
      LV.SetFocus
   Else
      KeyAscii = OnlyNumber(KeyAscii, True)
   End If

End Sub

Private Sub txtRate_LostFocus()
    txtRate.Visible = False
End Sub
