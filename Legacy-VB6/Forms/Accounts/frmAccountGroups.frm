VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAccountGroups 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts Groups"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5355
   ScaleMode       =   0  'User
   ScaleWidth      =   6690.98
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FLV 
      Height          =   5310
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.ComboBox cmbTypeEdit 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3465
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   2445
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ComboBox cmbType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3885
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   450
         Width           =   1755
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   1530
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox txtCountry 
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
         Height          =   300
         Left            =   90
         TabIndex        =   1
         Top             =   465
         Width           =   3780
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4065
         Left            =   90
         TabIndex        =   4
         Top             =   750
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   7170
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Group Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   3096
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Sorting"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSForms.TextBox txtCurrency 
         Height          =   315
         Left            =   5655
         TabIndex        =   2
         Top             =   450
         Width           =   795
         VariousPropertyBits=   679495707
         Size            =   "1402;556"
         SpecialEffect   =   3
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   240
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   7125
         BackColor       =   11517387
         Caption         =   " Group Name                                                            Type                                Sort No."
         Size            =   "12568;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAdd 
         Height          =   345
         Left            =   6480
         TabIndex        =   3
         Top             =   420
         Width           =   735
         Caption         =   " Add"
         PicturePosition =   65543
         Size            =   "1296;609"
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
         Left            =   4290
         TabIndex        =   5
         Top             =   4890
         Visible         =   0   'False
         Width           =   1425
         ForeColor       =   0
         Caption         =   "Print         "
         PicturePosition =   327683
         Size            =   "2514;635"
         Picture         =   "frmAccountGroups.frx":0000
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
         Left            =   5775
         TabIndex        =   6
         Top             =   4890
         Width           =   1425
         ForeColor       =   0
         Caption         =   " Close     "
         PicturePosition =   327683
         Size            =   "2514;635"
         Picture         =   "frmAccountGroups.frx":0112
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
            Picture         =   "frmAccountGroups.frx":0224
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAccountGroups.frx":03E0
            Key             =   "Up"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAccountGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iColNo As Integer
Private Sub cmbCat_Click()
    Call FillList
End Sub

Private Sub cmbTypeEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        LV.SelectedItem.ListSubItems(1) = cmbTypeEdit.Text
        LV.SelectedItem.ListSubItems(1).Tag = cmbTypeEdit.ListIndex
        con.Execute "UPDATE GroupsForExpenseReport SET Type=" & cmbTypeEdit.ListIndex & " WHERE EntryID=" & Val(LV.SelectedItem.key)
        cmbTypeEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        cmbTypeEdit.Visible = False
    End If

End Sub

Private Sub cmbTypeEdit_LostFocus()
    cmbTypeEdit.Visible = False
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    
    LV.Sorted = False
    
    Dim Affect As Long
    
    If txtCountry = "" Then
        MsgBox "Invalid Grous Name ", vbInformation
        Exit Sub
    ElseIf cmbType.ListIndex = -1 Then
        MsgBox "Please Select Type.", vbInformation
        Exit Sub
    End If
    
    con.Execute "INSERT INTO GroupsForExpenseReport(GroupName,Type,SortNo) " & _
     "VALUES('" & txtCountry & "'," & cmbType.ListIndex & "," & Me.txtCurrency & ")", Affect
     
    If Affect = 1 Then
        Call FillList
'        txtCurrency = ""
    End If
    txtCountry.SetFocus
    LV.Sorted = True
    Exit Sub
err:
    If err.Number = -2147217900 Then
        MsgBox "This Record Already Exists.", vbInformation
    Else
        MsgBox err.Description
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()
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
With cr1
    .ReportFileName = RptPath & "\Holidays.rpt"
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
    
    SQL = "SELECT * FROM GroupsForExpenseReport Order By SortNo"

    With rsList
        .CursorLocation = adUseClient
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.add(, ![EntryID] & "'", !GroupName)
            ITM.ListSubItems.add(, , IIf(Val(!Type & "") = 0, "Expense", "Income")).Tag = Val(!Type & "")
            ITM.ListSubItems.add , , Val(![SortNo] & "")
            .MoveNEXT
        Next
        .Close
    End With
    Set rsList = Nothing
    SQL = ""
    
    txtCurrency = GetSingleLongValue("MAX(SortNo)", "GroupsForExpenseReport") + 1
    
End Sub


Private Sub Form_Load()
    Call FillList
    With cmbType
        .AddItem "Expense"
        .AddItem "Income"
        .ListIndex = 0
    End With
    '--------------------------
    With cmbTypeEdit
        .AddItem "Expense"
        .AddItem "Income"
        .ListIndex = 0
    End With
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

    If LV.ListItems.count = 0 Then Exit Sub
    
    If iColNo = 2 Then
        With cmbTypeEdit
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
            .ListIndex = Val(LV.SelectedItem.ListSubItems(1).Tag)
            .Visible = True
            .SetFocus
        End With
    Else
        With txtEdit
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
            If iColNo = 1 Then
                .Text = LV.SelectedItem.Text
            Else
                .Text = LV.SelectedItem.SubItems(iColNo - 1)
            End If
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    End If
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'PicEdit.Visible = False
'      With cmdDelete
'        .Move LV.Left + LV.ColumnHeaders(LV.ColumnHeaders.count).Width + LV.ColumnHeaders(LV.ColumnHeaders.count).Left + 10, Item.Top + LV.Top
'        .Visible = True
'    End With
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
        
    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        If MsgBox("Remove This Item ?", vbYesNo) = vbNo Then Exit Sub
        Dim Affect As Long
        con.Execute "Delete FROM GroupsForExpenseReport WHERE EntryID=" & Val(LV.SelectedItem.key)
    
        Call FillList
    End If
    
    Exit Sub
err:
    MsgBox err.Description
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
        If iColNo = 1 Then
            If txtEdit = "" Then
                MsgBox "Invalid Group Name.", vbInformation
                Exit Sub
            End If
            LV.SelectedItem.Text = txtEdit
            con.Execute "UPDATE GroupsForExpenseReport SET GroupName='" & txtEdit & "' WHERE EntryID=" & Val(LV.SelectedItem.key)
        Else
            LV.SelectedItem.SubItems(iColNo - 1) = Val(txtEdit)
            con.Execute "UPDATE GroupsForExpenseReport SET SortNo=" & Val(txtEdit) & " WHERE EntryID=" & Val(LV.SelectedItem.key)
        End If
        
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
End Sub
