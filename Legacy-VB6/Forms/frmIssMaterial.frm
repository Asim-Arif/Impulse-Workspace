VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmIssMaterial 
   BorderStyle     =   0  'None
   Caption         =   "Issue Material To Contractor/Supervisors"
   ClientHeight    =   6645
   ClientLeft      =   1050
   ClientTop       =   -105
   ClientWidth     =   11505
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
   ScaleHeight     =   6645
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbItem 
      Height          =   285
      Left            =   7410
      TabIndex        =   23
      Top             =   1260
      Width           =   3405
      _ExtentX        =   6006
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
   Begin ComboList.Usercontrol1 cmbMaterials 
      Height          =   285
      Left            =   1455
      TabIndex        =   21
      Top             =   2115
      Visible         =   0   'False
      Width           =   3390
      _ExtentX        =   5980
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
   Begin VB.CheckBox chkOrderNo 
      Caption         =   "Show Customer Order No's"
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
      Height          =   240
      Left            =   510
      TabIndex        =   16
      Top             =   6135
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print When Saved"
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
      Height          =   240
      Left            =   510
      TabIndex        =   6
      Top             =   5880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2580
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   8070
      TabIndex        =   11
      Top             =   705
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dddd dd-MMM-yyyy"
      Format          =   20774915
      CurrentDate     =   -8461
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4305
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdUnAssign 
      Height          =   240
      Left            =   10800
      Picture         =   "frmIssMaterial.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "UnAssign Selected Item..."
      Top             =   1545
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   600
      Top             =   2055
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
            Picture         =   "frmIssMaterial.frx":0102
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIssMaterial.frx":02BE
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   465
      Top             =   3330
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
            Picture         =   "frmIssMaterial.frx":047A
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIssMaterial.frx":08CE
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4080
      Left            =   225
      TabIndex        =   1
      Top             =   1545
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7197
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S.No"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Material ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Material Name"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Qty In Stock"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Qty Required"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Prev Issued"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Qty To Issue"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Issue Qty"
         Object.Width           =   2028
      EndProperty
   End
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   225
      TabIndex        =   7
      Top             =   0
      Width           =   10575
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Material To Contractor/Supervisors"
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
         Left            =   2385
         TabIndex        =   8
         Top             =   120
         Width           =   6435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Material To Contractor/Supervisors"
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
         Left            =   2400
         TabIndex        =   9
         Top             =   135
         Width           =   6435
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   5535
      TabIndex        =   13
      Top             =   5805
      Width           =   5265
      Begin MSForms.CommandButton cmdSave 
         Height          =   360
         Left            =   75
         TabIndex        =   3
         Top             =   180
         Width           =   1665
         Caption         =   "Save & New   "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   110
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCancel 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   3510
         TabIndex        =   5
         Top             =   180
         Width           =   1665
         Caption         =   "Cancel            "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdSaveClose 
         Height          =   360
         Left            =   1800
         TabIndex        =   4
         Top             =   180
         Width           =   1665
         Caption         =   "Save & Close  "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label lbltotalBalls 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3945
         TabIndex        =   18
         Top             =   150
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Balls Issued :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   1
         Left            =   2340
         TabIndex        =   17
         Top             =   180
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin MSForms.TextBox txtIncharge 
      Height          =   285
      Left            =   2340
      TabIndex        =   22
      Top             =   1260
      Width           =   2775
      VariousPropertyBits=   746604575
      BorderStyle     =   1
      Size            =   "4895;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbOrderNo 
      Height          =   285
      Left            =   5115
      TabIndex        =   20
      Top             =   1260
      Width           =   2310
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4075;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbProcess 
      Height          =   285
      Left            =   225
      TabIndex        =   19
      Top             =   1260
      Width           =   2130
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3757;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issued On :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   7065
      TabIndex        =   15
      Top             =   735
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issue # :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   10
      Left            =   225
      TabIndex        =   14
      Top             =   750
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSForms.TextBox txtRecNo 
      Height          =   285
      Left            =   1095
      TabIndex        =   12
      Top             =   705
      Visible         =   0   'False
      Width           =   1980
      VariousPropertyBits=   746604575
      BorderStyle     =   1
      Size            =   "3492;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   240
      Left            =   225
      TabIndex        =   10
      Top             =   1020
      Width           =   10590
      BackColor       =   11517387
      Caption         =   " Process                                Incharge Name                                 Consolidation No.                   Item"
      Size            =   "18680;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmIssMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VendCapacity As Long
Dim IsItemsFirst As Boolean
Dim EditMaterial As Boolean


Private Sub chkOrderNo_Click()
'   Call cmbCountry_Click
   
      For i = 1 To LV.ListItems.count
         With LV.ListItems(i)
            .SubItems(1) = IIf(chkOrderNo = vbChecked, Replace(.ListSubItems(1).key, "'", ""), .ListSubItems(1).Tag)
         End With
      Next
   
End Sub


Private Sub cmbItem_matched()

    On Error GoTo err
    
    If Not cmbItem.MatchFound Then Exit Sub
    
    Dim LastGroup As String
    Dim rs As New ADODB.Recordset
    With rs
        
        .Open "Select * From VRefORdMatToIssue Where ItemID='" & cmbItem.ID & _
        "' And ProcessID=" & cmbProcess.List(cmbProcess.ListIndex, 1) & " and RefORderNo='" & _
        cmbOrderNo.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        
        Dim ITM As ListItem
        LV.ListItems.Clear
        
        
        Do Until .EOF
        
            Set ITM = LV.ListItems.add(, ![MatID], LV.ListItems.count + 1)
            
            
            ITM.Tag = ![Unit] & ""
            ITM.ListSubItems.add , , ![MatID] & ""
            
            ITM.ListSubItems.add , , ![MaterialName]
            
            ITM.ListSubItems.add , , Val(![InHand] & "") & " " & ![Unit]
            
            ITM.ListSubItems.add , , Val(![QtyRequired] & "") & " " & ![Unit]
            
            ITM.ListSubItems.add , , Val(![QtyIssued] & "") & " " & ![Unit]
            
            ITM.ListSubItems.add , , Val(![QtyToIssue] & "") & " " & ![Unit]
            
            'Set Tool Tips
            
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ToolTipText = ITM.ListSubItems(i).Text
            Next
            
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
    'Show Lots Refresh
    If LV.ListItems.count > 0 Then Call LV_ItemClick(LV.ListItems(1))
    
Exit Sub
err:
   MsgBox err.Description

    
End Sub


Private Sub cmbMaterials_GotFocus()
    Me.KeyPreview = False
End Sub

Private Sub cmbMaterials_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err
    If KeyCode = 13 And cmbMaterials.MatchFound Then
    
        With LV.SelectedItem
            .key = cmbMaterials.ID
            .SubItems(1) = cmbMaterials.ID
            .SubItems(2) = cmbMaterials.Text
            
            Dim rs As New ADODB.Recordset
            rs.Open "Select Inhand,Unit From Material Where MaterialID='" & cmbMaterials.ID & "'", con, adOpenForwardOnly, adLockReadOnly
            
            .SubItems(3) = Val(rs![InHand] & "") & " " & rs![Unit]
            .SubItems(4) = Val(.SubItems(4)) & " " & rs![Unit]
            .SubItems(5) = ""
            .Tag = rs![Unit] & ""
        End With
        
        LV.SetFocus
    ElseIf KeyCode = 27 Then
        LV.SetFocus
    End If
Exit Sub
err:
    MsgBox err.Description, vbInformation
End Sub

Private Sub cmbMaterials_LostFocus()
    Me.KeyPreview = True
    cmbMaterials.Visible = False
End Sub

Private Sub cmbOrderNo_Click()
    cmbItem.ClearVals
    cmbItem.AddVals con, "ItemName + '(' + ItemSize +') <' + OrderColor +'>'", "VRefOrderItems", "ItemID", "Where RefORderNo='" & cmbOrderNo & "'"
End Sub

Private Sub cmbProcess_Click()
    If Not cmbProcess.MatchFound Then Exit Sub
    txtIncharge = cmbProcess.List(cmbProcess.ListIndex, 2)
    Call AddToCombo(cmbOrderNo, "Distinct RefORderNo", "vRefOrderList", "Order By RefOrderNo")
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    If Saved Then
        Unload Me
        ShowForms.ShowForm Me, frmFlash
    End If
    
End Sub

Private Sub cmdSaveClose_Click()
    
    If Saved Then
        Unload Me
    End If
End Sub

Private Sub cmdUnAssign_Click()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    Dim iAffect As Integer
    
    'con.Execute "Delete From VendAssItems Where EntryID=" & Val(LV.SelectedItem.Key), iAffect
    'If iAffect > 0 Then
    '    'Call cmbNick_Click
    '    cmdUnAssign.Visible = False
    'End If
    
    
    LV.ListItems.Remove LV.SelectedItem.Index
    cmbVend.Enabled = LV.ListItems.count = 0
    Call GetTotal
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub Form_Unload(cancel As Integer)
    HideMainForm
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub


    If Not LV.SelectedItem.Checked Then
        LV.SelectedItem.Checked = True
        Exit Sub
    End If

    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Width = Val(LV.ColumnHeaders(i).Tag)
    Next
    
        With txtRate
            .Move LV.Left + LV.ColumnHeaders(8).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(8).Width
            .Text = LV.SelectedItem.SubItems(7)
            .Visible = True
            .Font.Size = 8
            .SetFocus
        End With
    

    

End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    

    
    cmbItem.ListHeight = 2600
    
    Call AddToCombo(cmbProcess, "Description,ProcessID,InchargeName", "VProcessIncharges", " Where ProcessCat=0 Order By SNo")

    
    DTPicker1 = GetServerDate
    
    txtRecNo = GetNextIssueID(DTPicker1)
    
    
    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
    Next
    
End Sub

Private Sub Form_Resize()

 On Error Resume Next
 fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
 Dim TitleBottom As Long

 TitleBottom = cmbProcess.Top + cmbProcess.Height
 
 LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + Frame6.Height + 100)
 
 Frame6.Move Me.ScaleWidth - Frame6.Width - 50, Me.ScaleHeight - Frame6.Height - 50

 chkPrint.Move LV.Left, Frame6.Top + 50
 chkOrderNo.Move LV.Left, Frame6.Top + 350
 
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)



    'With cmdUnAssign
    '    .Move LV.Left + LV.ColumnHeaders(LV.ColumnHeaders.Count).Width + LV.ColumnHeaders(LV.ColumnHeaders.Count).Left + 10, Item.Top + LV.Top
    '    .Visible = True
    'End With
    'txtRate.Visible = False
End Sub


Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    EditMaterial = CBool(x >= LV.ColumnHeaders(2).Left And _
                x <= (LV.ColumnHeaders(3).Left + LV.ColumnHeaders(3).Width)) And _
                LCase(cmbProcess.Text) = "matching"
End Sub

Private Sub txtRate_GotFocus()
    txtRate.SelStart = 0
    txtRate.SelLength = Len(txtRate.Text)
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        
            With LV.SelectedItem
            
                If Val(txtRate) > Val(.SubItems(3)) Then
                    MsgBox "Cannot Issue More Than Quantity Available In Stock.", vbInformation
                    KeyAscii = 0
                    Exit Sub
                End If
                
                .SubItems(7) = Val(txtRate) & " " & LV.SelectedItem.Tag
                
                
            End With
            
            txtRate.Visible = False
          '  Call GetTotal
    ElseIf KeyAscii = 27 Then
        txtRate.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Function Saved() As Boolean
    Saved = False
    On Error GoTo err
    Dim TempIndex As Integer
    TempIndex = 1
    
    For i = 1 To LV.ListItems.count
        If Val(LV.ListItems(i).SubItems(7)) > 0 Then Exit For
    Next
    
    If i = LV.ListItems.count + 1 Then
        MsgBox "Cannot Save, No Items Issued.", vbInformation
        Exit Function
    End If
    
     
    Call StartTrans(con)
    
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.SubItems(7)) > 0 Then
                con.Execute "Insert Into MaterialIssued(DT,ProcessID,RefORderNo,ItemID,MaterialID,Qty) Values('" & _
                Format(DTPicker1, "dd-MMM-yyyy") & "'," & cmbProcess.List(cmbProcess.ListIndex, 1) & ",'" & cmbOrderNo.Text & _
                "','" & cmbItem.ID & "','" & LV.ListItems(i).key & "'," & Val(LV.ListItems(i).SubItems(7)) & ")"
                
                'Update Stock Status
                con.Execute "Update Material Set InHand=Inhand-" & Val(LV.ListItems(i).SubItems(7)) & " Where MaterialID='" & LV.ListItems(i).key & "'"
                
                
            End If
        End With
    Next
    
   
   
   
    con.CommitTrans
    Saved = True
    
    'If chkPrint = vbChecked Then Call PrintMaterialIssueReport(txtRecNo)
    
    Screen.MousePointer = vbDefault
    
    Exit Function
    
err:

    MsgBox err.Description
    con.RollbackTrans
    Screen.MousePointer = vbDefault
    
End Function

Private Sub txtRate_LostFocus()
    txtRate.Visible = False
    txtRate.Font.Size = 6
End Sub

Private Sub PrintMaterialIssueReport(LOTNo As String)
    
    Screen.MousePointer = vbHourglass


    With frmFlash.cr1
        .Connect = ConnectStr
        .ReportFileName = RptPath & "\MakerIssue.rpt"
        .Formulas(0) = "Company='" & company & "'"
        .SelectionFormula = "{VMakerIssItems.RecieptID}='" & LOTNo & "'"
        .Action = 1
        .PageZoomNext
        .Formulas(0) = ""
        .SelectionFormula = ""
    End With



End Sub


Private Sub GetTotal()
'Dim total As Double
'For i = 1 To LV.ListItems.Count
'   total = total + Val(LV.ListItems(i).SubItems(7))
'Next
'lbltotalBalls = total
End Sub



Public Sub ShowForm(ProcessName As String)
    cmbProcess.Text = ProcessName
    If Not cmbProcess.MatchFound Then
        MsgBox "Please Define Incharge for This Process.", vbInformation
        Unload Me
        Exit Sub
    End If
    cmbProcess.Locked = True
    ShowForms.ShowForm Me, frmFlash, , "Issue Material for " & ProcessName
End Sub



