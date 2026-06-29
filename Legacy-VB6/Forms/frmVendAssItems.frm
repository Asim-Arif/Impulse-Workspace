VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmVendAssItems 
   ClientHeight    =   7395
   ClientLeft      =   1110
   ClientTop       =   60
   ClientWidth     =   11880
   ControlBox      =   0   'False
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
   ScaleHeight     =   7395
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbProcess 
      Height          =   285
      Left            =   7275
      TabIndex        =   1
      Top             =   855
      Width           =   3375
      _ExtentX        =   5953
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
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   855
      Width           =   6540
      _ExtentX        =   11536
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
   Begin ComboList.Usercontrol1 cmbItemSearch 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   9930
      _ExtentX        =   17515
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
   Begin ComboList.Usercontrol1 cmbItemName 
      Height          =   285
      Left            =   2415
      TabIndex        =   4
      Top             =   1800
      Width           =   7320
      _ExtentX        =   12912
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
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   345
      TabIndex        =   32
      Top             =   2580
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox PicEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4260
      ScaleHeight     =   285
      ScaleWidth      =   3990
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2700
      Visible         =   0   'False
      Width           =   3990
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   15
         TabIndex        =   18
         Top             =   0
         Width           =   1000
      End
      Begin VB.TextBox txtSnaff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2985
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.TextBox txtStamp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1950
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.TextBox txtPlant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   975
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   1000
      End
   End
   Begin VB.PictureBox PicSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   720
      ScaleHeight     =   4260
      ScaleWidth      =   9915
      TabIndex        =   26
      Top             =   2085
      Visible         =   0   'False
      Width           =   9945
      Begin VB.TextBox txtCommonRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   1395
         TabIndex        =   31
         Top             =   3885
         Width           =   1755
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   9915
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "Hide"
         Height          =   300
         Left            =   8655
         TabIndex        =   28
         Top             =   3915
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Common Rate :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   60
         TabIndex        =   30
         Top             =   3960
         Width           =   1290
      End
      Begin MSForms.ListBox lstSearch 
         Height          =   3525
         Left            =   0
         TabIndex        =   27
         Top             =   360
         Width           =   9915
         BorderStyle     =   1
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "17489;6218"
         MatchEntry      =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.TextBox txtItemCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   25
      Top             =   1800
      Width           =   1680
   End
   Begin VB.ComboBox cmbUnit 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9105
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   3690
      Visible         =   0   'False
      Width           =   1815
   End
   Begin ComboList.Usercontrol1 cmbItem 
      Height          =   285
      Left            =   5340
      TabIndex        =   3
      Top             =   3930
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
   Begin VB.Frame fupdate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   390
      Left            =   120
      TabIndex        =   15
      Top             =   6465
      Width           =   5250
      Begin VB.CheckBox chkLoadCombo 
         Caption         =   "Load Combo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   24
         Top             =   30
         Width           =   1470
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   375
         Left            =   15
         TabIndex        =   23
         Top             =   0
         Width           =   1500
         ForeColor       =   0
         Caption         =   "Print"
         PicturePosition =   327683
         Size            =   "2646;661"
         Accelerator     =   80
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
         Left            =   15
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
         ForeColor       =   0
         Caption         =   " Assign All"
         PicturePosition =   327683
         Size            =   "2646;661"
         Accelerator     =   104
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
         Left            =   1560
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
         ForeColor       =   0
         Caption         =   "Unassign All"
         PicturePosition =   327683
         Size            =   "2646;661"
         Accelerator     =   108
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.CommandButton cmdUnAssign 
      Height          =   240
      Left            =   9585
      Picture         =   "frmVendAssItems.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "UnAssign Selected Item..."
      Top             =   2310
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   0
      TabIndex        =   10
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
         Left            =   3840
         TabIndex        =   11
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
         Left            =   3855
         TabIndex        =   12
         Top             =   135
         Width           =   2295
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4290
      Left            =   180
      TabIndex        =   9
      Top             =   2160
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   7567
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Group"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Unit"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "D. Rate"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Remarks"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Maker Description"
         Object.Width           =   4233
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   2
      Left            =   720
      TabIndex        =   33
      Top             =   1110
      Width           =   9930
      BackColor       =   11517387
      Caption         =   " Item Search"
      Size            =   "17515;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   1
      Left            =   720
      TabIndex        =   16
      Top             =   1575
      Width           =   9015
      BackColor       =   11517387
      Caption         =   " Code                           Item"
      Size            =   "15901;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdAssign 
      Height          =   525
      Left            =   9765
      TabIndex        =   5
      Top             =   1575
      Width           =   900
      ForeColor       =   6244673
      Caption         =   "Assign   "
      PicturePosition =   327683
      Size            =   "1587;926"
      Accelerator     =   65
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   0
      Left            =   720
      TabIndex        =   13
      Top             =   630
      Width           =   9930
      BackColor       =   11517387
      Caption         =   $"frmVendAssItems.frx":0102
      Size            =   "17515;423"
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
      TabIndex        =   8
      Top             =   6510
      Width           =   1665
      ForeColor       =   0
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "2937;661"
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
Dim iColNo As Integer
Dim bCanChangeRate As Boolean, bHideRate As Boolean

Private Sub chkLoadCombo_Click()

    If chkLoadCombo.Value = vbChecked Then
        If cmbCust.MatchFound Then
            '
            Call AddInCombo
        End If
    End If
    
End Sub

Private Sub AddInCombo()
    Call cmbItemName.AddVals(con, "'{' + ItemID + '} '  + IsNull(ItemName,'') + ' ' + IsNull(ItemSize,'') + ' ' + IsNull(SizeUnit,'') + ' ' + IsNull(FinQuality,'') ", "VItems", "ItemID", " WHERE ItemID IN(SELECT ItemID FROM ItemProcesses WHERE ProcessID=" & cmbProcess.ID & ") AND ItemID NOT IN(SELECT ItemID FROM VendAssItems WHERE VendID=" & cmbCust.ID & " AND ProcessID=" & cmbProcess.ID & ")  ORDER BY ItemName")
End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    
    cmbProcess.ClearVals
    Call cmbProcess.AddVals(con, "Description", "VMakerAssignedProcesses", "ProcessID", " WHERE VendID=" & cmbCust.ID)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillItems()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim strRate As String, strRate_Dozen As String
    
    With rs
        .Open "Select EntryID,ProcessID,ItemID,ItemName,Description,Rate,PlantRate,SnaffRate,StampRate,PhaseID,Unit,AssignedUnit,Remarks,MakerDescription FROM VMakerAssItems WHERE VendID=" & cmbCust.ID & " AND ProcessID=" & cmbProcess.ID, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![EntryID] & "'", ![ITemID])
            ITM.Tag = Val(!ProcessID & "")
            ITM.ListSubItems.add , , ![ItemName] & ""
            ITM.ListSubItems.add , , "-" '![Type] & ""
            ITM.ListSubItems.add , , ![Description] & ""
            If bHideRate Then
                strRate = "-"
                strRate_Dozen = "-"
            Else
                strRate = Val(!Rate & "")
                strRate_Dozen = Val(!Rate & "") * 12
            End If
            ITM.ListSubItems.add , , strRate
            ITM.ListSubItems.add(, , !AssignedUnit & "").Tag = !Unit & ""
            ITM.ListSubItems.add(, , strRate_Dozen).Tag = ""
            
            ITM.ListSubItems.add , , !Remarks & "" & ""
            ITM.ListSubItems.add , , !MakerDescription & ""
'            If cmbCust.Tag = 7 Then
'                ITM.SubItems(5) = ![PlantRate] & ""
'                ITM.SubItems(6) = ![StampRate] & ""
'                ITM.SubItems(7) = ![SnaffRate] & ""
'                ITM.SubItems(8) = Val(ITM.SubItems(7) & "") + Val(ITM.SubItems(6) & "") + Val(ITM.SubItems(5) & "") + Val(ITM.SubItems(4) & "")
'            End If
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    cmbItemSearch.ClearVals
    Call cmbItemSearch.AddVals(con, "'{' + ItemID + '} '  + IsNull(ItemName,'')", "VMakerAssItems", "EntryID", " WHERE VendID=" & cmbCust.ID & " AND ProcessID=" & cmbProcess.ID)
    
End Sub
Private Sub cmbCust_UnMatched()
    LV.ListItems.Clear
    'cmbCountry.ClearVals
End Sub

Private Sub cmbItem_matched()
    cmbItemName.ID = cmbItem.ID
End Sub

Private Sub cmbItemName_matched()
    cmbItem.ID = cmbItemName.ID
    txtItemCode.Text = cmbItemName.ID
End Sub

Private Sub cmbItemSearch_matched()
    LV.ListItems(cmbItemSearch.ID & "'").Selected = True
    LV.SelectedItem.EnsureVisible
End Sub

Private Sub cmbProcess_matched()

    cmbItemName.ClearVals
    
    If chkLoadCombo.Value = vbChecked Then
        Call cmbItemName.AddVals(con, "'{' + ItemID + '} '  + IsNull(ItemName,'') + ' ' + IsNull(ItemSize,'') + ' ' + IsNull(SizeUnit,'') + ' ' + IsNull(Type1,'') + ' ' + IsNull(FinQuality,'') ", "VItems", "ItemID", " Where ItemID In(Select ItemID From ItemProcesses Where ProcessID=" & cmbProcess.ID & ") and ItemID Not In(Select ItemID From VendAssItems Where VendID=" & cmbCust.ID & " AND ProcessID=" & cmbProcess.ID & ")  Order By ItemName")
    End If
            
    cmbCust.Tag = cmbProcess.ID
    
    Call FillItems
    
End Sub

Private Sub cmbUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        con.Execute "UPDATE VendAssItems SET Unit='" & cmbUnit.Text & "' Where EntryID=" & Val(LV.SelectedItem.key)
        LV.SelectedItem.SubItems(5) = cmbUnit
        cmbUnit.Visible = False
    ElseIf KeyAscii = 27 Then
        cmbUnit.Visible = False
    End If
End Sub

Private Sub AssignItem(strItemCode As String, Optional dRate As Double = 0)

    If cmbCust.MatchFound = False Then Exit Sub
    
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Text = strItemCode Then
            MsgBox "Already Assigned !!!", vbInformation
            Exit Sub
        End If
    Next
    Dim Saved As Integer
    Dim cmd As New ADODB.Command
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_VendAssItem"
        .Parameters("@VendID").Value = cmbCust.ID
        .Parameters("@ItemID").Value = strItemCode 'cmbItemName.ID 'cmbNick.List(cmbNick.ListIndex, 1)
        .Parameters("@ProcessID").Value = cmbCust.Tag
        .Parameters("@Rate").Value = dRate 'IIf(dRate = 0, 1, dRate)
        .Execute Saved
    End With
    
    If Saved = 1 Then
        Call FillItems
    End If
    
End Sub

Private Sub cmdAssign_Click()

    On Error GoTo err
    If cmbItemName.MatchFound = False Then Exit Sub
    If txtItemCode.Text = "" Then Exit Sub
    
    
    If cmbCust.MatchFound = False Then Exit Sub
    Dim strItemCode As String
    If cmbItemName.MatchFound Then
        strItemCode = cmbItemName.ID
    Else
        strItemCode = txtItemCode
    End If
    
    Call AssignItem(strItemCode)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub



Private Sub cmdCheckAll_Click()
        
    Exit Sub
     
    If cmbCust.MatchFound = False Then Exit Sub
    
    'If cmbItem.MatchFound = False Then Exit Sub
    
    Dim Saved As Integer
    Dim cmd As New ADODB.Command
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_VendAssItem"
        .Parameters("@VendID").Value = cmbCust.ID 'cmbCust.ID
        .Parameters("@ItemID").Value = "0" 'cmbCountry.ID
        .Parameters("@ProcessID").Value = cmbCust.Tag
        .Execute Saved
    End With
    
    
    If Saved > 0 Then
        'Call RefreshLV
        'Call CmbCust's Matched Event To Remove Current Item
        'From The cmbItem...
        Call cmbCust_matched
        'cmdAssign.SetFocus
    End If

End Sub

Private Sub cmdClearAll_Click()

    Exit Sub
    On Error GoTo err
    Dim iAffect As Integer
    If MsgBox("Are You Sure To UnAssign All Items?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        con.Execute "Delete From VendAssItems Where VendID=" & cmbCust.ID, iAffect
        If iAffect > 0 Then
            Call cmbCust_matched
        End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdHide_Click()
    PicSearch.Visible = False
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo err
    If bHideRate Then Exit Sub
    
    If cmbCust.MatchFound = False Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\MakerAssignedItems.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VendAssItems.VendID}=" & cmbCust.ID, rpt
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdUnAssign_Click()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT ItemCode FROM VVendIssdDetail WHERE VendID=" & cmbCust.ID & " AND IssQty>RcvdQty AND ProcessID=" & cmbProcess.ID & " AND ItemCode='" & LV.SelectedItem.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            MsgBox "Can't Unassign this item. It's Receivable.", vbInformation
            .Close
            Exit Sub
        End If
        .Close
    End With
    Set rs = Nothing
    
    Dim iAffect As Integer
    
    con.Execute "DELETE FROM VendAssItems WHERE EntryID=" & Val(LV.SelectedItem.key), iAffect
    If iAffect > 0 Then
        Call cmbProcess_matched
        cmdUnAssign.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        SendKeys "{Tab}"
    End If
End Sub

Private Sub lstSearch_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    If KeyCode.Value = 13 Then
        If lstSearch.ListIndex > -1 Then
            Call AssignItem(lstSearch.List(lstSearch.ListIndex, 1), Val(txtCommonRate))
        End If
    ElseIf KeyCode.Value = vbKeyBack And Len(txtSearch) > 0 Then
        txtSearch = Left(txtSearch, Len(txtSearch) - 1)
        txtSearch.SetFocus
    ElseIf KeyCode.Value = vbKeyEscape Then
        PicSearch.Visible = False
        txtItemCode.SetFocus
    End If
    
End Sub


Private Sub lstSearch_KeyPress(KeyAscii As MSForms.ReturnInteger)
'    MsgBox KeyAscii
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    LV.Sorted = True
    If LV.SortKey = ColumnHeader.Index - 1 Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Up"
        Else
            LV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Down"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        'Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If

    On Error Resume Next
    'LV.SelectedItem.EnsureVisible

End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    If cmbCust.MatchFound = False Then Exit Sub
    If bCanChangeRate = False Then
        MsgBox "You don't have the Access to Change Rate.", vbInformation
        Exit Sub
    End If
    If bHideRate Then
        Exit Sub
    End If
    
    If Not LV.SelectedItem.Checked Then
        LV.SelectedItem.Checked = True
        'Call LV_ItemCheck(LV.SelectedItem)
        Exit Sub
    End If
    
    If iColNo = 6 Then
        With cmbUnit
            .Clear
            If LV.SelectedItem.ListSubItems(5).Tag & "" <> "" Then .AddItem LV.SelectedItem.SubItems(5)
            If LV.SelectedItem.ListSubItems(5) & "" <> "" And LV.SelectedItem.ListSubItems(5) <> LV.SelectedItem.ListSubItems(5).Tag Then .AddItem LV.SelectedItem.SubItems(5)
            If LV.SelectedItem.ListSubItems(5) <> "Kgs" Then
                cmbUnit.AddItem "Kgs"
            End If
            .Move LV.Left + LV.ColumnHeaders(6).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(6).Width
            .Text = LV.SelectedItem.ListSubItems(5)
            .Visible = True
            .SetFocus
        End With
        Exit Sub
    End If
           
'    With txtRate
'        PicEdit.Move LV.Left + LV.ColumnHeaders(5).Left, LV.Top + LV.SelectedItem.Top ', LV.ColumnHeaders(5).Width
'        If LV.ColumnHeaders.count = 9 Then
'            Dim rs As New ADODB.Recordset
'            With rs
'                .Open "Select IsNull(Planter,0),IsNull(Snaffer,0),IsNull(Stamper,0) From Venders Where VendID=" & cmbCust.ID, con, adOpenForwardOnly, adLockReadOnly
'
'                txtPlant.Enabled = False
'                txtSnaff.Enabled = False
'                txtStamp.Enabled = False
'                .Close
'
'            End With
'
'            PicEdit.Width = 3990
'            txtPlant.Visible = True
'            txtSnaff.Visible = True
'            txtStamp.Visible = True
'            txtPlant.Text = LV.SelectedItem.SubItems(5)
'            txtStamp.Text = LV.SelectedItem.SubItems(6)
'            txtSnaff.Text = LV.SelectedItem.SubItems(7)
'        Else
'            PicEdit.Width = .Width
'            txtPlant.Visible = False
'            txtSnaff.Visible = False
'            txtStamp.Visible = False
'        End If
'        PicEdit.Visible = True
'        .Text = LV.SelectedItem.SubItems(4)
'        .Visible = True
'        .SelStart = 0
'        .SelLength = Len(.Text)
'        .SetFocus
'    End With
    
    If iColNo <> 8 Then iColNo = 5
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
        .Text = LV.SelectedItem.SubItems(iColNo - 1)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With

End Sub





Private Sub Form_Load()

    cmbCust.ListHeight = 4000
    
    cmbCust.ClearVals
    Call cmbCust.AddVals(con, "'{ ' + VendID1 + ' } ' + MakerName", "VMakers", "VendID")
    
    'cmbNick.ListHeight = 2600
    cmbItem.ListHeight = 2600
    cmbItemName.ListHeight = 2600
    cmbProcess.ListHeight = 2600
    
    'cmbProcess.AddItem "REPAIRERS", "0"
    
    iColNo = 0
    
    
    If CurrentUserName <> "Administrator" Then
        bCanChangeRate = GetSingleBooleanValue("ChangeRateMakerAssign", "Users", " WHERE UserID=" & UserID)
    Else
        bCanChangeRate = True
    End If
        
    If CurrentUserName <> "Administrator" Then
        bHideRate = GetSingleBooleanValue("HideRateMakerAssign", "Users", " WHERE UserID=" & UserID)
    Else
        bHideRate = False
    End If
    
End Sub

Public Sub RefreshLV()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select *  from FCustomerCatalog Where CustCode='" & cmbCust.ID & "' and Country='" & cmbCountry.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![CustCode] & ![Country] & ![ITemID], ![ITemID] & "")
            ITM.SubItems(1) = ![Description] & ""
            ITM.SubItems(2) = ![Unit] & ""
            ITM.SubItems(3) = ![FOB] & ""
            ITM.SubItems(4) = ![CIFSea] & ""
            ITM.SubItems(5) = ![CIFAir] & ""
            ITM.SubItems(6) = ![CnFSea] & ""
            ITM.SubItems(7) = ![CnFAir] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
End Sub
Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
    Dim TitleBottom As Long

    TitleBottom = cmbItemName.Top + cmbItemName.Height 'fTitle.Top + fTitle.Height
 
    LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100)
 
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
    fupdate.Move LV.Left, LV.Top + LV.Height + 50
    
    'cmdCheckAll.Move LV.Left, LV.Top + LV.Height + 50
 
 
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With cmdUnAssign
        .Move LV.Left + LV.ColumnHeaders(LV.ColumnHeaders.count).Width + LV.ColumnHeaders(LV.ColumnHeaders.count).Left + 10, Item.Top + LV.Top
        .Visible = True
    End With
    PicEdit.Visible = False
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 1 Then Exit Sub
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If X > LV.ColumnHeaders(i).Left And X < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub



Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 2 Then Exit Sub
    If bHideRate Then Exit Sub
    
    Dim strRet As String
    strRet = ShowPopUpMenu1("Show History")
        
    If strRet = "Show History" Then
        Dim f As New frmMakerAssItemHistory
        Load f
        f.ViewHistory (Val(LV.SelectedItem.key))
    End If
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        If iColNo = 5 Then
            Dim iAffect As Integer
            con.Execute "UPDATE VendAssItems SET Rate=" & Val(txtEdit) & " WHERE EntryID=" & Val(LV.SelectedItem.key), iAffect
            If iAffect = 1 Then
                con.Execute "INSERT INTO VendAssItems_Revisions(VAI_RefID,Rate,UserName,Remarks) VALUES(" & Val(LV.SelectedItem.key) & "," & Val(LV.SelectedItem.SubItems(4)) & ",'" & CurrentUserName & "','" & LV.SelectedItem.SubItems(7) & "')"
                LV.SelectedItem.SubItems(4) = Val(txtEdit)
                LV.SelectedItem.SubItems(6) = Val(txtEdit) * 12
            End If
        Else
            con.Execute "UPDATE VendAssItems SET Remarks='" & txtEdit & "' WHERE EntryID=" & Val(LV.SelectedItem.key)
            LV.SelectedItem.SubItems(iColNo - 1) = txtEdit
        End If
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        'txtedit.Visible = False
        txtEdit.Visible = False
    End If
    'Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub txtItemCode_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Dim rs As New ADODB.Recordset
        Dim bFound As Boolean
        If cmbProcess.MatchFound = False Then
            MsgBox "Please Select Process.", vbInformation
            Exit Sub
        End If
        With rs
            .Open "SELECT ItemID FROM Items WHERE ItemID='" & txtItemCode & "' AND ItemID IN(SELECT ItemID FROM ItemProcesses WHERE ProcessID=" & cmbProcess.ID & ") AND ItemID NOT IN(SELECT ItemID FROM VendAssItems WHERE VendID=" & cmbCust.ID & " AND ProcessID=" & cmbProcess.ID & ")  ORDER BY ItemID", con, adOpenForwardOnly, adLockReadOnly
            If .EOF Then
                bFound = False
            Else
                Call AssignItem(txtItemCode)
                txtItemCode.SetFocus
                bFound = True
            End If
            .Close
            
            If bFound = False Then
                txtSearch.Text = txtItemCode
                txtSearch.SelStart = Len(txtSearch.Text) + 1
                PicSearch.Visible = True
                txtSearch.SetFocus
            End If
            Set rs = Nothing
        End With
    End If
    
End Sub

Private Sub txtPlant_KeyPress(KeyAscii As Integer)
    On Error GoTo err
    
    If KeyAscii = 13 Then
        'txtStamp.SetFocus
        If Not (txtSnaff.Enabled Or txtStamp.Enabled) Then
            Call txtSnaff_KeyPress(13)
        End If
    ElseIf KeyAscii = 27 Then
        'txtRate.Visible = False
        PicEdit.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        If iColNo = 4 Then
            Dim iAffect As Integer
            con.Execute "UPDATE VendAssItems SET Rate=" & Val(txtRate) & " WHERE EntryID=" & Val(LV.SelectedItem.key), iAffect
            If iAffect = 1 Then
                con.Execute "INSERT INTO VendAssItems_Revisions(VAI_RefID,Rate,UserName) VALUES(" & Val(LV.SelectedItem.key) & "," & Val(txtRate) & ",'" & CurrentUserName & "')"
                LV.SelectedItem.SubItems(4) = Val(txtRate)
                LV.SelectedItem.SubItems(6) = Val(txtRate) * 12
                PicEdit.Visible = False
            End If
        Else
            con.Execute "UPDATE VendAssItems SET Remarks='" & txtRate & "' WHERE EntryID=" & Val(LV.SelectedItem.key)
        End If
    ElseIf KeyAscii = 27 Then
        'txtRate.Visible = False
        PicEdit.Visible = False
    End If
    'Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtSearch_Change()

    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT '{' + ItemID + '} '  + ItemName + ' ' + ItemSize + SizeUnit,ItemID FROM VItems WHERE ItemID LIKE '%" & txtSearch & "%' AND ItemID IN(SELECT ItemID FROM ItemProcesses WHERE ProcessID=" & cmbProcess.ID & ") AND ItemID NOT IN(SELECT ItemID FROM VendAssItems WHERE VendID=" & cmbCust.ID & " AND ProcessID=" & cmbProcess.ID & ")", con, adOpenForwardOnly, adLockReadOnly
        lstSearch.Clear
        Do Until .EOF
            lstSearch.AddItem .Fields(0).Value & ""
            lstSearch.List(lstSearch.ListCount - 1, 1) = !ITemID & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then
        If lstSearch.ListCount > 0 Then
            If lstSearch.ListIndex < (lstSearch.ListCount - 1) Then
                lstSearch.ListIndex = lstSearch.ListIndex + 1
            End If
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyUp Then
        If lstSearch.ListCount > 0 Then
            If lstSearch.ListIndex > 0 Then
                lstSearch.ListIndex = lstSearch.ListIndex - 1
            End If
            KeyCode = 0
        End If
    End If
    
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If lstSearch.ListIndex > -1 Then
            Call AssignItem(lstSearch.List(lstSearch.ListIndex, 1))
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        PicSearch.Visible = False
        txtItemCode.SetFocus
    End If
End Sub

Private Sub txtSnaff_KeyPress(KeyAscii As Integer)
    On Error GoTo err
    
    If KeyAscii = 13 Then
        Dim iAffect As Integer
        con.Execute "Update VendAssItems Set Rate=" & Val(txtRate) & ",PlantRate=" & Val(txtPlant) & ",StampRate=" & Val(txtStamp) & ",SnaffRate=" & Val(txtSnaff) & "  Where EntryID=" & Val(LV.SelectedItem.key), iAffect
        If iAffect = 1 Then
            LV.SelectedItem.SubItems(4) = Val(txtRate)
            LV.SelectedItem.SubItems(5) = Val(txtPlant)
            LV.SelectedItem.SubItems(6) = Val(txtStamp)
            LV.SelectedItem.SubItems(7) = Val(txtSnaff)
            LV.SelectedItem.SubItems(8) = Val(txtSnaff) + Val(txtStamp) + Val(txtPlant) + Val(txtRate)
            PicEdit.Visible = False
        End If
    ElseIf KeyAscii = 27 Then
        'txtRate.Visible = False
        PicEdit.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtStamp_KeyPress(KeyAscii As Integer)
    On Error GoTo err
    
    If KeyAscii = 13 Then
        'txtSnaff.SetFocus
        If Not (txtSnaff.Enabled) Then
            Call txtSnaff_KeyPress(13)
        End If
    ElseIf KeyAscii = 27 Then
        'txtRate.Visible = False
        PicEdit.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
