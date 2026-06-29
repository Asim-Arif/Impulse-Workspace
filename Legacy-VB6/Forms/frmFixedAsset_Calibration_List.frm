VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmFixedAsset_Calibration_List 
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1200
      Top             =   1425
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
            Picture         =   "frmFixedAsset_Calibration_List.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFixedAsset_Calibration_List.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   8400
      Left            =   60
      TabIndex        =   3
      Top             =   -45
      Width           =   11865
      Begin MSComDlg.CommonDialog CD1 
         Left            =   750
         Top             =   1365
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Please Select Picture"
         Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
      End
      Begin VB.ComboBox cmbRange 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmFixedAsset_Calibration_List.frx":0378
         Left            =   4440
         List            =   "frmFixedAsset_Calibration_List.frx":0391
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   345
         Width           =   2010
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   8190
         TabIndex        =   1
         Top             =   345
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   113770499
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6465
         TabIndex        =   0
         Top             =   345
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   113770499
         CurrentDate     =   37055
      End
      Begin ComboList.Usercontrol1 cmbAsset 
         Height          =   285
         Left            =   60
         TabIndex        =   6
         Top             =   345
         Width           =   4365
         _ExtentX        =   7699
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
      Begin MSComctlLib.ListView LV 
         Height          =   7110
         Left            =   60
         TabIndex        =   7
         Top             =   750
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   12541
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Asset"
            Object.Width           =   5609
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Calibration Date"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Cert. #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Remarks"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Next Due Date"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "PDF"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10455
         TabIndex        =   8
         Top             =   7935
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   9810
         TabIndex        =   2
         Top             =   315
         Width           =   1500
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2646;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmFixedAsset_Calibration_List.frx":03F0
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
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   120
         Width           =   9645
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add New Entry"
      End
      Begin VB.Menu mnuPrintSlip 
         Caption         =   "Print Slip"
         Visible         =   0   'False
      End
      Begin VB.Menu dd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu d1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintList 
         Caption         =   "Print List"
      End
   End
End
Attribute VB_Name = "frmFixedAsset_Calibration_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbRange_Change()
    Call cmbRange_Click
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = Date
        Case 1
            DtFrom = DateAdd("d", -15, Date)
        Case 2
            DtFrom = DateAdd("d", -30, Date)
        Case 3
            DtFrom = DateAdd("d", -60, Date)
        Case 4
            DtFrom = DateAdd("d", -90, Date)
        Case 5
            DtFrom = DateAdd("d", -180, Date)
        Case 6
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()

    On Error GoTo err
    Dim rs As New ADODB.Recordset, ITM As ListItem
    Dim strCondition As String
    
    strCondition = " WHERE DT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'"
    
    If cmbAsset.ID <> "0" Then
        strCondition = strCondition & " AND AssetID='" & cmbAsset.ID & "'"
    End If
    
    With rs
        .Open "SELECT * FROM VFixedAssets_Calibration " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add(, , "(" & !AssetID & ") " & !Description & "").Tag = !AssetID & ""
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , !CertificateNo & ""
            ITM.ListSubItems.add , , !Remarks & ""
            ITM.ListSubItems.add , , Format(!NextDueDT, "dd-MMM-yyyy")
            
            If IsNull(!PDF) Then
                ITM.ListSubItems.add , , "-"
            Else
                ITM.ListSubItems.add , , "Attached"
            End If
            
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Load()

    cmbRange.ListIndex = 0
    
    cmbAsset.ClearVals
    cmbAsset.AddItem "<All Assets>", "0"
    cmbAsset.ListHeight = 4000
    cmbAsset.AddVals con, "'{' + AssetID + '} ' + Description", "FixedAssets", "AssetID", " ORDER BY AssetID"
    cmbAsset.ID = "0"
    
End Sub

Private Sub LV_DblClick()
    
    If LV.SelectedItem.SubItems(6) = "Attached" Then
        Call ShowTechDrawing("FixedAssets_Calibration", "PDF", " WHERE EntryID=" & Val(LV.SelectedItem.Tag), LV.SelectedItem.Text, CD1)
    Else
        MsgBox "No PDF attached.", vbInformation
        Exit Sub
    End If
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    Me.PopUpMenu mnupop
End Sub

Private Sub mnuAdd_Click()
    Dim f As New frmFixedAssets_Calibration_AE
    Load f
    f.AddNew
End Sub

Private Sub mnuEdit_Click()
    Dim f As New frmFixedAssets_Calibration_AE
    Load f
    Call f.EditEntry(Val(LV.SelectedItem.Tag))
End Sub

Private Sub mnuPrintList_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\FixedAssets_CalibrationList.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{FixedAssets_Calibration.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#", rpt
    
End Sub

Private Sub mnuPrintSlip_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\FixedAssets_RepairSlip.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{FixedAssets_RepairLog.EntryID}=" & Val(LV.SelectedItem.Tag), rpt
    
End Sub
