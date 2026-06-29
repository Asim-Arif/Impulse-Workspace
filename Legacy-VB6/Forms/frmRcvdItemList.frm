VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmRcvdItemList 
   ClientHeight    =   7695
   ClientLeft      =   75
   ClientTop       =   -30
   ClientWidth     =   12780
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleMode       =   0  'User
   ScaleWidth      =   11447.22
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rcvd Entries"
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
         Index           =   8
         Left            =   75
         TabIndex        =   2
         Top             =   165
         Width           =   10425
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Rcvd Entries"
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
         Height          =   465
         Index           =   9
         Left            =   90
         TabIndex        =   3
         Top             =   180
         Width           =   10410
      End
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   7050
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   12720
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
         ItemData        =   "frmRcvdItemList.frx":0000
         Left            =   75
         List            =   "frmRcvdItemList.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   375
         Width           =   2100
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5820
         Left            =   60
         TabIndex        =   4
         Top             =   720
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   10266
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Iss No."
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Order No"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Rcvd By"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Rcvd From"
            Object.Width           =   6174
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   4155
         TabIndex        =   7
         Top             =   375
         Width           =   1920
         _ExtentX        =   3387
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
         Format          =   50724867
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   2205
         TabIndex        =   8
         Top             =   375
         Width           =   1920
         _ExtentX        =   3387
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
         Format          =   50724867
         CurrentDate     =   37055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   " Range                                   From                                 To"
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
         TabIndex        =   10
         Top             =   150
         Width           =   6000
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   345
         Left            =   6120
         TabIndex        =   9
         Top             =   360
         Width           =   1545
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2725;609"
         Accelerator     =   82
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
         Left            =   11040
         TabIndex        =   5
         Top             =   6600
         Width           =   1620
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2857;635"
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
Attribute VB_Name = "frmRcvdItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String

Private Sub cmbMaterial_matched()
    Call RefreshLV
End Sub

Private Sub cmbRange_Change()
    Call cmbRange_Click
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = GetServerDate
        Case 1
            DtFrom = DateAdd("d", -15, GetServerDate)
        Case 2
            DtFrom = DateAdd("d", -30, GetServerDate)
        Case 3
            DtFrom = DateAdd("d", -60, GetServerDate)
        Case 4
            DtFrom = DateAdd("d", -90, GetServerDate)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub


Private Sub cmdClose_Click()
   Unload Me
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub

Private Sub Form_Load()
    
'    cmbRange.ListIndex = 0
    Call RefreshLV
    cmbRange.ListIndex = 0
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim DTRange As String
    
    Screen.MousePointer = vbHourglass
    
'    If cmbRange.ListIndex = 0 Or cmbRange.ListIndex = -1 Then
'        DTRange = " DT Between '" & DtFrom & "' and '" & DtTo & "'"
'    ElseIf cmbRange.ListIndex = 1 Then
'        DTRange = " DT Between '" & DateAdd("d", -30, GetServerDate) & "' and '" & GetServerDate & "'"
'    ElseIf cmbRange.ListIndex = 2 Then
'        DTRange = " DT Between '" & DateAdd("d", -60, GetServerDate) & "' and '" & GetServerDate & "'"
'    ElseIf cmbRange.ListIndex = 3 Then
'        DTRange = " DT Between '" & DateAdd("d", -90, GetServerDate) & "' and '" & GetServerDate & "'"
'    ElseIf cmbRange.ListIndex = 4 Then
'        DTRange = " DT Between '" & DtFrom & "' and '" & DtTo & "'"
'    End If
    
    DTRange = " Where Cast(Convert(varchar(8),DT,1) As DateTime) Between '" & DtFrom & "' and '" & DtTo & "'"
    
    With rs
        SQL = "Select * From RcvItemsSimple " & DTRange & " Order By DT,RcvNo"
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        'LV.Sorted = False
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, ![RcvNo] & "'", .AbsolutePosition)
            
            ITM.ListSubItems.add , , !RcvNo
            ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
            
            ITM.ListSubItems.add , , !OrderNo
            ITM.ListSubItems.add , , !RcvdBy
            ITM.ListSubItems.add , , !RcvdFrom
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 2 Then Exit Sub
 
    Dim strRet As String
'    If LV.ListItems.count = 0 Then
'        strRet = ShowPopUpMenu1("&Issue New Material")
'    Else
'        strRet = ShowPopUpMenu1("&Issue New Material", "-", "&Print This Entry", "&Return Stock", "Print &List")
'    End If
    strRet = ShowPopUpMenu1("&Issue New Material")
    Select Case strRet
    
        Case "&Issue New Material"
        
            frmIssItemsSimple.Show
            
        Case "&Print This Entry"
        
'            Dim f As New frmPrevRpt
'            Dim rpt As CRAXDDRT.Report
'            Set rpt = rptApp.OpenReport(rptProductionPath & "\rptIssItemsSimple.rpt")
'            f.ShowReport "{IssItemsSimple.IssNo}='" & LV.SelectedItem.SubItems(1) & "' ", rpt
            
        Case "&Return Stock"
        
'            Dim frm As New frmReturnIssdItems
'            frm.ShowDetail (LV.SelectedItem.SubItems(1))
        
        Case "Print &List"
        
'            Set rpt = rptApp.OpenReport(rptProductionPath & "\rptMaterialIssList.rpt")
'            rpt.FormulaFields(1).Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo, "dd-MMM-yyyy") & "'"
            
'            f.ShowReport "{IssItemsSimple.DT}=#" & Format(DtFrom.value, "dd-MMM-yyyy") & "# To #" & Format(DtTo.value, "dd-MMM-yyyy") & "#", rpt
            
    End Select
        
End Sub


