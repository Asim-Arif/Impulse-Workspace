VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmStockStatus 
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleMode       =   0  'User
   ScaleWidth      =   11675.62
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbItemName 
      Height          =   285
      Left            =   3150
      TabIndex        =   9
      Top             =   420
      Width           =   5070
      _ExtentX        =   8943
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
   End
   Begin ComboList.Usercontrol1 cmbItem 
      Height          =   285
      Left            =   870
      TabIndex        =   10
      Top             =   420
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
   End
   Begin MSComCtl2.DTPicker DtTo 
      Height          =   300
      Left            =   5670
      TabIndex        =   2
      Top             =   6375
      Visible         =   0   'False
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   55312387
      CurrentDate     =   37081
   End
   Begin MSComCtl2.DTPicker DtFrom 
      Height          =   300
      Left            =   4290
      TabIndex        =   3
      Top             =   6375
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   55312387
      CurrentDate     =   37055
   End
   Begin VB.Frame FLV 
      Height          =   6045
      Left            =   30
      TabIndex        =   0
      Top             =   780
      Width           =   12615
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   465
         Left            =   8295
         TabIndex        =   6
         Top             =   5475
         Width           =   3030
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1530
            TabIndex        =   8
            Top             =   75
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
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   60
            TabIndex        =   7
            Top             =   75
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
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5265
         Left            =   105
         TabIndex        =   1
         Top             =   195
         Width           =   12390
         _ExtentX        =   21855
         _ExtentY        =   9287
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15199215
         BorderStyle     =   1
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
            Text            =   "Item Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Description"
            Object.Width           =   5997
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Process"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "In Stock"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Out Stock"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Wastage"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Lost"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockStatus.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockStatus.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   885
      TabIndex        =   5
      Top             =   165
      Width           =   7305
      BackColor       =   11517387
      Caption         =   $"frmStockStatus.frx":0564
      Size            =   "12885;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   8265
      TabIndex        =   4
      Top             =   375
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
End
Attribute VB_Name = "frmStockStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrintSQL As String
Dim ProcessID As Long
Dim FreeStock As Boolean
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    Screen.MousePointer = vbHourglass
    
    Dim rpt As craxddrt.Report 'New rptStockStatus
    If ProcessID = 1 Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptStockStatusForging.rpt") '
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptStockStatus.rpt") '
    End If
    
    rpt.SQLQueryString = PrintSQL
    
    Dim F As New frmPrevRpt
    Load F
    
    F.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
err:
    MsgBox err.Description
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub cmdRefresh_Click()

    Call RefreshLV

End Sub

Private Sub Form_Load()

    'DtTo = ServerDate
    'DtFrom = DateAdd("M", -1, ServerDate)
    
    cmbItem.ListHeight = 2200
    cmbItemName.ListHeight = 2200
    
    
End Sub
Private Sub Form_Resize()

    On Error Resume Next
     
    FLV.Move 50, 800, Me.ScaleWidth - 100, Me.ScaleHeight - 900
     
    LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
     
    fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rsLedger As New ADODB.Recordset
    
    Dim Cond As String
    
    Cond = " Where ProcessID=" & ProcessID
    
    If cmbItem.ID = "0" Then
        
    Else
        Cond = Cond & " and ItemID='" & cmbItem.ID & "'"
    End If
    
    If FreeStock Then
        If ProcessID = 0 Then
            PrintSQL = "Select CompItemCode,ItemDescription,ProcessID" & _
                ",Description,Sum(InStock) As InStock,Sum(OutStock) As OutStock," & _
                "Sum(TotWastage) As TotWastage,Sum(TotLost) As TotLost From VStockStatusRpt " & _
                IIf(cmbItem.ID = "0", "", " Where ItemID='" & cmbItem.ID & "'") & " Group By CompItemCode," & _
                " ItemDescription,ProcessID,Description"
        Else
            PrintSQL = "Select CompItemCode,ItemDescription,ProcessID" & _
                ",Description,Sum(InStock) As InStock,Sum(OutStock) As OutStock," & _
                "Sum(TotWastage) As TotWastage,Sum(TotLost) As TotLost From VStockStatusRpt " & _
                IIf(cmbItem.ID = "0", "", " AND ItemID='" & cmbItem.ID & "'") & " Group By CompItemCode," & _
                " ItemDescription,ProcessID,Description"
        End If
    Else
        PrintSQL = "Select * From VStockStatusrpt " & Cond
    End If
    
    With rsLedger
        
        .Open PrintSQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Dim ITM As ListItem
        For i = 0 To .RecordCount - 1
            
            Set ITM = LV.ListItems.Add(, , ![CompItemCode] & "")
            ITM.Tag = Val(!ProcessID & "")
            ITM.SubItems(1) = ![ItemDescription] & ""
            ITM.SubItems(2) = ![Description] & ""
            ITM.SubItems(3) = Val(![InStock] & "")
            ITM.SubItems(4) = Val(![OutStock] & "")
            ITM.SubItems(5) = Val(![TotWastage] & "")
            ITM.SubItems(6) = Val(![TotLost] & "")
            .MoveNext
        Next
        .Close
    End With
    LV.Refresh
    Set rsLedger = Nothing
        
    Cond = ""
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbItem_matched()
    cmbItemName.ID = cmbItem.ID
End Sub

Private Sub cmbItemName_matched()
    cmbItem.ID = cmbItemName.ID
    'Call RefreshLV
End Sub
Public Sub ShowStockStatus(lProcessID As Long, Optional bFreeStock As Boolean = False)
    
    On Error GoTo err
    
    FreeStock = bFreeStock
    Me.Hide
    
    ProcessID = lProcessID
    
    cmbItem.AddItem "<All Items>", 0
    cmbItemName.AddItem "<All Items>", 0
    If lProcessID = 0 Then
        Call cmbItem.AddVals(con, "Distinct ItemID", "VStockStatusrpt", "ItemID")
        Call cmbItemName.AddVals(con, "Distinct ItemDescription", "VStockStatusrpt", "ItemID")
    Else
        Call cmbItem.AddVals(con, "Distinct ItemID", "VStockStatusrpt", "ItemID", " Where ProcessID=" & lProcessID & " ")
        Call cmbItemName.AddVals(con, "Distinct ItemDescription", "VStockStatusrpt", "ItemID", " Where ProcessID=" & lProcessID & " ")
    End If
    PrintSQL = ""
    
    cmbItem.Text = "<All Items>"
    
    Me.Show
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.Count = 0 Then Exit Sub
    Dim F As New frmVendIssdDetailS
    Load F
    F.ShowDetail LV.SelectedItem.Text, LV.SelectedItem.Tag
    
End Sub
