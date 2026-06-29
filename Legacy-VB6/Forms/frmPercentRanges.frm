VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPercentRanges 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Ranges...."
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleMode       =   0  'User
   ScaleWidth      =   5508.638
   Begin VB.Frame FLV 
      Height          =   6780
      Left            =   60
      TabIndex        =   0
      Top             =   795
      Width           =   6045
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         ScaleHeight     =   285
         ScaleWidth      =   4155
         TabIndex        =   8
         Top             =   540
         Visible         =   0   'False
         Width           =   4185
         Begin VB.TextBox txtToE 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1380
            TabIndex        =   11
            Top             =   0
            Width           =   1400
         End
         Begin VB.TextBox txtFromE 
            Height          =   300
            Left            =   -15
            TabIndex        =   10
            Top             =   0
            Width           =   1400
         End
         Begin VB.TextBox txtPerE 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2775
            TabIndex        =   9
            Top             =   0
            Width           =   1400
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6060
         Left            =   90
         TabIndex        =   3
         Top             =   195
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   10689
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "From"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "To"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Percentage"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   3045
         TabIndex        =   12
         Top             =   6330
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
         Left            =   4530
         TabIndex        =   1
         Top             =   6330
         Width           =   1425
         ForeColor       =   0
         Caption         =   " Close     "
         PicturePosition =   327683
         Size            =   "2514;635"
         Picture         =   "frmPercentRanges.frx":0000
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
      Left            =   5670
      Top             =   225
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
      Left            =   5595
      Top             =   -195
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
            Picture         =   "frmPercentRanges.frx":0112
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPercentRanges.frx":0564
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   600
      Left            =   3810
      TabIndex        =   7
      Top             =   120
      Width           =   615
      Caption         =   " Add"
      PicturePosition =   65543
      Size            =   "1085;1058"
      Picture         =   "frmPercentRanges.frx":0676
      Accelerator     =   65
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtPer 
      Height          =   285
      Left            =   2685
      TabIndex        =   6
      Top             =   390
      Width           =   1110
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "1958;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtTo 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   390
      Width           =   1095
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "1931;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtFrom 
      Height          =   285
      Left            =   435
      TabIndex        =   4
      Top             =   390
      Width           =   1095
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "1931;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   435
      TabIndex        =   2
      Top             =   165
      Width           =   3360
      BackColor       =   11517387
      Caption         =   "  Qty. From        Qty. To            %"
      Size            =   "5927;423"
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
Attribute VB_Name = "frmPercentRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()

    On Error GoTo ERR
    If Val(txtFrom) = 0 Or Val(txtTo) = 0 Or Val(txtPer) = 0 Then
        MsgBox "Can't Insert Null Values.", vbInformation
        Exit Sub
    End If
    If Val(txtFrom) > Val(txtTo) Then
        MsgBox "Invalid Entry"
        Exit Sub
    End If
    
    
    
    Dim Affect As Long
    
        
    con.Execute "insert into PercentageRanges (QtyFrom," & _
     "QtyTo,PerIncrease) Values(" & _
     txtFrom & "," & txtTo & "," & txtPer & ")", Affect
     
     
    If Affect = 1 Then
        Dim Item As ListItem
        Set Item = LV.ListItems.Add(, "'" & txtFrom, txtFrom)
        With Item
            .SubItems(1) = txtTo
            .SubItems(2) = txtPer
            '.SubItems(3) = t
        End With
        txtFrom = ""
        txtTo = ""
        txtPer = ""
        txtDeduct = ""
        txtFrom = ""
    End If
    
    Exit Sub
ERR:
    MsgBox ERR.Description
End Sub

    
Private Sub cmdCheckAll_Click()
    With LV
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = True
        Next i
    End With
End Sub

Private Sub cmdClearAll_Click()
    With LV
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = False
        Next i
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
On Error GoTo ERR
Screen.MousePointer = vbHourglass

With cr1
    .ReportFileName = RptPRPath & "\TaxRanges.rpt"
    
    '.DataFiles(0) = DatabasePath
    .Connect = ConnectStr
    
    '.SelectionFormula = "{VOrdItems.VendID}='" & cmbCustID & _
      "' AND ({VOrdItems.Dated}=date(" & _
      Format(DtFrom.Tag, "yyyy,MM,dd") & _
      ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
      
    '.Formulas(0) = "CompanyName='" & getCompanyName & "'"
    '.Formulas(1) = "FromTo='From " & Format(DtFrom.Tag, "dd-MMM-yyyy") & " To " & Format(DtTo.Tag, "dd-MMM-yyyy") & "'"
    '.Formulas(2) = "Ledger='Ledger of " & cmbCustName & "'"
    '.Formulas(2) = "SubReport=" & LVItems.Visible
    'if Cheques Are Present then Show Cheque Detail Also
    'If LVItems.Visible Then
    '    .SubreportToChange = "ChqLedger"
    '    .DataFiles(0) = DBName
    '    .SelectionFormula = "{VChqLedger.AccNo}='" & cmbCustID & "' AND ({VChqLedger.CDate}=date(" & Format(DtFrom.Tag, "yyyy,MM,dd") & ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
    'End If
    .Action = 1
    .PageZoomNext
End With

Screen.MousePointer = vbDefault
Exit Sub
ERR:
    MsgBox ERR.Description
    Screen.MousePointer = vbDefault
End Sub






Private Sub Form_Load()

    Call FillList
End Sub
Private Sub FillList()

    
    Dim rsList As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Sql As String
    
    
    
    Sql = "select * from PercentageRanges Order By EntryID"

    With rsList
        .CursorLocation = adUseClient
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.Add(, ![EntryID] & "'", ![QtyFrom])
            ITM.SubItems(1) = ![QtyTo]
            ITM.SubItems(2) = ![PerIncrease]
            'ITM.SubItems(3) = ![Reduction]
            .MoveNext
        Next
        .Close
    End With
    Set rsLedger = Nothing
    Sql = ""
    
End Sub




Private Sub LV_DblClick()
    If LV.ListItems.Count = 0 Then Exit Sub
    
    With LV.SelectedItem
        Pic.Move LV.Left, LV.Top + LV.SelectedItem.Top
        txtFromE = .Text
        txtToE = .SubItems(1)
        txtPerE = .SubItems(2)
        'txtDedE = .SubItems(3)
    End With
    Pic.Visible = True
    With txtFromE
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub txtFromE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With txtToE
            .SelStart = 0
            .SelLength = Len(.Text)
            txtToE.SetFocus
        End With
    ElseIf KeyAscii = 27 Then
        txtFromE = ""
        Pic.Visible = False
    End If
End Sub
Private Sub txtToE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With txtPerE
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyAscii = 27 Then
        txtFromE = ""
        txtToE = ""
        txtPerE = ""
        Pic.Visible = False
    End If
End Sub
Private Sub txtPerE_KeyPress(KeyAscii As Integer)

    On Error GoTo ERR
    Dim theKey As String
    Dim Affect As Long
    theKey = Val(LV.SelectedItem.key)
    'theKey = Right(theKey, Len(theKey) - 1)
    
    If KeyAscii = 13 Then
        con.Execute "Update PercentageRanges Set QtyFrom=" & _
         txtFromE & ",QtyTo=" & txtToE & ",PerIncrease=" & _
         txtPerE & " Where EntryID=" & Val(theKey), Affect
        If Affect = 1 Then
            With LV.SelectedItem
                .key = "'" & txtFromE
                .Text = txtFromE
                .SubItems(1) = txtToE
                .SubItems(2) = txtPerE
                '.SubItems(3) = txtDedE
            End With
        End If
        Pic.Visible = False
    ElseIf KeyAscii = 27 Then
        txtFromE = ""
        txtToE = ""
        txtPerE = ""
        'txtDedE = ""
        Pic.Visible = False
    End If
    
    Exit Sub
ERR:
    MsgBox ERR.Description

End Sub
Private Sub txtDedE_KeyPress(KeyAscii As Integer)
End Sub
