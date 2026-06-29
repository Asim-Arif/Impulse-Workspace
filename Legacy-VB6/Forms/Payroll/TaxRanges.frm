VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form TaxRanges 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Settings...."
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleMode       =   0  'User
   ScaleWidth      =   5548.945
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5085
      Left            =   90
      TabIndex        =   1
      Top             =   300
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8969
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tax Deduction"
      TabPicture(0)   =   "TaxRanges.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdAdd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtDeduction"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtTo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtFrom"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FLV"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Tax Reduction"
      TabPicture(1)   =   "TaxRanges.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAdd1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtFrom1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtTo1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtReduction"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame1 
         Height          =   4080
         Left            =   -74970
         TabIndex        =   15
         Top             =   900
         Width           =   5955
         Begin VB.PictureBox Pic1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   180
            ScaleHeight     =   285
            ScaleWidth      =   5565
            TabIndex        =   22
            Top             =   510
            Visible         =   0   'False
            Width           =   5595
            Begin VB.TextBox txtToE1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   1830
               TabIndex        =   25
               Top             =   0
               Width           =   1850
            End
            Begin VB.TextBox txtFromE1 
               BackColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   -15
               TabIndex        =   24
               Top             =   0
               Width           =   1850
            End
            Begin VB.TextBox txtPerE1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   3690
               TabIndex        =   23
               Top             =   0
               Width           =   1860
            End
         End
         Begin MSComctlLib.ListView LV1 
            Height          =   3720
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   6562
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
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
               Object.Width           =   3263
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "To"
               Object.Width           =   3263
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Reduction"
               Object.Width           =   3175
            EndProperty
         End
      End
      Begin VB.Frame FLV 
         Height          =   4080
         Left            =   30
         TabIndex        =   2
         Top             =   900
         Width           =   5955
         Begin VB.PictureBox Pic 
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            ScaleHeight     =   255
            ScaleWidth      =   5565
            TabIndex        =   3
            Top             =   990
            Visible         =   0   'False
            Width           =   5595
            Begin VB.TextBox txtPerE 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   3690
               TabIndex        =   6
               Top             =   0
               Width           =   1860
            End
            Begin VB.TextBox txtFromE 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   -15
               TabIndex        =   5
               Top             =   0
               Width           =   1850
            End
            Begin VB.TextBox txtToE 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1830
               TabIndex        =   4
               Top             =   0
               Width           =   1850
            End
         End
         Begin MSComctlLib.ListView LV 
            Height          =   3720
            Left            =   120
            TabIndex        =   7
            Top             =   270
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   6562
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
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
               Object.Width           =   3263
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "To"
               Object.Width           =   3263
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Deduction"
               Object.Width           =   3175
            EndProperty
         End
      End
      Begin MSForms.CommandButton cmdAdd1 
         Height          =   390
         Left            =   -70260
         TabIndex        =   21
         Top             =   510
         Width           =   1065
         Caption         =   " Add"
         PicturePosition =   131072
         Size            =   "1879;688"
         Picture         =   "TaxRanges.frx":0038
         Accelerator     =   65
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label Label1 
         Height          =   240
         Left            =   -74925
         TabIndex        =   20
         Top             =   390
         Width           =   4305
         BackColor       =   11517387
         VariousPropertyBits=   8388627
         Caption         =   "  Sal. From             Sal. To                  Reduction %"
         Size            =   "7594;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtFrom1 
         Height          =   300
         Left            =   -74940
         TabIndex        =   19
         Top             =   645
         Width           =   1425
         VariousPropertyBits=   746604571
         Size            =   "2514;529"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtTo1 
         Height          =   300
         Left            =   -73485
         TabIndex        =   18
         Top             =   645
         Width           =   1425
         VariousPropertyBits=   746604571
         Size            =   "2514;529"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtReduction 
         Height          =   300
         Left            =   -72030
         TabIndex        =   17
         Top             =   645
         Width           =   1425
         VariousPropertyBits=   746604571
         Size            =   "2514;529"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.Label Label2 
         Height          =   240
         Left            =   75
         TabIndex        =   13
         Top             =   390
         Width           =   4305
         BackColor       =   11517387
         VariousPropertyBits=   8388627
         Caption         =   "  Sal. From            Sal. To                    Deduction %"
         Size            =   "7594;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtFrom 
         Height          =   300
         Left            =   60
         TabIndex        =   12
         Top             =   645
         Width           =   1425
         VariousPropertyBits=   746604571
         Size            =   "2514;529"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtTo 
         Height          =   300
         Left            =   1515
         TabIndex        =   11
         Top             =   645
         Width           =   1425
         VariousPropertyBits=   746604571
         Size            =   "2514;529"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtDeduction 
         Height          =   300
         Left            =   2970
         TabIndex        =   10
         Top             =   645
         Width           =   1425
         VariousPropertyBits=   746604571
         Size            =   "2514;529"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.CommandButton cmdAdd 
         Height          =   390
         Left            =   4710
         TabIndex        =   9
         Top             =   480
         Width           =   1065
         Caption         =   " Add"
         PicturePosition =   131072
         Size            =   "1879;688"
         Picture         =   "TaxRanges.frx":014A
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
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   1245
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   4290
      Top             =   3690
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
      Left            =   4770
      Top             =   3840
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
            Picture         =   "TaxRanges.frx":025C
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TaxRanges.frx":06AE
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   3210
      TabIndex        =   14
      Top             =   5430
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
      Left            =   4650
      TabIndex        =   8
      Top             =   5430
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
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuDeductions 
         Caption         =   "Tax Deductions"
      End
      Begin VB.Menu mnuReductions 
         Caption         =   "Tax Reductions"
      End
   End
End
Attribute VB_Name = "TaxRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()

    'On Error GoTo Err
    
    If PrevRange Then
        MsgBox "Range Already Exist !", vbInformation
        Exit Sub
    End If
    
    If Val(txtFrom) = 0 Or Val(txtTo) = 0 Or txtDeduction = "" Then
        MsgBox "Can't Insert Null Values.", vbInformation
        Exit Sub
    End If
    
    If Val(txtFrom) > Val(txtTo) Then
        MsgBox "Starting Range Could Not Be Greater Than The Ending Range !", vbInformation
        Exit Sub
    End If
    
    
    
    Dim Affect As Long
    
        
    con.Execute "insert into TaxRanges (SalFrom," & _
     "SalTo,Deduction) Values(" & _
     txtFrom & "," & txtTo & "," & txtDeduction & ")", Affect
     
    If Affect = 1 Then
        Dim Item As ListItem
        Set Item = LV.ListItems.add(, "'" & txtFrom, txtFrom)
        With Item
            .SubItems(1) = txtTo
            .SubItems(2) = txtDeduction & " %"
        End With
        txtFrom = ""
        txtTo = ""
        txtDeduction = ""
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Function PrevRange() As Boolean

Dim Temp As New ADODB.Recordset
Dim SalFrom  As Double, SalTo As Double

With Temp
    .CursorLocation = adUseClient
    .Open "SELECT * FROM TaxRanges", con, adOpenKeyset, adLockReadOnly
    If .RecordCount > 0 Then
        For i = 1 To .RecordCount
            SalFrom = Val(!SalFrom & "")
            SalTo = Val(!SalTo & "")
            
            If Val(txtFrom) > SalFrom And Val(txtFrom) < SalTo Then
                PrevRange = True
                Exit For
            ElseIf Val(txtFrom) = SalFrom Then
                PrevRange = True
                Exit For
            End If
            .MoveNext
        Next
    
    End If
    
End With
End Function
    
Private Sub cmdAdd1_Click()

    'On Error GoTo Err
    
    If PrevRange1 Then
        MsgBox "Range Already Exist !", vbInformation
        Exit Sub
    End If
    
    If Val(txtFrom1) = 0 Or Val(txtTo1) = 0 Or txtReduction = "" Then
        MsgBox "Can't Insert Null Values.", vbInformation
        Exit Sub
    End If
    
    If Val(txtFrom1) > Val(txtTo1) Then
        MsgBox "Starting Range Could Not Be Greater Than The Ending Range !", vbInformation
        Exit Sub
    End If
    
    
    
    Dim Affect As Long
    
        
    con.Execute "insert into TaxReduction (SalFrom," & _
     "SalTo,Reduction) Values(" & _
     txtFrom1 & "," & txtTo1 & "," & txtReduction & _
     ")", Affect
     
    If Affect = 1 Then
        Dim Item As ListItem
        Set Item = LV1.ListItems.add(, "'" & txtFrom1, "Rs." & txtFrom1)
        With Item
            .SubItems(1) = txtTo1
            .SubItems(2) = txtReduction & " %"
        End With
        txtFrom1 = ""
        txtTo1 = ""
        txtReduction = ""
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Function PrevRange1() As Boolean

Dim Temp As New ADODB.Recordset
Dim SalFrom  As Double, SalTo As Double

With Temp
    .CursorLocation = adUseClient
    .Open "SELECT * FROM TaxReduction", con, adOpenKeyset, adLockReadOnly
    If .RecordCount > 0 Then
        For i = 1 To .RecordCount
            SalFrom = Val(!SalFrom & "")
            SalTo = Val(!SalTo & "")
            
            If Val(txtFrom1) > SalFrom And Val(txtFrom1) < SalTo Then
                PrevRange1 = True
                Exit For
            ElseIf Val(txtFrom1) = SalFrom Then
                PrevRange1 = True
                Exit For
            End If
            .MoveNext
        Next
    
    End If
    
End With
End Function
    
    

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
'MsgBox getTaxAmt(Val(Text1) * 12) / 12
'Exit Sub

'On Error GoTo ERR
'Screen.MousePointer = vbHourglass

With cr1
'    .ReportFileName = RptPath & "\TaxRanges.rpt"
    
    '.DataFiles(0) = DatabasePath
'    .Connect = ConnectStr
    
    '.SelectionFormula = "{VOrdItems.VendID}='" & cmbCustID & _
      "' AND ({VOrdItems.Dated}=date(" & _
      Format(DtFrom.Tag, "yyyy,MM,dd") & _
      ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
      
    '.Formulas(0) = "CompanyName='" & getCompanyName & "'"
    '.Formulas(1) = "FromTo='From " & Format(DtFrom.Tag, "dd-MMM-yyyy") & " To " & Format(DtTo.Tag, "dd-MMM-yyyy") & "'"
'    .Formulas(0) = "Company='" & company & "'"
'    .Action = 1
'    .PageZoomNext
End With

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptTaxRanges.rpt")

Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    End If
Next

'SelFormula = "Year({Holidays.DT})=" & cmbYear
Me.Hide
f.ShowReport " ", rpt

Screen.MousePointer = vbDefault

End Sub






Private Sub Form_Load()

    Call FillList
    Call FillList1
End Sub
Private Sub FillList()

    
    Dim rsList As New ADODB.Recordset
    Dim ITM As ListItem
    Dim SQL As String
    
    
    
    SQL = "select * from TaxRanges Order By SalFrom"

    With rsList
        .CursorLocation = adUseClient
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.add(, "'" & ![SalFrom], "Rs. " & ![SalFrom])
            ITM.SubItems(1) = "Rs. " & ![SalTo]
            ITM.SubItems(2) = ![Deduction] & " %"
            .MoveNext
        Next
        .Close
    End With
    Set rsLedger = Nothing
    SQL = ""
    
End Sub


Private Sub FillList1()

    
    Dim rsList As New ADODB.Recordset
    Dim ITM As ListItem
    Dim SQL As String
    
    
    
    SQL = "select * from TaxReduction Order By SalFrom"

    With rsList
        .CursorLocation = adUseClient
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV1.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV1.ListItems.add(, "'" & ![SalFrom], "Rs. " & ![SalFrom])
            ITM.SubItems(1) = "Rs. " & ![SalTo]
            ITM.SubItems(2) = ![Reduction] & " %"
            .MoveNext
        Next
        .Close
    End With
    Set rsLedger = Nothing
    SQL = ""
    
End Sub


Private Sub LV_Click()
Pic.Visible = False
End Sub


Private Sub LV1_Click()
Pic1.Visible = False
End Sub


Private Sub LV_DblClick()
    If LV.ListItems.Count = 0 Then Exit Sub
    
    With LV.SelectedItem
        Pic.Move LV.Left, LV.Top + LV.SelectedItem.Top
        
        txtFromE = Replace(.Text, "Rs. ", "")
        txtToE = Replace(.SubItems(1), "Rs. ", "")
        txtPerE = Replace(.SubItems(2), " %", "")

    End With
    
    Pic.Visible = True
    With txtFromE
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LV1_DblClick()
    If LV1.ListItems.Count = 0 Then Exit Sub
    
    With LV1.SelectedItem
        Pic1.Move LV1.Left, LV1.Top + LV1.SelectedItem.Top
        txtFromE1 = Replace(.Text, "Rs. ", "")
        txtToE1 = Replace(.SubItems(1), "Rs. ", "")
        txtPerE1 = Replace(.SubItems(2), " %", "")
    End With
    
    Pic1.Visible = True
    With txtFromE1
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


Private Sub txtFromE1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With txtToE1
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyAscii = 27 Then
        txtFromE1 = ""
        Pic1.Visible = False
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

Private Sub txtToE1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With txtPerE1
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyAscii = 27 Then
        txtFromE1 = ""
        txtToE1 = ""
        txtPerE1 = ""
        Pic1.Visible = False
    End If
End Sub


Private Sub txtPerE1_KeyPress(KeyAscii As Integer)
    'On Error GoTo Err
    
    
    Dim theKey As String
    Dim Affect As Long
    theKey = LV1.SelectedItem.key
    theKey = Right(theKey, Len(theKey) - 1)
    
    If KeyAscii = 13 Then
        
        con.Execute "Update TaxReduction Set SalFrom=" & _
         txtFromE1 & ",SalTo=" & txtToE1 & ",Reduction=" & _
         txtPerE1 & " Where SalFrom=" & theKey, Affect
         
        If Affect = 1 Then
            With LV1.SelectedItem
                .key = "'" & txtFromE1
                .Text = "Rs. " & txtFromE1
                .SubItems(1) = "Rs. " & txtToE1
                .SubItems(2) = txtPerE1 & " %"
            End With
        End If
        Pic1.Visible = False
    ElseIf KeyAscii = 27 Then
        txtFromE1 = ""
        txtToE1 = ""
        txtPerE1 = ""
        Pic1.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub txtPerE_KeyPress(KeyAscii As Integer)
    On Error GoTo err
    
    
    Dim theKey As String
    Dim Affect As Long
    theKey = LV.SelectedItem.key
    theKey = Right(theKey, Len(theKey) - 1)
    
    If KeyAscii = 13 Then
        
        con.Execute "Update TaxRanges Set SalFrom=" & _
         txtFromE & ",SalTo=" & txtToE & ",Deduction=" & _
         txtPerE & " Where SalFrom=" & theKey, Affect
         
        If Affect = 1 Then
            With LV.SelectedItem
                .key = "'" & txtFromE
                .Text = "Rs. " & txtFromE
                .SubItems(1) = "Rs. " & txtToE
                .SubItems(2) = txtPerE & " %"
            End With
        End If
        Pic.Visible = False
    ElseIf KeyAscii = 27 Then
        txtFromE = ""
        txtToE = ""
        txtPerE = ""
        Pic.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
