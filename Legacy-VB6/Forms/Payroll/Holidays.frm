VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Holidays 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Holidays"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleMode       =   0  'User
   ScaleWidth      =   5454.896
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDesc 
      Height          =   300
      Left            =   1200
      TabIndex        =   9
      Top             =   645
      Width           =   4050
   End
   Begin VB.Frame FLV 
      Height          =   5040
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   5955
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2385
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   1200
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4320
         Left            =   90
         TabIndex        =   2
         Top             =   195
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   7620
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ColHdrIcons     =   "ImageList1"
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
            Text            =   "Day"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   5997
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   2925
         TabIndex        =   10
         Top             =   4620
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
         Left            =   4440
         TabIndex        =   1
         Top             =   4620
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
   Begin Crystal.CrystalReport cr1 
      Left            =   6345
      Top             =   240
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   285
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   503
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
      CustomFormat    =   "dddd  dd-MMM-yyyy"
      Format          =   130023427
      CurrentDate     =   37055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -15
      Top             =   -15
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
            Picture         =   "Holidays.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Holidays.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Holiday :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   2355
      TabIndex        =   7
      Top             =   330
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Year :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   720
      TabIndex        =   6
      Top             =   330
      Width           =   435
   End
   Begin MSForms.ComboBox cmbYear 
      Height          =   300
      Left            =   1185
      TabIndex        =   4
      Top             =   285
      Width           =   1110
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "1958;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   600
      Left            =   5355
      TabIndex        =   3
      Top             =   345
      Width           =   615
      Caption         =   " Add"
      PicturePosition =   65543
      Size            =   "1085;1058"
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
Attribute VB_Name = "Holidays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MinDate As Date
Dim MaxDate As Date
Private Sub cmbYear_Change()
    Call cmbYear_Click
End Sub

Private Sub cmbYear_Click()
    
    MinDate = DateSerial(cmbYear, 1, 1)
    MaxDate = DateSerial(cmbYear, 12, 31)
    DTPicker1 = MinDate
    Call FillList
    
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    'If Txt = 0 Or Val(txtTo) = 0 Or Val(txtPer) = 0 Or Val(txtDeduct) = 0 Then
    '    MsgBox "Can't Insert Null Values.", vbInformation
    '    Exit Sub
    'End If
    LV.Sorted = False
    
    Dim Affect As Long
    
        
    con.Execute "insert into Holidays (" & _
     "DT,Description) Values('" & DTPicker1 & "','" & txtDesc & _
     "')", Affect
     
    If Affect = 1 Then
        Dim Item As ListItem
        Set Item = LV.ListItems.add(, Format(DTPicker1, "dd-MMM-yyyy"), Format(DTPicker1, "dddd"))
        With Item
            .SubItems(1) = Format(DTPicker1, "dd-MMM-yyyy")
            .SubItems(2) = txtDesc
        End With
        txtDesc = ""
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

    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptHolidays.rpt")

    SelFormula = "Year({Holidays.DT})=" & cmbYear
    
    Me.Hide
    f.ShowReport SelFormula & " ", rpt
    Unload Me
    
End Sub

Private Sub DTPicker1_Change()
    With DTPicker1
        If .Value > MaxDate Then
            .Value = MaxDate
        ElseIf .Value < MinDate Then
            .Value = MinDate
        End If
    End With
End Sub

Private Sub Form_Load()
      
    Dim CYear As Long
    CYear = Format(GetServerDate, "yyyy")
    For i = 2002 To CYear + 1
        cmbYear.AddItem i
    Next i
    cmbYear = CYear
    DTPicker = DateSerial(cmbYear, 1, 1)
    
End Sub
Private Sub FillList()

    
    Dim rsList As New ADODB.Recordset
    Dim ITM As ListItem
    Dim SQL As String
    
    
    
    SQL = "SELECT * FROM Holidays WHERE YEAR(DT)='" & cmbYear & "' ORDER BY DT "

    With rsList
        .CursorLocation = adUseClient
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.add(, Format(![DT], "dd-MMM-yyyy"), Format(![DT], "dddd"))
            ITM.SubItems(1) = Format(![DT], "dd-MMM-yyyy")
            ITM.SubItems(2) = ![Description]
            .MoveNEXT
        Next
        .Close
    End With
    Set rsLedger = Nothing
    SQL = ""
    
End Sub




Private Sub txtDedE_KeyPress(KeyAscii As Integer)
    On Error GoTo err
    Dim theKey As String
    Dim Affect As Long
    theKey = LV.SelectedItem.key
    theKey = Right(theKey, Len(theKey) - 1)
    
    If KeyAscii = 13 Then
        con.Execute "Update TaxRanges Set SalFrom=" & _
         txtFromE & ",SalTo=" & txtToE & ",PerDeduct=" & _
         txtPerE & ",Reduction=" & txtDedE & _
         " Where SalFrom=" & theKey, Affect
        If Affect = 1 Then
            With LV.SelectedItem
                .key = "'" & txtFromE
                .Text = txtFromE
                .SubItems(1) = txtToE
                .SubItems(2) = txtPerE
                .SubItems(3) = txtDedE
            End With
        End If
        Pic.Visible = False
    ElseIf KeyAscii = 27 Then
        txtFromE = ""
        txtToE = ""
        txtPerE = ""
        txtDedE = ""
        Pic.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
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
    
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(3).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(3).Width
        .Text = LV.SelectedItem.SubItems(2)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    On Error Resume Next
    If KeyCode = 46 Then
        If MsgBox("Do You Want To Delete This Holiday?", vbQuestion + vbYesNo) = vbYes Then
            If getDBPassword("Delete_Holiday") = False Then Exit Sub
            Dim Affect As Integer
            con.Execute "DELETE FROM Holidays WHERE DT='" & LV.SelectedItem.key & "'", Affect
            If Affect = 1 Then
                LV.ListItems.Remove (LV.SelectedItem.Index)
            End If
        End If
    End If
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        con.Execute "UPDATE Holidays SET Description='" & txtEdit & "' WHERE DT='" & LV.SelectedItem.SubItems(1) & "'"
        LV.SelectedItem.SubItems(2) = txtEdit
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub
