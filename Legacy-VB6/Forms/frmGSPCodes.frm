VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmGSPCodes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSP Codes"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6285
   ScaleMode       =   0  'User
   ScaleWidth      =   5428.024
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FLV 
      Height          =   6255
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5955
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   90
         TabIndex        =   10
         Top             =   465
         Width           =   2520
      End
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   5400
         Picture         =   "frmGSPCodes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Remove This Entry From Voucher"
         Top             =   885
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDesc 
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
         Height          =   495
         Left            =   2625
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   465
         Width           =   2550
      End
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   270
         ScaleHeight     =   255
         ScaleWidth      =   4140
         TabIndex        =   4
         Top             =   1020
         Visible         =   0   'False
         Width           =   4170
         Begin MSForms.TextBox txtDescription 
            Height          =   300
            Left            =   600
            TabIndex        =   6
            Top             =   -15
            Width           =   645
            VariousPropertyBits=   -1467987941
            Size            =   "1138;529"
            SpecialEffect   =   3
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox txtSNo 
            Height          =   300
            Left            =   -15
            TabIndex        =   5
            Top             =   -15
            Width           =   645
            VariousPropertyBits=   679495707
            Size            =   "1138;529"
            SpecialEffect   =   3
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4815
         Left            =   90
         TabIndex        =   3
         Top             =   975
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Country"
            Object.Width           =   4445
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSForms.Label Label2 
         Height          =   240
         Left            =   90
         TabIndex        =   11
         Top             =   225
         Width           =   5085
         Caption         =   "  Country                                        Code"
         Size            =   "8969;423"
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
         Left            =   5205
         TabIndex        =   8
         Top             =   420
         Width           =   705
         Caption         =   " Add"
         PicturePosition =   65543
         Size            =   "1244;609"
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
         Left            =   2445
         TabIndex        =   2
         Top             =   5820
         Visible         =   0   'False
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
         Left            =   3960
         TabIndex        =   1
         Top             =   5820
         Width           =   1425
         ForeColor       =   0
         Caption         =   " Close     "
         PicturePosition =   327683
         Size            =   "2514;635"
         Picture         =   "frmGSPCodes.frx":0532
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
            Picture         =   "frmGSPCodes.frx":0644
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGSPCodes.frx":0800
            Key             =   "Up"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGSPCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdAdd_Click()

    On Error GoTo err
    
    LV.Sorted = False
    
    Dim Affect As Long
    
    If txtID = "" Then
        MsgBox "Invalid Country", vbInformation
        Exit Sub
    ElseIf txtDesc = "" Then
        MsgBox "Invalid Code", vbInformation
        Exit Sub
    End If
    
    con.Execute "Insert Into GSPCodes(Country,Code) " & _
                "Values('" & txtID & "','" & txtDesc & "')", Affect
     
    If Affect = 1 Then
        Call FillList
        txtDesc = ""
    End If
    
    txtID.SetFocus
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
On Error GoTo err
If MsgBox("Remove This Item ?", vbYesNo) = vbNo Then Exit Sub
    Dim Affect As Long
    con.Execute "Delete From Colors where ColorID='" & LV.SelectedItem.key & "'", Affect
    
If Affect > 0 Then Call FillList
Exit Sub
err:
MsgBox "Cannot Remove This Item Due To Error :" & vbNewLine & _
        err.Description, vbCritical
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
    
    SQL = "Select * from GSPCodes Order By Country"

    With rsList
        .CursorLocation = adUseClient
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.add(, , ![Country] & "")
            ITM.SubItems(1) = ![code]
            .MoveNEXT
        Next
        .Close
    End With
    Set rsList = Nothing
    SQL = ""
    
End Sub


Private Sub Form_Load()
   Call FillList
End Sub

Private Sub Form_Resize()
    Me.Top = 1700 'MainForm.Toolbar1.Top + MainForm.Toolbar1.Height + 100
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
'    Exit Sub
    'Edit This Record...
    If LV.ListItems.count = 0 Then Exit Sub
    '-------------------------------------------------------
    PicEdit.Move LV.Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(1).Width + LV.ColumnHeaders(2).Width
    
    
    txtSNo.Move 0, 0, LV.ColumnHeaders(1).Width
    txtSNo.Text = LV.SelectedItem
    
    txtDescription.Move txtSNo.Width - 10, 0, LV.ColumnHeaders(2).Width
    txtDescription.Text = LV.SelectedItem.SubItems(1)
    
    PicEdit.Visible = True
    
    With txtSNo
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PicEdit.Visible = False
      With cmdDelete
        .Move LV.Left + LV.Width, Item.Top + LV.Top
        .Visible = True
    End With
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'For Now
    Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    On Error Resume Next
    If KeyCode = 46 Then
        If MsgBox("Do You Want To Delete This Component?", vbQuestion + vbYesNo) = vbYes Then
            Dim Affect As Integer
            con.Execute "Delete from Components Where CompID=" & Val(LV.SelectedItem.key), Affect
            If Affect = 1 Then
                LV.ListItems.Remove (LV.SelectedItem.Index)
            End If
        End If
    End If
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub txtDescription_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

    Dim iAffect As Integer
    If KeyCode.Value = 13 And Shift = 0 Then
        'Debug.Print Shift
        con.Execute "Update GSPCodes Set Country='" & txtSNo & _
         "',Code='" & txtDescription & "' Where Country='" & _
         LV.SelectedItem & "'", iAffect
        If iAffect = 1 Then
            PicEdit.Visible = False
            Call FillList
        End If
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtSNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    If KeyCode.Value = 13 Then
        With txtDescription
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyCode.Value = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub
