VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmCommercialCovering 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commercial Covering Letter"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCommercialCovering.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7860
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   5565
      Left            =   75
      TabIndex        =   3
      Top             =   645
      Width           =   7725
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         ScaleHeight     =   225
         ScaleWidth      =   6630
         TabIndex        =   18
         Top             =   1245
         Visible         =   0   'False
         Width           =   6660
         Begin VB.TextBox txtDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            BorderStyle     =   0  'None
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   -15
            MaxLength       =   150
            TabIndex        =   20
            Top             =   0
            Width           =   5000
         End
         Begin VB.TextBox txtCopies 
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
            BorderStyle     =   0  'None
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   5010
            TabIndex        =   19
            Top             =   0
            Width           =   1600
         End
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   75
         TabIndex        =   7
         Top             =   480
         Width           =   2280
         _ExtentX        =   4022
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
         Height          =   4005
         Left            =   75
         TabIndex        =   16
         Top             =   810
         Width           =   7560
         _ExtentX        =   13335
         _ExtentY        =   7064
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Description"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No. Of Copies"
            Object.Width           =   2822
         EndProperty
      End
      Begin VB.Frame FraSave 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4080
         TabIndex        =   13
         Top             =   4830
         Visible         =   0   'False
         Width           =   3555
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   90
            TabIndex        =   15
            Top             =   180
            Width           =   1665
            Caption         =   "Save        "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmCommercialCovering.frx":014A
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   1800
            TabIndex        =   14
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
      End
      Begin VB.Frame FraAdd 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   705
         TabIndex        =   4
         Top             =   4830
         Width           =   6945
         Begin MSForms.CommandButton cmdEdit 
            Height          =   360
            Left            =   1785
            TabIndex        =   12
            Top             =   180
            Width           =   1665
            Caption         =   "Edit Selected   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   69
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAdd 
            Height          =   360
            Left            =   75
            TabIndex        =   11
            Top             =   180
            Width           =   1665
            Caption         =   "Add New     "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmCommercialCovering.frx":025C
            Accelerator     =   78
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
            Left            =   5175
            TabIndex        =   6
            Top             =   180
            Width           =   1665
            Caption         =   "Close            "
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
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   3480
            TabIndex        =   5
            Top             =   180
            Width           =   1665
            Caption         =   "Print             "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmCommercialCovering.frx":036E
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
      Begin MSForms.ComboBox cmbInvNoNew 
         Height          =   285
         Left            =   4680
         TabIndex        =   17
         Tag             =   "AutoCatID"
         Top             =   480
         Visible         =   0   'False
         Width           =   2430
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4286;503"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2385
         TabIndex        =   10
         Top             =   480
         Width           =   2265
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3995;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Customer                                 Country                                   Invoice No"
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
         Index           =   15
         Left            =   75
         TabIndex        =   9
         Top             =   240
         Width           =   7035
      End
      Begin MSForms.ComboBox cmbInvNo 
         Height          =   285
         Left            =   4680
         TabIndex        =   8
         Tag             =   "AutoCatID"
         Top             =   480
         Width           =   2430
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4286;503"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   435
      TabIndex        =   0
      Top             =   0
      Width           =   6960
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Covering Letter"
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
         Left            =   30
         TabIndex        =   1
         Top             =   150
         Width           =   6765
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Covering Letter"
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
         Left            =   45
         TabIndex        =   2
         Top             =   165
         Width           =   6750
      End
   End
End
Attribute VB_Name = "frmCommercialCovering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PInvoice As String
Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    
    cmbCust.Tag = cmbCust.Text & cmbCountry
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    
    If strProforma <> "" Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select Distinct CustomInvoice,CCoveringLetterL From VCommercialCoveringLetter Where CustCode+Country='" & cmbCust.ID & cmbCountry.Text & "' ", con, adOpenForwardOnly, adLockReadOnly
        cmbInvNo.Clear
        Do Until .EOF
            cmbInvNo.AddItem .Fields(0) & ""
            cmbInvNo.List(cmbInvNo.ListCount - 1, 1) = .Fields(1) & ""
            .MoveNEXT
        Loop
        .Close
        
        .Open "Select CustomInvoice From CustomInvoice Where CustomInvoice Not In (Select CustomInvoice From CommercialCoveringLetter) and CustCode+Country='" & cmbCust.ID & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbInvNoNew.Clear
        Do Until .EOF
            cmbInvNoNew.AddItem .Fields(0) & ""
            
            .MoveNEXT
        Loop
        .Close
        
    End With
    Set rs = Nothing
    
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbCust_matched()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Country,Curr From ForeignCustomers Where CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbinvno_Change()
    Call cmbinvno_click
End Sub

Private Sub cmbInvNoNew_Click()

    On Error GoTo err
    If cmbInvNoNew.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From VrptCustomInvoiceComplete Where CustomInvoice='" & cmbInvNoNew & "'", con, adOpenForwardOnly, adLockReadOnly
        
        Set ITM = LV.ListItems.add(, , "Commercial Invoice")
        ITM.SubItems(1) = "(3)"
        
        Set ITM = LV.ListItems.add(, , "PACKING LIST")
        ITM.SubItems(1) = "(1)"
        
        Set ITM = LV.ListItems.add(, , "GSP FORM (A)")
        ITM.SubItems(1) = "(1)"
        
        Set ITM = LV.ListItems.add(, , "Bill of lading")
        ITM.SubItems(1) = "(3)"
        
        Set ITM = LV.ListItems.add(, , "FDA Package")
        ITM.SubItems(1) = "(2)"
        
        Set ITM = LV.ListItems.add(, , "Letter to Customer")
        ITM.SubItems(1) = "(1)"
        
        Set ITM = LV.ListItems.add(, , "Any Instruction Letter")
        ITM.SubItems(1) = "(1)"
      
        Set ITM = LV.ListItems.add(, , "AWB/EMS/Courier Receipts")
        ITM.SubItems(1) = "(1)"
        
        Set ITM = LV.ListItems.add(, "New", "<<      Dbl Click Here To Add More.     >>")
        
        
        .Close
        
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbinvno_click()

    On Error GoTo err
    If cmbInvNo.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "Select * From VCommercialCoveringLetter Where CustomInvoice='" & cmbInvNo & "' Order By EntryID", con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , ![EntryText] & "")
            ITM.SubItems(1) = ![EntryDetail] & ""
            .MoveNEXT
        Loop
        
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmdAdd_Click()
        
        
    Call lok(False)
    PInvoice = ""
    LV.ListItems.Clear
    
End Sub
Private Sub lok(Locked As Boolean)
    
    If Locked Then PInvoice = ""
    
    FraAdd.Visible = Locked
    FraSave.Visible = Not Locked
    
    cmbInvNo.Visible = Locked
    cmbInvNoNew.Visible = Not Locked
    
    cmbInvNo.Locked = False
    
End Sub


Private Sub cmdCancel_Click()
    Call cmbinvno_click
    Call lok(True)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    
    If cmbInvNo.MatchFound = False Then Exit Sub
'    If cmbInvNo.List(cmbInvNo.ListIndex, 1) Then
'        MsgBox "This Record Has Been Locked.", vbInformation
'        Exit Sub
'    End If
    PInvoice = cmbInvNo
    LV.ListItems.add , "New", "<<      Dbl Click Here To Add More.     >>"
    FraSave.Visible = True
    FraAdd.Visible = False
    cmbInvNo.Locked = True
    
    
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    If cmbInvNo.MatchFound = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report 'New rptCommercialCoveringLetter
    Set rpt = rptApp.OpenReport(RptPath & "\rptCommercialCoveringLetter.rpt") '
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VCommercialCoveringLetter.CustomInvoice}='" & cmbInvNo.Text & "' ", rpt
    'frmPrevRpt.ShowReport
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
    Call StartTrans(con)
    
    If PInvoice <> "" Then
        con.Execute "Delete From CommercialCoveringLetter Where CustomInvoice='" & PInvoice & "'"
    End If
    
    For i = 1 To LV.ListItems.count - 1
        con.Execute "Insert Into CommercialCoveringLetter(CustomInvoice,EntryText,EntryDetail) Values('" & _
         IIf(PInvoice = "", cmbInvNoNew, cmbInvNo) & "','" & LV.ListItems(i) & "','" & _
         LV.ListItems(i).SubItems(1) & "')"
    Next i
    
    con.CommitTrans
    
    Call lok(True)
    Call cmbCountry_Click
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub Form_Load()

    On Error GoTo err
    
    cmbCust.ListHeight = 2100
    
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where custCode<>'Hico'")
    
    Call SaveLV(LV)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_Click()
    PicEdit.Visible = False
End Sub

Private Sub LV_DblClick()
    
    On Error GoTo err
    
    Call ResetLV(LV)
    PicEdit.Move PicEdit.Left, LV.Top + LV.SelectedItem.Top
    
    With txtDesc
        .Text = LV.SelectedItem.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    txtCopies = LV.SelectedItem.SubItems(1) & ""
    
    PicEdit.Visible = True
    txtDesc.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        LV.ListItems.Remove LV.SelectedItem.Index
    End If
    
End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If txtCopies = "" Then
            MsgBox "Invalid Entry!", vbInformation
            Exit Sub
        Else
            LV.SelectedItem.Text = txtDesc
            LV.SelectedItem.SubItems(1) = txtCopies
                        
            If LV.SelectedItem.key = "New" Then
                LV.SelectedItem.key = ""
                LV.ListItems.add , "New", "<<      Dbl Click Here To Add More.     >>"
            End If

            PicEdit.Visible = False
        End If
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If txtDesc = "" Then
            MsgBox "Invalid Entry!", vbInformation
            Exit Sub
        Else
            With txtCopies
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
        End If
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub
