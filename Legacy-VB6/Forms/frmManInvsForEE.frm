VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmManInvsForEE 
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12465
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   12465
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9615
      Left            =   45
      TabIndex        =   16
      Top             =   30
      Width           =   12375
      Begin VB.TextBox txtFormE 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7890
         MaxLength       =   150
         TabIndex        =   5
         Top             =   1050
         Width           =   1950
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   1035
         Width           =   1740
         _ExtentX        =   3069
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
      Begin VB.Frame FNew 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   285
         TabIndex        =   22
         Top             =   1350
         Width           =   10230
         Begin VB.TextBox txtPakAmt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3900
            MaxLength       =   150
            TabIndex        =   8
            Top             =   270
            Width           =   1920
         End
         Begin VB.TextBox txtForAmt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1950
            MaxLength       =   150
            TabIndex        =   7
            Top             =   270
            Width           =   1920
         End
         Begin VB.TextBox txtTotAmt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   0
            MaxLength       =   150
            TabIndex        =   6
            Top             =   270
            Width           =   1920
         End
         Begin VB.CommandButton cmdAddItem 
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   9030
            Picture         =   "frmManInvsForEE.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Add This Entry To Voucher"
            Top             =   15
            UseMaskColor    =   -1  'True
            Width           =   540
         End
         Begin MSComCtl2.DTPicker DTRel 
            Height          =   300
            Left            =   5850
            TabIndex        =   9
            Top             =   270
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   529
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
            CustomFormat    =   "dddd, MMMM dd, yyyy"
            Format          =   56950787
            CurrentDate     =   37250
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Total Amount                 Foreign Amout                Pak Amount                      Realisation Date"
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
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   8985
         End
      End
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3810
         MaxLength       =   150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1005
      End
      Begin VB.TextBox txtPInvoice 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4845
         MaxLength       =   150
         TabIndex        =   3
         Top             =   1050
         Width           =   1470
      End
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11715
         Picture         =   "frmManInvsForEE.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Remove This Entry From Voucher"
         Top             =   1965
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   540
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   300
         Left            =   6345
         TabIndex        =   4
         Top             =   1035
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   529
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
         Format          =   56950787
         CurrentDate     =   37250
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6885
         Left            =   285
         TabIndex        =   11
         Top             =   1965
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   12144
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         TextBackground  =   -1  'True
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
            Text            =   "S.No"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PInvoice"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Form E"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "F Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Pak Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "R Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   6390
         TabIndex        =   17
         Top             =   8850
         Width           =   5280
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   3510
            TabIndex        =   14
            Top             =   195
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
         Begin MSForms.CommandButton cmdSClear 
            Height          =   360
            Left            =   90
            TabIndex        =   12
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSClose 
            Height          =   360
            Left            =   1800
            TabIndex        =   13
            Top             =   195
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
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   12360
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Invoices For EE"
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
            TabIndex        =   20
            Top             =   165
            Width           =   12225
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoices For EE"
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
            TabIndex        =   21
            Top             =   180
            Width           =   12210
         End
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmManInvsForEE.frx":0A64
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
         Index           =   7
         Left            =   285
         TabIndex        =   18
         Top             =   795
         Width           =   9555
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2055
         TabIndex        =   1
         Top             =   1050
         Width           =   1725
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3043;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmManInvsForEE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strDelList As String
Private Sub FillList()
    
    On Error GoTo err
    
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select * From ManualInvsForEE", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
        
            Set ITM = LV.ListItems.Add(, ![EntryID] & "'", .AbsolutePosition)
            
            ITM.ListSubItems.Add , , ![CustCode] & " " & ![Country]
            ITM.ListSubItems.Add , , ![PInvoice] & ""
            ITM.ListSubItems.Add , , Format(![DT], "dd-MMM-yyyy")
            ITM.ListSubItems.Add , , ![FormE] & ""
            ITM.ListSubItems.Add , , Val(![TotalCommercialAmt] & "")
            ITM.ListSubItems.Add , , Val(![FAmount] & "")
            ITM.ListSubItems.Add , , Val(![PAmount] & "")
            ITM.ListSubItems.Add , , Format(![RealisationDate], "dd-MMM-yyyy")
   
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing

    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    
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
            .MoveNext
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdAddItem_Click()
    
    On Error GoTo err
    

    Dim ITM As ListItem
    Dim Sno As Integer
    If LV.ListItems.Count > 0 Then
        Sno = Val(LV.ListItems(LV.ListItems.Count).Text) + 1
    Else
        Sno = 1
    End If
    
    Set ITM = LV.ListItems.Add(, txtFormE, Sno)
    ITM.Tag = "New"
    ITM.ListSubItems.Add , , cmbCust.Text & " " & cmbCountry
    
    ITM.ListSubItems(1).Tag = cmbCust.Text
    
    ITM.ListSubItems.Add , , txtPInvoice
    ITM.ListSubItems(2).Tag = cmbCountry.Text
    
    ITM.ListSubItems.Add , , Format(DT.value, "dd-MMM-yyyy")
    ITM.ListSubItems(3).Tag = txtCurrency
    
    ITM.ListSubItems.Add , , txtFormE
    ITM.ListSubItems.Add , , Val(txtTotAmt)
    ITM.ListSubItems.Add , , Val(txtForAmt)
    ITM.ListSubItems.Add , , Val(txtPakAmt)
    ITM.ListSubItems.Add , , Format(DTRel.value, "dd-MMM-yyyy")
   
   
    txtPInvoice = ""
    txtFormE = ""
    txtTotAmt = ""
    txtForAmt = ""
    txtPakAmt = ""
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()

    If LV.ListItems.Count = 0 Then Exit Sub
    
    If LV.SelectedItem.ListSubItems(6).Tag <> "New" Then
        strDelList = strDelList & Val(LV.SelectedItem.Key) & ","
        LV.ListItems.Remove LV.SelectedItem.Index
    End If
    
End Sub

Private Sub cmdSClear_Click()
    If Saved Then
        Unload Me
        Me.Show
    End If
End Sub

Private Sub cmdSClose_Click()
    If Saved Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
   
    On Error GoTo err
    strOrderNo = ""
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'HICO'")
    DT = ServerDate
    DTRel = DateAdd("m", 1, ServerDate)
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    strDelList = ""
    Call FillList
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Private Sub Form_Resize()
On Error Resume Next
   FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub



Private Sub LV_Click()
    If LV.ListItems.Count > 0 Then
        Call LV_ItemClick(LV.SelectedItem)
    End If
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With cmdDelete
        .Move .Left, LV.Top + Item.Top
        .Visible = True
    End With
End Sub

Private Function Saved() As Boolean
    
    Saved = False
    
    Call StartTrans(con)
    
    
    If strDelList <> "" Then
        strDelList = Left(strDelList, Len(strDelList) - 1)
             
        con.Execute "Delete From ManualInvsForEE Where EntryID In(" & strDelList & ")"
    End If
    
    'Now Save Items
    For i = 1 To LV.ListItems.Count
        If LV.ListItems(i).Tag = "New" Then
            With LV.ListItems(i)
            
                con.Execute "Insert into ManualInvsForEE(CustCode,Country,PInvoice," & _
                 "FormE,DT,TotalCommercialAmt,Curr,FAmount,PAmount,RealisationDate) " & _
                 "Values ('" & .ListSubItems(1).Tag & "','" & .ListSubItems(2).Tag & "','" & _
                 .SubItems(2) & "','" & .SubItems(4) & "','" & .SubItems(3) & "'," & _
                 Val(.SubItems(5)) & ",'" & .ListSubItems(3).Tag & "'," & _
                 Val(.SubItems(6)) & "," & Val(.SubItems(7)) & ",'" & .SubItems(8) & "')"
                 
            End With
        End If
    Next

    con.CommitTrans

    Saved = True
    Exit Function

err:
    MsgBox err.Description
    
End Function
