VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmNewSteel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Steel Type"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5205
      Left            =   60
      TabIndex        =   8
      Top             =   45
      Width           =   5475
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   4140
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3150
         Visible         =   0   'False
         Width           =   4170
         Begin VB.TextBox txtDescription 
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
            Height          =   285
            Left            =   465
            TabIndex        =   13
            Top             =   -15
            Width           =   495
         End
         Begin VB.TextBox txtName 
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
            Height          =   285
            Left            =   -15
            TabIndex        =   12
            Top             =   -15
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   2655
         Left            =   120
         TabIndex        =   4
         Top             =   2475
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Symbol"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "LVal"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "UVal"
            Object.Width           =   1058
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Stripe Max. Weight"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   1980
         TabIndex        =   15
         Top             =   1905
         Width           =   1605
      End
      Begin MSForms.TextBox txtMax 
         Height          =   285
         Left            =   1950
         TabIndex        =   3
         Top             =   2115
         Width           =   1815
         VariousPropertyBits=   679495707
         MaxLength       =   100
         BorderStyle     =   1
         ScrollBars      =   2
         Size            =   "3201;503"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Stripe Min. Weight"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   14
         Top             =   1905
         Width           =   1545
      End
      Begin MSForms.TextBox txtMin 
         Height          =   285
         Left            =   105
         TabIndex        =   2
         Top             =   2115
         Width           =   1815
         VariousPropertyBits=   679495707
         MaxLength       =   100
         BorderStyle     =   1
         ScrollBars      =   2
         Size            =   "3201;503"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TBox 
         Height          =   285
         Index           =   0
         Left            =   105
         TabIndex        =   0
         Top             =   435
         Width           =   5280
         VariousPropertyBits=   679495707
         MaxLength       =   100
         BorderStyle     =   1
         ScrollBars      =   2
         Size            =   "9313;503"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TBox 
         Height          =   810
         Index           =   1
         Left            =   105
         TabIndex        =   1
         Top             =   1050
         Width           =   5280
         VariousPropertyBits=   -1467987941
         MaxLength       =   255
         BorderStyle     =   1
         ScrollBars      =   2
         Size            =   "9313;1429"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   10
         Top             =   825
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Steel Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   9
         Top             =   225
         Width           =   900
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   5610
      TabIndex        =   7
      Top             =   1065
      Width           =   1665
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmNewSteel.frx":0000
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   375
      Left            =   5595
      TabIndex        =   5
      Top             =   135
      Width           =   1665
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmNewSteel.frx":0112
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   375
      Left            =   5610
      TabIndex        =   6
      Top             =   600
      Width           =   1665
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmNewSteel.frx":0224
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
Attribute VB_Name = "frmNewSteel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CatID As Integer
Dim SelItm As ListItem
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()

    If Saved Then
        Unload Me
        Me.Show 1
    End If

End Sub

Public Sub EditSteel(iCatID As Integer)
    
    CatID = iCatID
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * From SteelTypes Where SteelID=" & iCatID, con, adOpenForwardOnly, adLockReadOnly
        TBox(0) = ![SteelType] & ""
        TBox(1) = ![Description] & ""
        txtMin = Val(![StripeMinW] & "")
        txtMax = Val(![StripeMaxW] & "")
        .Close
        
        'Now Get The Comps..
        .Open "Select * From SteelTypeComps Where SteelID=" & iCatID, con, adOpenForwardOnly, adLockReadOnly
        For i = 1 To .RecordCount
            LV.ListItems(![CompID] & "'").Checked = True
            LV.ListItems(![CompID] & "'").SubItems(2) = Val(![LVal] & "")
            LV.ListItems(![CompID] & "'").SubItems(3) = Val(![UVal] & "")
            .MoveNext
        Next i
        .Close
    End With
    
    
    Me.Show 1
    
    Set rs = Nothing
    
End Sub
Private Function Saved() As Boolean

On Error GoTo err

    If Trim(TBox(0)) = "" Then
        MsgBox "Steel Type Can't Be Null.", vbInformation
        SaveItem = False
        Exit Function
    End If

    Call StartTrans(con)
    
    
    If CatID = 0 Then
        Dim iNewID As Long
        iNewID = Val(con.Execute("Select Max(SteelID) From SteelTypes").Fields(0).value & "")
        iNewID = iNewID + 1
        
        con.Execute "Insert into SteelTypes(SteelID,SteelType,Description,StripeMinW,StripeMaxW) " & _
         "Values(" & iNewID & ",'" & TBox(0).Text & "','" & TBox(1) & "'," & Val(txtMin) & "," & Val(txtMax) & ")"
         
        CatID = con.Execute("Select Max(SteelID) From SteelTypes").Fields(0).value
    Else
        con.Execute "Delete From SteelTypeComps Where SteelID=" & CatID
        
        con.Execute "Update SteelTypes Set SteelType='" & TBox(0) & _
        "',Description='" & TBox(1) & "',StripeMinW=" & Val(txtMin) & _
        ",StripeMaxW=" & Val(txtMax) & "  Where SteelID=" & CatID
        
    End If
    
    
    
    For i = 1 To LV.ListItems.Count
        If LV.ListItems(i).Checked Then
            con.Execute "Insert Into SteelTypeComps(SteelID,CompID," & _
             "LVal,UVal) Values(" & CatID & "," & _
             Val(LV.ListItems(i).Key) & "," & _
             Val(LV.ListItems(i).SubItems(2)) & "," & _
             Val(LV.ListItems(i).SubItems(3)) & ")"
            
        End If
    Next i
    
    con.CommitTrans
    CatID = 0
    Saved = True
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub Form_Load()
    Call AddCompositions
End Sub
Private Sub AddCompositions()
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    
    With rs
        .Open "Select * From Compositions Order By Symbol", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set itm = LV.ListItems.Add(, ![ID] & "'", ![Symbol] & "")
            itm.SubItems(1) = ![Name] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.Count = 0 Then Exit Sub
    '-------------------------------------------------------
    If LV.SelectedItem.Checked = False Then Exit Sub
    Set SelItm = LV.SelectedItem
    
    PicEdit.Tag = "Edit"
    PicEdit.Move LV.Left + LV.ColumnHeaders(3).Left, LV.SelectedItem.Top + LV.Top, LV.ColumnHeaders(3).Width + LV.ColumnHeaders(4).Width
    
    
    
    txtName.Move 0, 0, LV.ColumnHeaders(3).Width
    txtName.Text = Val(LV.SelectedItem.SubItems(2))
    
    txtDescription.Move txtName.Width - 10, 0, LV.ColumnHeaders(4).Width
    txtDescription.Text = Val(LV.SelectedItem.SubItems(3))
    
    PicEdit.Visible = True
    
    With txtName
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If Item.Checked = False Then
        LV.SelectedItem.SubItems(2) = ""
        LV.SelectedItem.SubItems(3) = ""
        Exit Sub
    End If
    Set SelItm = Item
    '-------------------------------------------------------
    
    PicEdit.Move LV.Left + LV.ColumnHeaders(3).Left, Item.Top + LV.Top, LV.ColumnHeaders(3).Width + LV.ColumnHeaders(4).Width
    
    PicEdit.Tag = "Add"
    
    txtName.Move 0, 0, LV.ColumnHeaders(3).Width
    txtName.Text = Val(Item.SubItems(2))
    
    txtDescription.Move txtName.Width - 10, 0, LV.ColumnHeaders(4).Width
    txtDescription.Text = Val(Item.SubItems(3))
    
    PicEdit.Visible = True
    
    With txtName
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'If PicEdit.Visible Then PicEdit.Visible = False
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtDescription) < Val(txtName) Then
            MsgBox "Please Upper Value Should Be More Than Lower Value.", vbInformation
            Exit Sub
        End If
        SelItm.SubItems(2) = Val(txtName)
        SelItm.SubItems(3) = Val(txtDescription)
        PicEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
        If PicEdit.Tag = "Add" Then
            SelItm.Checked = False
        End If
    End If
    
    Call OnlyNums(KeyAscii)
    
End Sub
Private Sub txtMax_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii.value = OnlyNums1(KeyAscii.value)
End Sub


Private Sub txtMin_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii.value = OnlyNums1(KeyAscii.value)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        With txtDescription
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
        If PicEdit.Tag = "Add" Then
            SelItm.Checked = False
        End If
    End If
    
    Call OnlyNums(KeyAscii)
    
End Sub
