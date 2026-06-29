VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmNewCustomPList 
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12480
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
   Icon            =   "frmNewCustomPList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   12480
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
      TabIndex        =   15
      Top             =   30
      Width           =   12375
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
         TabIndex        =   23
         Top             =   1350
         Visible         =   0   'False
         Width           =   11925
         Begin VB.TextBox txtQty 
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1470
            MaxLength       =   150
            TabIndex        =   6
            Top             =   240
            Width           =   945
         End
         Begin VB.TextBox txtTo 
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   855
            MaxLength       =   150
            TabIndex        =   5
            Top             =   240
            Width           =   600
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
            Left            =   11370
            Picture         =   "frmNewCustomPList.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Add This Entry To List."
            Top             =   15
            UseMaskColor    =   -1  'True
            Width           =   540
         End
         Begin VB.TextBox txtFrom 
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   0
            MaxLength       =   150
            TabIndex        =   4
            Top             =   240
            Width           =   600
         End
         Begin MSForms.ComboBox cmbItemDesc 
            Height          =   285
            Left            =   6015
            TabIndex        =   8
            Top             =   240
            Width           =   5295
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "9340;503"
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
            Caption         =   " to"
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
            Height          =   285
            Index           =   3
            Left            =   585
            TabIndex        =   28
            Top             =   240
            Width           =   285
         End
         Begin MSForms.ComboBox cmbItemID 
            Height          =   285
            Left            =   2430
            TabIndex        =   7
            Top             =   240
            Width           =   3555
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "6271;503"
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
            Caption         =   "        Carton #             Qnty        Item ID                                                                 Item Description"
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
            TabIndex        =   24
            Top             =   0
            Width           =   11310
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
         Top             =   1035
         Width           =   1005
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
         Width           =   12375
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Packing List"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   390
            Index           =   8
            Left            =   75
            TabIndex        =   20
            Top             =   165
            Width           =   12255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Packing List"
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
            Width           =   12240
         End
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
         Left            =   11745
         Picture         =   "frmNewCustomPList.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Remove This Entry From List."
         Top             =   1965
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   540
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6495
         Left            =   285
         TabIndex        =   10
         Top             =   1935
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   11456
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Carton #"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Qnty"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   8035
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
         Left            =   6405
         TabIndex        =   16
         Top             =   8760
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
      Begin MSForms.ComboBox cmbCustomInvoice 
         Height          =   285
         Left            =   4830
         TabIndex        =   3
         Top             =   1035
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount :"
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
         Height          =   270
         Index           =   1
         Left            =   8400
         TabIndex        =   27
         Top             =   8460
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount (In Words) :"
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
         Height          =   270
         Index           =   2
         Left            =   285
         TabIndex        =   26
         Top             =   8460
         Width           =   1830
      End
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   270
         TabIndex        =   25
         Top             =   9180
         Width           =   1545
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "2725;609"
         Value           =   "0"
         Caption         =   "Print On Save"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label lblTotWords 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2100
         TabIndex        =   22
         Top             =   8460
         Width           =   6315
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Customer                     Country                       Currency       Invoice No."
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
         Width           =   6270
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   9270
         TabIndex        =   17
         Top             =   8460
         Width           =   2400
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2055
         TabIndex        =   1
         Top             =   1035
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
Attribute VB_Name = "frmNewCustomPList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String
Public Sub EditOrder(TempOrdNo As String)
    
    On Error GoTo err
    strOrderNo = TempOrdNo
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Str As String
    Dim Qty As String
    
    
    With rs
        .Open "Select * From VCustomPList Where PInvoice='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        
        
        With cmbCustomInvoice
            .Clear
            .AddItem strOrderNo
            .ListIndex = 0
            .Locked = True
        End With
    
        Do Until .EOF
            If Val(![CartonTo] & "") = 0 Then
                Str = Format(![CartonFrom] & "", "00")
                Qty = ![Qty] & " Pcs"
            Else
                Str = Format(![CartonFrom] & "", "00 to ") & Format(![CartonTo] & "", "00")
                Qty = ![Qty] & " Pcs Each"
            End If
            
            
            Set ITM = LV.ListItems.Add(, Val(![CartonFrom] & "") & "'", Str)
                
            For i = 0 To cmbItemID.ListCount - 1
                If Val(cmbItemID.List(i, 1)) = ![ITemID] Then
                    ITM.Tag = i
                    cmbItemID.List(i, 3) = 0
                    Exit For
                End If
            Next i
            
            ITM.ListSubItems.Add , , Qty
            ITM.ListSubItems(1).Tag = Val(![CartonTo] & "")
            ITM.ListSubItems.Add(, , ![ITemID] & "").Tag = ![ITemID] & ""
            
            ITM.ListSubItems.Add(, , ![Description] & "").Tag = "" '![CustomInvoiceItemID] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    txtQty = 0
    
    Set rs = Nothing
    Me.Show
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
    cmbCountry.Tag = cmbCust.Text & cmbCountry
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_CustomInvoiceForPList"
        
        .Parameters("@Customer").value = cmbCountry.Tag
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
    End With
    
    cmbCustomInvoice.Clear
    With rs
        Do Until .EOF
            cmbCustomInvoice.AddItem .Fields(0) & ""
            
            .MoveNext
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

Private Sub cmbCustomInvoice_Change()
    Call cmbCustomInvoice_Click
End Sub

Private Sub cmbCustomInvoice_Click()

    On Error GoTo err
    If cmbCustomInvoice.MatchFound = False Then Exit Sub
    
    If strOrderNo = "" Then LV.ListItems.Clear
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select ItemID,Description,Qty From VCustomInvoice Where PInvoice='" & cmbCustomInvoice.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            Set ITM = LV.ListItems.Add
            .MoveNext
        Loop
        .Close
    End With
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbItemDesc_Change()
    Call cmbItemDesc_Click
End Sub

Private Sub cmbItemDesc_Click()
    If cmbItemDesc.MatchFound = False Then Exit Sub
    cmbItemID.ListIndex = cmbItemDesc.ListIndex
End Sub

Private Sub cmbItemID_Change()
    Call cmbItemID_Click
End Sub

Private Sub cmbItemID_Click()
    If cmbItemID.MatchFound = False Then Exit Sub
    cmbItemDesc.ListIndex = cmbItemID.ListIndex
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdAddItem_Click()
    
    On Error GoTo err
    
    
    If Val(txtQty) > Val(cmbItemID.List(cmbItemID.ListIndex, 3)) Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    End If
    
    If Not cmbItemID.MatchFound Or cmbItemID = "" Then
        MsgBox "Invalid Item Selected", vbInformation
        Exit Sub
    
    ElseIf Val(txtQty) = 0 Then
        MsgBox "Quantity Cannot Be Zero", vbInformation
        Exit Sub
    End If

    If Val(txtFrom) = 0 Then
        MsgBox "Invalid Carton #", vbInformation
        Exit Sub
    End If
    
    Dim ITM As ListItem
    Dim Str As String
    Dim Qty As String
    Dim lTotalQty As Long
    Dim lNextCarton As Long
    
    If Val(txtTo) = 0 Then
        Str = Format(Val(txtFrom), "00")
        Qty = Val(txtQty) & " Pcs"
        lTotalQty = Val(txtQty)
        lNextCarton = Val(txtFrom) + 1
    Else
        Str = Format(Val(txtFrom), "00") & " to " & Format(Val(txtTo), "00")
        Qty = Val(txtQty) & " Pcs Each"
        lTotalQty = Val(txtQty) * (Val(txtTo) - Val(txtFrom) + 1)
        lNextCarton = Val(txtTo) + 1
    End If
    
    If lTotalQty > Val(cmbItemID.List(cmbItemID.ListIndex, 3)) Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    End If
    
    
    Set ITM = LV.ListItems.Add(, Val(txtFrom) & "'", Str)
    ITM.Tag = cmbItemID.ListIndex
    
    ITM.ListSubItems.Add , , Qty
    ITM.ListSubItems(1).Tag = Val(txtTo)
    
    ITM.ListSubItems.Add(, , cmbItemID).Tag = cmbItemID.List(cmbItemID.ListIndex, 1)    'AutoItemID
    
    ITM.ListSubItems.Add(, , cmbItemDesc).Tag = cmbItemID.List(cmbItemID.ListIndex, 2)    'CustomInvoiceItemID
    
    
    cmbItemID.List(cmbItemID.ListIndex, 3) = cmbItemID.List(cmbItemID.ListIndex, 3) - lTotalQty
    
    txtQty = 0
    
    txtFrom = lNextCarton
    txtTo = ""
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()

    If LV.ListItems.Count = 0 Then Exit Sub
    
    With LV.SelectedItem
        If InStr(.Text, "to") Then
            cmbItemID.List(.Tag, 3) = cmbItemID.List(.Tag, 3) + Val(.SubItems(1)) * Val(.ListSubItems(1).Tag - Val(.Key) + 1)
        Else
            cmbItemID.List(.Tag, 3) = cmbItemID.List(.Tag, 3) + Val(.SubItems(1))
        End If
        LV.ListItems.Remove .Index
       
    End With
    
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
      
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'HICO'")
    OrderDT = Date
    DeliveryDT = DateAdd("m", 1, Date)
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    
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


Private Sub txtQty_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Function Saved() As Boolean
    
    Saved = False
    If LV.ListItems.Count = 0 Then
        MsgBox "No Items Added To List.", vbInformation
        Exit Function
    End If
    
    For i = 0 To cmbItemID.ListCount - 1
        If Val(cmbItemID.List(i, 3)) > 0 Then
            MsgBox "Not All Of Custom Invoice Items Are Added To The List.", vbInformation
            Exit Function
        End If
    Next i
    
    Call StartTrans(con)
    '----------------------------------------------------------------
    If strOrderNo <> "" Then
        con.Execute "Delete from CustomPList Where CustomInvoice='" & _
         strOrderNo & "'"
    End If
    
    'Now Save Items
    Dim strItemCode As String
    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            con.Execute "Insert Into CustomPList(CustomInvoice,CartonFrom,CartonTo" & _
             ",Qty,ItemCode,CustomInvoiceItemID) Values('" & _
             cmbCustomInvoice & "'," & Val(.Key) & "," & _
             Val(.ListSubItems(1).Tag) & "," & Val(.SubItems(1)) & _
             ",'" & .ListSubItems(2).Tag & _
             "'," & Val(.ListSubItems(3).Tag) & ")"
        End With
    Next
      
    con.CommitTrans
    
    Saved = True
    Exit Function

err:
    MsgBox err.Description
End Function
