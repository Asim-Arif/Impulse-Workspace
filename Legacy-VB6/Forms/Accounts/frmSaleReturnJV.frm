VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSaleReturnJV 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   10920
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   7185
      Left            =   45
      TabIndex        =   14
      Top             =   60
      Width           =   10755
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   10755
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Return Voucher"
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
            TabIndex        =   31
            Top             =   165
            Width           =   10605
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Return Voucher"
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
            TabIndex        =   32
            Top             =   180
            Width           =   10590
         End
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   3720
         TabIndex        =   29
         Top             =   4020
         Visible         =   0   'False
         Width           =   4320
      End
      Begin VB.TextBox Tbox 
         Height          =   315
         Index           =   2
         Left            =   6150
         TabIndex        =   3
         Top             =   1545
         Width           =   3795
      End
      Begin MSComctlLib.ListView LVPymt 
         Height          =   1770
         Left            =   345
         TabIndex        =   23
         Top             =   4245
         Visible         =   0   'False
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   3122
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Advance Sno"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E3EDF9&
         CausesValidation=   0   'False
         Height          =   285
         Left            =   6735
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   1110
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   3825
         TabIndex        =   0
         Top             =   945
         Width           =   2535
         _ExtentX        =   4471
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
         CustomFormat    =   "dddd dd/MM/yyyy"
         Format          =   22740995
         CurrentDate     =   37250
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   315
         ScaleHeight     =   5175
         ScaleWidth      =   10245
         TabIndex        =   20
         Top             =   1905
         Width           =   10245
         Begin VB.CheckBox chkPrint 
            Caption         =   "Print When Saved"
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
            Left            =   45
            TabIndex        =   33
            Top             =   4800
            Value           =   1  'Checked
            Width           =   1830
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            Height          =   1395
            Left            =   1560
            TabIndex        =   28
            Top             =   855
            Visible         =   0   'False
            Width           =   4320
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H80000004&
            Height          =   660
            Left            =   4395
            TabIndex        =   22
            Top             =   4500
            Width           =   5280
            Begin MSForms.CommandButton cmdCancel 
               Height          =   360
               Left            =   3510
               TabIndex        =   13
               Top             =   195
               Width           =   1665
               Caption         =   "Cancel            "
               PicturePosition =   327683
               Size            =   "2937;635"
               Picture         =   "frmSaleReturnJV.frx":0000
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
               TabIndex        =   11
               Top             =   195
               Width           =   1665
               Caption         =   "Save & New   "
               PicturePosition =   327683
               Size            =   "2937;635"
               Picture         =   "frmSaleReturnJV.frx":0112
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
               TabIndex        =   12
               Top             =   195
               Width           =   1665
               Caption         =   "Save & Close  "
               PicturePosition =   327683
               Size            =   "2937;635"
               Picture         =   "frmSaleReturnJV.frx":0224
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
         Begin VB.CommandButton cmdSave 
            Caption         =   "Post"
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
            Left            =   9690
            Picture         =   "frmSaleReturnJV.frx":0382
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Add This Entry To Voucher"
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   525
         End
         Begin VB.CommandButton cmdDelete 
            CausesValidation=   0   'False
            Height          =   285
            Left            =   9690
            Picture         =   "frmSaleReturnJV.frx":08B4
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Remove This Entry From Voucher"
            Top             =   600
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox Tbox 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   8250
            TabIndex        =   7
            Top             =   225
            Width           =   1410
         End
         Begin MSComctlLib.ListView LView 
            Height          =   3525
            Left            =   15
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   600
            Width           =   9660
            _ExtentX        =   17039
            _ExtentY        =   6218
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Acc #"
               Object.Width           =   2910
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Acc Title"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Amount"
               Object.Width           =   2205
            EndProperty
         End
         Begin VB.TextBox Tbox 
            Height          =   315
            Index           =   0
            Left            =   4470
            TabIndex        =   6
            Top             =   225
            Width           =   3765
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Total Sale Amount :"
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
            Height          =   195
            Index           =   6
            Left            =   6510
            TabIndex        =   25
            Top             =   4215
            Width           =   1665
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8250
            TabIndex        =   24
            Top             =   4155
            Width           =   1410
         End
         Begin MSForms.ComboBox cmbAccNo 
            Height          =   315
            Left            =   30
            TabIndex        =   4
            Top             =   225
            Width           =   1965
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "3466;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox CmbAccTitle 
            Height          =   315
            Left            =   1995
            TabIndex        =   5
            Top             =   225
            Width           =   2475
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "4366;556"
            ListWidth       =   8819
            ListStyle       =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00AFBDCB&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmSaleReturnJV.frx":0DE6
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
            Left            =   30
            TabIndex        =   21
            Top             =   0
            Width           =   9630
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Last Balance :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   1
         Left            =   6525
         TabIndex        =   27
         Top             =   975
         Width           =   1395
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   7995
         TabIndex        =   26
         Top             =   945
         Width           =   1980
      End
      Begin MSForms.ComboBox cmbCustAccTitle 
         Height          =   315
         Left            =   2295
         TabIndex        =   2
         Top             =   1545
         Width           =   3855
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6800;556"
         ListWidth       =   8819
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbCustAccNo 
         Height          =   315
         Left            =   345
         TabIndex        =   1
         Top             =   1545
         Width           =   1965
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3466;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Customer Acc. #            Customer Acc. Title                                                  Description"
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
         Left            =   345
         TabIndex        =   19
         Top             =   1305
         Width           =   9615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Voucher # :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   3
         Left            =   345
         TabIndex        =   18
         Top             =   975
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   5
         Left            =   3210
         TabIndex        =   17
         Top             =   975
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1530
         TabIndex        =   16
         Top             =   975
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmSaleReturnJV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DontDropDown As Boolean
Dim DontShowList As Boolean
Dim PermissionDenied As Boolean


Private Sub cmbcustaccNo_Change()
If DontDropDown Then Exit Sub
    If cmbCustAccNo.MatchFound Then Call cmbcustaccNo_Click Else If cmbCustAccNo <> "" Then cmbCustAccNo.DropDown
End Sub
Private Sub cmbcustaccNo_Click()
    If cmbCustAccTitle.ListIndex <> cmbCustAccNo.ListIndex Then cmbCustAccTitle.ListIndex = cmbCustAccNo.ListIndex
    Call RefreshList
    Call ShowAccBalance
End Sub
Private Sub cmbcustaccNo_Validate(Cancel As Boolean)
    If Not cmbCustAccNo.MatchFound Then cmbCustAccNo.ListIndex = cmbCustAccTitle.ListIndex
End Sub
Private Sub cmbCustAccTitle_Change()
    If cmbCustAccTitle.MatchFound Then Call cmbcustacctitle_Click 'Else If cmbCustAccTitle <> "" Then cmbCustAccTitle.DropDown
End Sub
Private Sub cmbcustacctitle_Click()
    If cmbCustAccNo.ListIndex <> cmbCustAccTitle.ListIndex Then cmbCustAccNo.ListIndex = cmbCustAccTitle.ListIndex
End Sub

Private Sub cmbCustAccTitle_LostFocus()
   List2.Visible = False
End Sub

Private Sub cmbcustacctitle_Validate(Cancel As Boolean)
If Not cmbCustAccTitle.MatchFound Then cmbCustAccTitle.ListIndex = cmbCustAccNo.ListIndex
End Sub
Private Sub cmbCustAccTitle_DropButtonClick()
   List2.Visible = False
End Sub
Private Sub cmbCustAccTitle_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDown Then
   If List2.ListIndex < List2.ListCount - 1 Then List2.ListIndex = List2.ListIndex + 1
   KeyCode = 0
   DontShowList = True
ElseIf KeyCode = vbKeyUp Then
   If List2.ListIndex > 0 Then List2.ListIndex = List2.ListIndex - 1
   KeyCode = 0
   DontShowList = True
ElseIf KeyCode = 13 And List2.Visible Then
   cmbCustAccTitle = List2
   List2.Visible = False
   cmbCustAccTitle.SetFocus
End If

End Sub

Private Sub cmbCustAccTitle_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If DontShowList Then DontShowList = False Else ShowList List2, cmbCustAccTitle, cmbCustAccTitle
End Sub



Private Sub cmbAccNo_Change()
If DontDropDown Then Exit Sub
    If cmbAccNo.MatchFound Then Call cmbAccNo_Click Else If Not cmbAccNo <> "" Then cmbAccNo.DropDown
End Sub
Private Sub cmbAccNo_Click()
    If CmbAccTitle.ListIndex <> cmbAccNo.ListIndex Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
   
End Sub
Private Sub cmbAccNo_Validate(Cancel As Boolean)
    If Not cmbAccNo.MatchFound Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
End Sub



Private Sub CmbAccTitle_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = vbKeyDown Then
   If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex + 1
   KeyCode = 0
   DontShowList = True
ElseIf KeyCode = vbKeyUp Then
   If List1.ListIndex > 0 Then List1.ListIndex = List1.ListIndex - 1
   KeyCode = 0
   DontShowList = True
ElseIf KeyCode = 13 And List1.Visible Then
   CmbAccTitle = List1
   List1.Visible = False
   CmbAccTitle.SetFocus
End If

End Sub

Private Sub CmbAccTitle_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)
 If DontShowList Then DontShowList = False Else ShowList List1, CmbAccTitle, CmbAccTitle
End Sub
Private Sub CmbAccTitle_Change()
    If CmbAccTitle.MatchFound Then Call cmbAccTitle_Click 'Else If CmbAccTitle <> "" Then CmbAccTitle.DropDown
End Sub
Private Sub cmbAccTitle_Click()
    If cmbAccNo.ListIndex <> CmbAccTitle.ListIndex Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
End Sub

Private Sub CmbAccTitle_LostFocus()
   List1.Visible = False
End Sub

Private Sub CmbAccTitle_Validate(Cancel As Boolean)
If Not CmbAccTitle.MatchFound Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
End Sub
Private Sub CmbAccTitle_DropButtonClick()
   List1.Visible = False
End Sub

Private Sub RefreshList()
Screen.MousePointer = vbHourglass
Dim rsList As New ADODB.Recordset
 
rsList.CursorLocation = adUseClient
rsList.Open "Select Sno,TotalAmount,AmountLeft  from VAccPay where (AccNo='" & cmbCustAccNo & "') and (VDate<='" & DTPicker1 & "') and ((AmountLeft>0) or (AmountLeft is null))", con, adOpenForwardOnly, adLockReadOnly
rsList.Sort = "Sno"
 
With LVPymt.ListItems
    .Clear
    Dim TotalAmt As Double
    Dim itm As ListItem
    Do Until rsList.EOF
        If IsNull(rsList![AmountLeft]) Then
            TotalAmt = Val(rsList![TotalAmount])
        Else
            TotalAmt = Val(rsList![AmountLeft] & "")
        End If
        
        
        Set itm = .Add(, rsList![Sno] & "'", Val(TotalAmt) & "")
        
        
                
        rsList.MoveNext
    Loop

End With

rsList.Close
Set rsList = Nothing


' Apply To Advances

Dim SaleAmt As Double

SaleAmt = Val(lblTotal)
For i = 1 To LVPymt.ListItems.Count

    If SaleAmt <= 0 Then Exit For
    
    SaleAmt = SaleAmt - Val(LVPymt.ListItems(i))
    LVPymt.ListItems(i).Checked = True
        
Next



Screen.MousePointer = vbDefault
End Sub


Private Sub ShowAccBalance()
'Show Previous Balance of Account
If lblBalance.Visible Then
    lblBalance.Caption = Format(getBalance(cmbCustAccNo, DTPicker1), "#,##0.00 DR;#,##0.00 CR;0.00 ")
    lblBalance.ToolTipText = "Balance of '" & cmbCustAccTitle & "' As Of " & Format(DTPicker1.Value, "dd/MM/yyyy")
End If

End Sub


Private Sub cmdCancel_Click()
    ClearLView
    Unload Me
End Sub

Private Sub CmdDelete_Click()


With LView.ListItems(LView.SelectedItem.Index)
    cmbAccNo.AddItem .Text
    CmbAccTitle.AddItem .ListSubItems(1).Text
    LView.ListItems.Remove (LView.SelectedItem.Index)
    cmdDelete.Visible = False
    If LView.ListItems.Count > 0 Then
        Call GetTotal
    Else
        lblTotal = "0.00 "
    End If
End With

End Sub


Private Sub cmdSave_Click()
    Call AddToLView
    cmbAccNo.SetFocus
End Sub



Private Sub cmdSClear_Click()
    If SaveTrans Then
    
        'If MsgBox("Do You Want To Enter The Amount Returned ?", vbQuestion + vbYesNo) = vbYes Then
        '    With frmCashPV
        '        .cmbAccNo = cmbCustAccNo
        '        .Show
        '        .SetFocus
        '    End With
        '    Unload Me
        'Else
            MsgBox "Transaction Saved successfuly.", vbInformation
            Unload Me
            Me.Show
            Me.SetFocus
        'End If
        
    End If
End Sub



Private Sub cmdSClose_Click()
    If SaveTrans Then
        'If MsgBox("Do You Want To Enter The Amount Returned ?", vbQuestion + vbYesNo) = vbYes Then
        '    With frmCashPV
        '        .cmbAccNo = cmbCustAccNo
        '        .Show
        '        .SetFocus
        '    End With
            
            
        'Else
            MsgBox "Transaction Saved successfuly.", vbInformation
            
        '   End If
        Unload Me
    End If
End Sub

Private Sub DTPicker1_Change()
        Call setDateRange(DTPicker1)
    Label1(4).Caption = getNextVchrNo(DTPicker1.Value, "JV")
End Sub

Private Sub DTPicker1_Click()
    Label1(4).Caption = getNextVchrNo(DTPicker1.Value, "JV")
End Sub

Private Sub Form_Activate()
'Check For User Access
If PermissionDenied Then
   Me.Hide
   MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
   Unload Me
   Exit Sub
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
PermissionDenied = Not UserHasAccess("SaleReturnVoucher")

    Me.KeyPreview = True
    
    DTPicker1.Value = Date
    Call DTPicker1_Click
    Call FillAccCmbs

Call FindInCombo(CmbAccTitle, "Sales of Hardware")
chkPrint = IIf(frmMain.mnuPrintOnSave.Checked, vbChecked, vbUnchecked)
End Sub

Private Sub FillAccCmbs()
    Dim AccRec As New ADODB.Recordset
    With AccRec
        .CursorLocation = adUseClient
        .Open "Select AccNo,AccTitle from Accounts as Acc1 Where AccNo Like '" & SRTACC & "%' AND Parent=0 AND Active=1 ", con, adOpenForwardOnly, adLockReadOnly
              
        cmbAccNo.Clear
        CmbAccTitle.Clear
        Do Until .EOF
            cmbAccNo.AddItem ![AccNo]
            CmbAccTitle.AddItem ![AccTitle] & ""
            .MoveNext
        Loop
        .Close
    End With
    Set AccRec = Nothing
    
    
    Call AddToCombo(cmbCustAccNo, "AccNo", "VCustomers", " Where Active=1 Order By Accno")
    Call AddToCombo(cmbCustAccTitle, "AccTitle", "VCustomers", "Where Active=1 Order By Accno")
    
End Sub

Private Function AddToLView() As Boolean

    If cmbAccNo.Text = "" Then
        MsgBox "First select Account # from the list.", vbInformation
        cmbAccNo.SetFocus
        Exit Function
    ElseIf Val(Tbox(1).Text) = 0 Then
        MsgBox "First enter Some Amount.", vbInformation
        Tbox(1).SetFocus
        Exit Function
    End If
    
    Dim LItems As ListItem
    Set LItems = LView.ListItems.Add(, cmbAccNo.Text & "", cmbAccNo.Text)
    
    LItems.ListSubItems.Add 1, , CmbAccTitle.Text & ""
    LItems.ListSubItems.Add 2, , Tbox(0).Text & ""
    
    
    LItems.ListSubItems.Add 3, , Val(Tbox(1).Text)
    
    
    DontDropDown = True
   
    cmbAccNo.RemoveItem cmbAccNo.ListIndex
    CmbAccTitle.RemoveItem CmbAccTitle.ListIndex
    cmbAccNo.Text = ""
    CmbAccTitle.Text = ""
    Call GetTotal
    DontDropDown = False
    
        
    AddToLView = True
End Function

Private Sub GetTotal()

Dim i As Integer
Dim Total As Double
Total = 0

For i = 1 To LView.ListItems.Count
    Total = Total + LView.ListItems(i).ListSubItems(3).Text
Next i

lblTotal = Total

End Sub


Private Function SaveTrans() As Boolean


    If FakeAccounts Then
        SaveTrans = SaveFakeTrans
        Exit Function
    End If


If LView.ListItems.Count = 0 Then
    MsgBox "No Transaction to save!", vbInformation
    Exit Function
End If
    
    
''''''''''''''''''''''Check For Voucher No. Used Or Not '''''''''''''''''

Dim VNo As String
VNo = getNextVchrNo(DTPicker1.Value, "JV")
If Label1(4) <> VNo Then
   MsgBox "Voucher No. : " & Label1(4) & " Has Been Used By Other Transaction." & vbNewLine & _
          "Voucher No. : " & VNo & " Will Be Used For This Transaction.", vbInformation
       Label1(4) = VNo
End If
    
    
    
    Dim i As Integer
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    
    INextSno = getNextSno(DTPicker1.Value)
    
    'Remember Sno For Voucher Details
    SnoToUse = INextSno
    
    
    Bal = getBalance(cmbCustAccNo, DTPicker1.Value)
    Bal = Bal - Val(lblTotal)
    
    'Start New Transaction
    con.BeginTrans
    
    'Creit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & Label1(4) & "','" & cmbCustAccNo & _
        "','" & Tbox(0) & "',0," & Val(lblTotal) & "," & Bal & ",'',0)"
    
    
          
   Call adjustLedger(cmbCustAccNo, DTPicker1.Value, INextSno, -Val(lblTotal))
          
    
    
    
    
    
    INextSno = INextSno + 1
    ' Debit Entries
    
    
    
    For i = 1 To LView.ListItems.Count
        Bal = getBalance(LView.ListItems(i).Text, DTPicker1.Value)
        Bal = Bal + Val(LView.ListItems(i).ListSubItems(3).Text)
        
        con.Execute "insert into Vouchers Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
         Label1(4) & "','" & LView.ListItems(i).Text & "','" & LView.ListItems(i).ListSubItems(2).Text & "'," & _
         Val(LView.ListItems(i).ListSubItems(3).Text) & ",0," & _
         Bal & ",'',0)"
         
         
         Dim Amt As Double
         Amt = Val(LView.ListItems(i).ListSubItems(3).Text)
         Call adjustLedger(LView.ListItems(i).Text, DTPicker1, INextSno, Amt)
         INextSno = INextSno + 1
    Next i
    
    
   
    
    ' Check For Advance Payments And Adjust if Present
    
    'Call RefreshList
       
    'For i = 1 To LVPymt.ListItems.Count
    '    If LVPymt.ListItems(i).Checked Then Con.Execute "Insert into VoucherDetail Values (" & Val(LVPymt.ListItems(i).key) & "," & SnoToUse & ")"
    'Next
    
    
    
    
    
    con.CommitTrans
    
    
    Call SaveFakeTrans
    
    'Print If To Be Printed
        If chkPrint = vbChecked Then Call PrintVoucher(frmMain.cr1, Label1(4), True)
        
        Call PlaySaveSound
    
     SaveTrans = True
    
End Function
Private Function SaveFakeTrans() As Boolean

    If LView.ListItems.Count = 0 Then
        MsgBox "No Transaction to save!", vbInformation
        Exit Function
    End If
        
    
    ''''''''''''''''''''''Check For Voucher No. Used Or Not '''''''''''''''''

    Dim VNo As String
    VNo = getNextFakeVchrNo(DTPicker1.Value, "JV")

    
    
    
    Dim i As Integer
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    
    INextSno = getNextFakeSno(DTPicker1.Value)
    
    'Remember Sno For Voucher Details
    SnoToUse = INextSno
    
    
    Bal = getFakeBalance(cmbCustAccNo, DTPicker1.Value)
    Bal = Bal - Val(lblTotal)
    
    'Start New Transaction
    con.BeginTrans
    
    'Creit Entry
    con.Execute "insert into Fakevouchers Values(" & INextSno & ",'" & _
        Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & VNo & "','" & cmbCustAccNo & _
        "','" & Tbox(0) & "',0," & Val(lblTotal) & "," & Bal & ",'',0)"
    
    
          
   Call adjustFakeLedger(cmbCustAccNo, DTPicker1.Value, INextSno, -Val(lblTotal))
          
    
    
    
    
    
    INextSno = INextSno + 1
    ' Debit Entries
    
    
    
    For i = 1 To LView.ListItems.Count
        Bal = getFakeBalance(LView.ListItems(i).Text, DTPicker1.Value)
        Bal = Bal + Val(LView.ListItems(i).ListSubItems(3).Text)
        
        con.Execute "insert into Fakevouchers Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
         VNo & "','" & LView.ListItems(i).Text & "','" & LView.ListItems(i).ListSubItems(2).Text & "'," & _
         Val(LView.ListItems(i).ListSubItems(3).Text) & ",0," & _
         Bal & ",'',0)"
         
         
         Dim Amt As Double
         Amt = Val(LView.ListItems(i).ListSubItems(3).Text)
         Call adjustFakeLedger(LView.ListItems(i).Text, DTPicker1, INextSno, Amt)
         INextSno = INextSno + 1
    Next i
    
    
   
    
    ' Check For Advance Payments And Adjust if Present
    
    'Call RefreshList
       
    'For i = 1 To LVPymt.ListItems.Count
    '    If LVPymt.ListItems(i).Checked Then Con.Execute "Insert into VoucherDetail Values (" & Val(LVPymt.ListItems(i).key) & "," & SnoToUse & ")"
    'Next
    
    
    
    
    
    con.CommitTrans
    
    SaveFakeTrans = True
    
End Function





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    If LView.ListItems.Count > 0 Then
        If MsgBox("There is a Transaction active." & vbNewLine & "Save or Cancel Transaction first.", vbInformation + vbOKOnly) Then
            Cancel = True
        End If
    End If
End If
End Sub

Private Sub List1_Click()
If UCase(Screen.ActiveControl.Name) <> "CMBACCTITLE" Then
   CmbAccTitle = List1
   List1.Visible = False
   CmbAccTitle.SetFocus
End If
End Sub
Private Sub List2_Click()
If UCase(Screen.ActiveControl.Name) <> "CMBCUSTACCTITLE" Then
   cmbCustAccTitle = List2
   List2.Visible = False
   cmbCustAccTitle.SetFocus
End If
End Sub
Private Sub LView_Click()
    If LView.ListItems.Count > 0 Then
        With cmdDelete
            .Move .Left, LView.Top + (LView.SelectedItem.Top)
            .Visible = True
        End With
    End If
End Sub



Private Sub LView_DblClick()
    If LView.ListItems.Count > 0 Then
        With LView.SelectedItem
            If .ListSubItems(3).Text > 0 Then
                txtEdit.Width = LView.ColumnHeaders(4).Width + 30
                txtEdit.Move LView.ColumnHeaders(4).Left + LView.Left + 30, LView.Top + .Top
                txtEdit.Text = .ListSubItems(3).Text
                txtEdit.Tag = 3
            Else
                txtEdit.Width = LView.ColumnHeaders(5).Width + 30
                txtEdit.Move LView.ColumnHeaders(5).Left + LView.Left + 30, LView.Top + .Top
                txtEdit.Text = .ListSubItems(4).Text
                txtEdit.Tag = 4
            End If
        End With
        LView.Tag = LView.SelectedItem.Index
        txtEdit.Visible = True
        txtEdit.SelStart = 0
        txtEdit.SelLength = Len(txtEdit)
        txtEdit.SetFocus
    End If
End Sub

Private Sub LView_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ERR
If LView.ListItems.Count > 0 Then
    If KeyCode = 46 Then
        Call CmdDelete_Click
    End If
End If
Exit Sub
ERR:
    
    MsgBox ERR.Description
    

End Sub

Private Sub LView_Validate(Cancel As Boolean)
On Error Resume Next
    cmdDelete.Visible = False
    txtEdit.Visible = False
End Sub

Private Sub Tbox_Change(Index As Integer)
If Index = 2 Then Tbox(0) = Tbox(2)
End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As Integer)

If Index = 1 Then
    KeyAscii = OnlyNumber(KeyAscii, True)
End If

End Sub
Private Sub ClearLView()

Dim iTotal As Integer
iTotal = LView.ListItems.Count
    For i = 1 To iTotal
        LView.ListItems.Remove (1)
    Next i
Tbox(0).Text = ""
Tbox(1).Text = ""
cmdDelete.Visible = False
Call DTPicker1_Click
End Sub


Private Sub txtAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumber(KeyAscii, True)
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

With LView.ListItems(Val(LView.Tag))
    If KeyAscii = 13 Then
        .ListSubItems(Val(txtEdit.Tag)).Text = txtEdit
        Call GetTotal
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    Else
        KeyAscii = OnlyNumber(KeyAscii, True)
    End If
End With
End Sub

Private Sub Form_Resize()
   FALL.Move (Me.ScaleWidth - FALL.Width) / 2, (Me.ScaleHeight - FALL.Height) / 2
End Sub



