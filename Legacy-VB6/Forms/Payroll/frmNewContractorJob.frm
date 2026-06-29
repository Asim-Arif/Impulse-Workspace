VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmNewContractorJob 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Job"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6750
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   1065
      TabIndex        =   15
      Top             =   120
      Width           =   3120
      _ExtentX        =   5503
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
   Begin VB.Frame Frame1 
      Caption         =   "Job Details"
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
      Height          =   2130
      Left            =   90
      TabIndex        =   6
      Top             =   480
      Width           =   6480
      Begin VB.ComboBox cmbUnit 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3450
         TabIndex        =   2
         Top             =   1155
         Width           =   1395
      End
      Begin VB.CheckBox chkInactive 
         Caption         =   "Inactive Job"
         DataField       =   "Inactive"
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
         Height          =   240
         Left            =   4815
         TabIndex        =   4
         Tag             =   "False"
         Top             =   1785
         Width           =   1470
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   2
         Left            =   4875
         TabIndex        =   3
         Top             =   1155
         Width           =   1440
         VariousPropertyBits=   679495707
         Size            =   "2540;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unit"
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
         Index           =   6
         Left            =   3480
         TabIndex        =   11
         Top             =   930
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rate"
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
         Left            =   4905
         TabIndex        =   10
         Top             =   930
         Width           =   405
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
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   930
         Width           =   960
      End
      Begin MSForms.TextBox TBox 
         Height          =   870
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   1155
         Width           =   3165
         VariousPropertyBits=   -1467987941
         MaxLength       =   255
         ScrollBars      =   2
         Size            =   "5583;1535"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblItemID 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   225
         TabIndex        =   5
         Top             =   570
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Job ID"
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
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   540
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   0
         Left            =   2010
         TabIndex        =   0
         Top             =   570
         Width           =   4305
         VariousPropertyBits=   679495707
         Size            =   "7594;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Job Name"
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
         Left            =   2055
         TabIndex        =   7
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Department :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   16
      Top             =   135
      Width           =   960
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   405
      Left            =   4965
      TabIndex        =   14
      Top             =   2715
      Width           =   1620
      Caption         =   "Cancel        "
      PicturePosition =   327683
      Size            =   "2857;714"
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
      Height          =   405
      Left            =   1605
      TabIndex        =   13
      Top             =   2715
      Width           =   1620
      Caption         =   "Save & New"
      PicturePosition =   327683
      Size            =   "2857;714"
      Picture         =   "frmNewContractorJob.frx":0000
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
      Height          =   405
      Left            =   3285
      TabIndex        =   12
      Top             =   2715
      Width           =   1620
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2857;714"
      Picture         =   "frmNewContractorJob.frx":0112
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
Attribute VB_Name = "frmNewContractorJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public add As Boolean, edit As Boolean
Dim Loaded As Boolean
Dim AccType As String
Dim AccountNo As String
Dim SubAccNo As String
Dim dontShowList As Boolean
Private Sub chkParent_Click()
If chkParent = vbChecked Then
   TBox(1) = "0.00"
   TBox(1).Enabled = False
Else
   TBox(1).Enabled = True
End If
   
End Sub


Private Sub cmbCat_DropButtonClick()
    List1.Visible = False
End Sub

Private Sub cmbCat_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

If KeyCode = vbKeyDown Then
   If List1.ListIndex < List1.ListCount - 1 Then List1.ListIndex = List1.ListIndex + 1
   KeyCode = 0
   dontShowList = True
ElseIf KeyCode = vbKeyUp Then
   If List1.ListIndex > 0 Then List1.ListIndex = List1.ListIndex - 1
   KeyCode = 0
   dontShowList = True
ElseIf KeyCode = 13 And List1.Visible Then
   cmbCat = List1
   List1.Visible = False
   cmbCat.SetFocus
End If

End Sub


Private Sub cmbCat_LostFocus()
    List1.Visible = False
End Sub




Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If SaveItem Then
    Unload Me
    Me.add = True
    Me.edit = False
    Me.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If SaveItem Then Unload Me
End Sub


Private Sub Form_Activate()
   If edit And Not Loaded Then Call showdata
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    Me.KeyPreview = True
    Call AddToCombo(cmbUnit, "Distinct Unit", "Items", "Order By Unit")
    
    With cmbDept
        .ListHeight = 1600
        .AddVals con, "Name", "Departments", "DeptID", "Where TempDept=1"
    End With
    
    lblItemID = GetNextItemID
    
End Sub

Private Function SaveItem() As Boolean

On Error GoTo err

Dim OpBalance As Double
Dim Rec As New ADODB.Recordset



'SaleAcc = getAccountNo(cmbSaleAcc)
'PchAcc = getAccountNo(cmbPchAcc)

    If add Then
        con.Execute "Insert into Items(ItemID,ItemName,Unit,Description,Rate1,Rate2,Rate3,Rate4,Rate5,Rate6,Rate7,Active) " & _
        "Values('" & lblItemID & "','" & TBox(0).Text & "','" & cmbUnit & "','" & TBox(1) & "'," & Val(TBox(2)) & "," & Val(TBox(3)) & _
        "," & Val(TBox(4)) & "," & Val(TBox(5)) & "," & Val(TBox(6)) & "," & Val(TBox(7)) & "," & Val(Val(TBox(8))) & "," & CBool(chkInactive = vbUnchecked) & ")"
    Else
        'con.Execute "Update Items Set ItemName='" & TBox(0) & "',Unit='" & cmbUnit & "',Inhand=" & TBox(2) & ",Description='" & Val(TBox(1)) & "',PchPrice=" & Val(TBox(4)) & ",SalePrice=" & Val(TBox(5)) & ",Reorder =" & Val(TBox(3)) & ",Active=" & CBool(chkInactive <> vbChecked) & " Where ItemID='" & Me.Tag & "'"
        con.Execute "Update Items Set ItemName='" & TBox(0) & "',Unit='" & cmbUnit & "',Description='" & TBox(1) & "',Rate1=" & _
        Val(TBox(2)) & ",Rate2=" & Val(TBox(3)) & ",Rate3=" & Val(TBox(4)) & ",Rate4=" & Val(TBox(5)) & ",Rate5=" & Val(TBox(6)) & _
        ",Rate6=" & Val(TBox(7)) & ",Rate7=" & Val(TBox(8)) & ",Active=" & CBool(chkInactive <> vbChecked) & " Where ItemID='" & Me.Tag & "'"
    End If
 
 SaveItem = True
 
Exit Function
err:
    MsgBox err.Description
End Function




Private Sub LblAccType_Click()

End Sub

Private Sub List1_Click()

    If UCase(Screen.ActiveControl.Name) <> "CMBCAT" Then
        cmbCat = List1
        List1.Visible = False
        cmbCat.SetFocus
    End If

End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Select Case Index
    Case 2, 3, 4, 5, 6, 7, 8
        KeyAscii = OnlyNumber(Val(KeyAscii), True)
    End Select
End Sub



Private Function GetNextItemID() As String
    Dim NextVal As String
    NextVal = GetMax("val(Right(ItemID,3))", "Items")
    GetNextItemID = Format(Val(NextVal) + 1, "000")
End Function

Private Sub showdata()
Dim rs As New ADODB.Recordset


With rs
    .Open "Select * From Items Where ITemID='" & Me.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
    lblItemID = ![ItemID] & ""
    'cmbCatID = ![CatID] & ""
    TBox(0) = ![ItemName] & ""
    TBox(1) = ![Description] & ""
    TBox(2) = Val(![Rate1] & "")
    TBox(3) = Val(![Rate2] & "")
    TBox(4) = Val(![Rate3] & "")
    'Call FindInCombo(cmbSaleAcc, ![SaleAccount])
    'Call FindInCombo(cmbPchAcc, ![PurchaseAccount])
    Call FindInCombo(cmbUnit, ![Unit] & "")
    chkInactive = IIf(![Active], vbUnchecked, vbChecked)
    .Close
End With

Set rs = Nothing

'fCat.Enabled = False
'cmbSaleAcc.Enabled = False

'cmbPchAcc.Enabled = False
'TBox(2).Enabled = False

Loaded = True
End Sub
