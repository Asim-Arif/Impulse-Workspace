VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBankList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Profile"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8925
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankList.frx":0000
            Key             =   "CTY"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankList.frx":0454
            Key             =   "BNK"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankList.frx":08A8
            Key             =   "BCH"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FEdit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   3555
      Left            =   3555
      TabIndex        =   16
      Top             =   570
      Width           =   5325
      Begin VB.TextBox txtFields 
         DataField       =   "Manager"
         Height          =   285
         Index           =   5
         Left            =   2670
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "Contact Person Name"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Phone2"
         Height          =   285
         Index           =   2
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1635
         Width           =   1680
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Phone1"
         Height          =   285
         Index           =   1
         Left            =   75
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1635
         Width           =   1680
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Fax"
         Height          =   285
         Index           =   3
         Left            =   3540
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1635
         Width           =   1680
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Email"
         Height          =   285
         Index           =   4
         Left            =   75
         MaxLength       =   100
         TabIndex        =   7
         Top             =   2160
         Width           =   2550
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Address"
         Height          =   720
         Index           =   6
         Left            =   75
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Tag             =   "Address"
         Top             =   2760
         Width           =   5175
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Branch"
         Height          =   285
         Index           =   0
         Left            =   75
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "Contact Person Name"
         Top             =   1125
         Width           =   5130
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         Index           =   5
         X1              =   30
         X2              =   5300
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   4
         X1              =   30
         X2              =   5300
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name :"
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
         Index           =   3
         Left            =   105
         TabIndex        =   29
         Top             =   180
         Width           =   1035
      End
      Begin MSForms.ComboBox cmbBank 
         DataField       =   "Bank"
         Height          =   285
         Left            =   75
         TabIndex        =   1
         Top             =   390
         Width           =   2985
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "5265;503"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City :"
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
         Index           =   4
         Left            =   3150
         TabIndex        =   28
         Top             =   180
         Width           =   420
      End
      Begin MSForms.ComboBox cmbCity 
         DataField       =   "City"
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   390
         Width           =   2085
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3678;503"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manager Name"
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
         Index           =   2
         Left            =   2685
         TabIndex        =   27
         Top             =   1950
         Width           =   1275
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone 2"
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
         Index           =   5
         Left            =   1800
         TabIndex        =   22
         Top             =   1425
         Width           =   675
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
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
         Index           =   14
         Left            =   90
         TabIndex        =   21
         Top             =   915
         Width           =   1110
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone 1"
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
         Index           =   11
         Left            =   90
         TabIndex        =   20
         Top             =   1425
         Width           =   675
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FaxNo"
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
         Index           =   7
         Left            =   3540
         TabIndex        =   19
         Top             =   1425
         Width           =   510
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
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
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   1950
         Width           =   510
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   2535
         Width           =   690
      End
   End
   Begin MSComctlLib.TreeView Tv 
      Height          =   4095
      Left            =   45
      TabIndex        =   0
      Top             =   645
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   7223
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   8880
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Profile"
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
         Index           =   0
         Left            =   3480
         TabIndex        =   25
         Top             =   135
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Profile"
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
         Height          =   390
         Index           =   1
         Left            =   3495
         TabIndex        =   26
         Top             =   150
         Width           =   1950
      End
   End
   Begin VB.Frame FData 
      Height          =   585
      Left            =   3570
      TabIndex        =   30
      Top             =   4140
      Width           =   5310
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   330
         Left            =   3975
         TabIndex        =   13
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         Height          =   330
         Left            =   1395
         TabIndex        =   11
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   330
         Left            =   2685
         TabIndex        =   12
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&New "
         Height          =   330
         Left            =   105
         TabIndex        =   10
         Top             =   180
         Width           =   1230
      End
   End
   Begin VB.Frame FSave 
      Height          =   585
      Left            =   6075
      TabIndex        =   23
      Top             =   4140
      Visible         =   0   'False
      Width           =   2805
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   105
         TabIndex        =   14
         Top             =   180
         Width           =   1260
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   330
         Left            =   1425
         TabIndex        =   15
         Top             =   180
         Width           =   1260
      End
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Visible         =   0   'False
      Begin VB.Menu mnuBankAcc 
         Caption         =   "Define &New Bank Account in This Branch"
      End
   End
End
Attribute VB_Name = "frmBankList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsBank As New ADODB.Recordset
Dim Bank As String, City As String

Private Sub cmbBank_Click()
   Call AddToCombo(cmbCity, "Distinct City", "BankList", "Order By City")
End Sub

Private Sub cmdAdd_Click()
   'Call Tv_NodeClick(Tv.SelectedItem)
   rsBank.AddNew
   cmbBank = Bank
   cmbCity = City
   Call Ulok
End Sub

Private Sub cmdCancel_Click()
   rsBank.CancelUpdate
   Call lok
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub CmdDelete_Click()
   On Error GoTo DelErr
 
If MsgBox("Are You Sure To Delete This Record", vbQuestion + vbYesNo) = vbYes Then
   With rsBank
      .Delete
      .MoveNext
      If .EOF Then .MoveLast
   End With
End If

 Call RefreshTV

Exit Sub

DelErr:
If ERR.Number = -2147467259 Then
 MsgBox "Cannot Delete This Record." & vbNewLine & "It Has Some Related Records In Policy Proposals, Delete Them First", vbInformation, "Delete Error"
Else
 MsgBox "Cannot Delete Record,Following Error Occured" & vbNewLine & ERR.Description, vbCritical, "Error"
End If

End Sub

Private Sub cmdedit_Click()
Call Ulok
End Sub

Private Sub cmdSave_Click()

If Trim(cmbBank) = "" Then
 MsgBox "Select The Bank First", vbInformation
 cmbBank.SetFocus
 Exit Sub
ElseIf Trim(cmbCity) = "" Then
 MsgBox "Select The City First", vbInformation
 cmbBank.SetFocus
 Exit Sub
ElseIf Trim(txtFields(0)) = "" Then
 MsgBox "You Must Enter The Branch Name", vbInformation
 cmbBank.SetFocus
 Exit Sub
End If


On Error GoTo ERR
 rsBank![Bank] = cmbBank
 rsBank![City] = cmbCity
 rsBank.Update
 rsBank.Requery
Call lok
Call RefreshTV
Exit Sub
ERR:

   MsgBox "Cannot Save, Following Error Occured" & vbNewLine & ERR.Description, vbInformation

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Me.KeyPreview = True
 Call AddToCombo(cmbBank, "Distinct Bank", "BankList", "Order By Bank")
 rsBank.Open "Select * from BankList Order By BankID", Con, adOpenDynamic, adLockOptimistic
 
 For i = 0 To txtFields.Count - 1
  Set txtFields(i).DataSource = rsBank
 Next
 
 Set cmbBank.DataSource = rsBank
 Set cmbCity.DataSource = rsBank
 
 Call RefreshTV
 Call lok
End Sub

Private Sub RefreshTV()
Screen.MousePointer = vbHourglass
Tv.Nodes.Clear

Dim rsBnk As New ADODB.Recordset
Dim rsCty As New ADODB.Recordset
Dim rsBch As New ADODB.Recordset

rsBnk.Open "Select Distinct Bank From BankList Order By Bank", Con

For i = 0 To rsBnk.RecordCount - 1
   Tv.Nodes.Add , , rsBnk![Bank] & "", rsBnk![Bank] & "", "BNK"
' Add Cities
   If rsCty.State = 1 Then rsCty.Close
   rsCty.Open "Select distinct  City from BankList Where Bank='" & rsBnk![Bank] & "'", Con
      
   For j = 0 To rsCty.RecordCount - 1
      Tv.Nodes.Add rsBnk![Bank] & "", tvwChild, rsCty![City] & "CTY" & rsBnk![Bank], rsCty![City] & "", "CTY"
      
      If rsBch.State = 1 Then rsBch.Close
      rsBch.Open "Select * From BankList Where Bank='" & rsBnk![Bank] & "' and City='" & rsCty![City] & "'", Con
      
      For k = 0 To rsBch.RecordCount - 1
         Tv.Nodes.Add rsCty![City] & "CTY" & rsBnk![Bank], tvwChild, rsBch![BankID] & "BCH", rsBch![Branch] & "", "BCH"
      
         rsBch.MoveNext
      Next
      
      
      rsCty.MoveNext
   Next
   rsBnk.MoveNext
Next

'Tv.Nodes(rsBch![BankID] & "BCH").Selected = True

Set rsBnk = Nothing
Set rsCty = Nothing
Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
Set rsBank = Nothing
End Sub

Private Sub lok()
FEdit.Enabled = False
Tv.Enabled = True
FSave.Visible = False
FData.Visible = True

End Sub

Private Sub Ulok()
FEdit.Enabled = True
Tv.Enabled = False
FSave.Visible = True
FData.Visible = False

End Sub



Private Sub mnuBankAcc_Click()
Dim f As New frmNewAccount
With f
   .chkBankAcc = vbChecked
   Call FindInCombo(.cmbBank, Tv.SelectedItem.Parent.Parent)
   Call FindInCombo(.cmbBankBranch, Tv.SelectedItem)
   .Show 1

End With
End Sub

Private Sub Tv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 2 Or Tv.Nodes.Count = 0 Then Exit Sub
If Tv.SelectedItem.Image = "BCH" Then Me.PopupMenu mnupop
 
End Sub

Private Sub Tv_NodeClick(ByVal Node As MSComctlLib.Node)
If Tv.SelectedItem.Image = "BCH" Then
 rsBank.Find "BankID=" & Val(Tv.SelectedItem.key), 0, adSearchForward, 1
 Bank = Tv.SelectedItem.Parent.Parent
 City = Tv.SelectedItem.Parent
ElseIf Tv.SelectedItem.Image = "BNK" Then
 Bank = Tv.SelectedItem
 City = ""
ElseIf Tv.SelectedItem.Image = "CTY" Then
 Bank = Tv.SelectedItem.Parent
 City = Tv.SelectedItem
End If
End Sub
