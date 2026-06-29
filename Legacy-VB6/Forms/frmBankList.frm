VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmBankList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Profile"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   8925
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   60
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   6480
      Left            =   3555
      TabIndex        =   16
      Top             =   570
      Width           =   5325
      Begin VB.TextBox txtFields 
         DataField       =   "Remittance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Index           =   10
         Left            =   75
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Tag             =   "Remittance"
         Top             =   4845
         Width           =   5175
      End
      Begin VB.TextBox txtFields 
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
         Index           =   9
         Left            =   1845
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   37
         Top             =   6105
         Visible         =   0   'False
         Width           =   3390
      End
      Begin VB.TextBox txtFields 
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
         Index           =   8
         Left            =   1845
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   36
         Top             =   6105
         Width           =   3390
      End
      Begin VB.TextBox txtFields 
         DataField       =   "AccNo"
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
         Index           =   7
         Left            =   3510
         MaxLength       =   20
         TabIndex        =   31
         Top             =   1755
         Width           =   1680
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Manager"
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
         Index           =   5
         Left            =   2670
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "Contact Person Name"
         Top             =   2790
         Width           =   2535
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Phone2"
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
         Index           =   2
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2265
         Width           =   1680
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Phone1"
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
         Index           =   1
         Left            =   75
         MaxLength       =   20
         TabIndex        =   4
         Top             =   2265
         Width           =   1680
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Fax"
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
         Index           =   3
         Left            =   3510
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2265
         Width           =   1680
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Email"
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
         Index           =   4
         Left            =   75
         MaxLength       =   100
         TabIndex        =   7
         Top             =   2790
         Width           =   2550
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Index           =   6
         Left            =   75
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Tag             =   "Address"
         Top             =   3390
         Width           =   5175
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Branch"
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
         Index           =   0
         Left            =   75
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "Contact Person Name"
         Top             =   1755
         Width           =   3390
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Swift Code"
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
         Index           =   12
         Left            =   90
         TabIndex        =   40
         Top             =   4620
         Width           =   885
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fin. Account Name :"
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
         Index           =   10
         Left            =   90
         TabIndex        =   38
         Top             =   6105
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Financial Account # :"
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
         Index           =   9
         Left            =   90
         TabIndex        =   35
         Top             =   6105
         Width           =   1740
      End
      Begin MSForms.ComboBox cmbCompany 
         Height          =   285
         Left            =   75
         TabIndex        =   34
         Top             =   405
         Width           =   5130
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "9049;503"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company :"
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
         Index           =   8
         Left            =   105
         TabIndex        =   33
         Top             =   165
         Width           =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   1
         X1              =   0
         X2              =   5270
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         Index           =   0
         X1              =   0
         X2              =   5270
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account No."
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
         Left            =   3540
         TabIndex        =   32
         Top             =   1545
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         Index           =   5
         X1              =   30
         X2              =   5300
         Y1              =   1395
         Y2              =   1395
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         Index           =   4
         X1              =   30
         X2              =   5300
         Y1              =   1410
         Y2              =   1410
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
         Top             =   810
         Width           =   1035
      End
      Begin MSForms.ComboBox cmbBank 
         DataField       =   "Bank"
         Height          =   285
         Left            =   75
         TabIndex        =   1
         Top             =   1020
         Width           =   2985
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "5265;503"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
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
         Top             =   810
         Width           =   420
      End
      Begin MSForms.ComboBox cmbCity 
         DataField       =   "City"
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   1020
         Width           =   2085
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3678;503"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IBAN No."
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
         Top             =   2580
         Width           =   705
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
         Top             =   2055
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
         Top             =   1545
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
         Top             =   2055
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
         Top             =   2055
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
         Top             =   2580
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
         Top             =   3165
         Width           =   690
      End
   End
   Begin MSComctlLib.TreeView Tv 
      Height          =   6915
      Left            =   45
      TabIndex        =   0
      Top             =   660
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   12197
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
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
      Top             =   6990
      Width           =   5310
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3975
         TabIndex        =   13
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1395
         TabIndex        =   11
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2685
         TabIndex        =   12
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&New "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Top             =   6990
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
End
Attribute VB_Name = "frmBankList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents rsBank  As ADODB.Recordset
Attribute rsBank.VB_VarHelpID = -1
Dim Bank As String, City As String
Dim rsBank1 As ADODB.Record
Attribute rsBank1.VB_VarHelpID = -1
Private Sub cmbBank_Change()
    Call cmbBank_Click
End Sub

Private Sub cmbBank_Click()
    Call AddToCombo(cmbCity, "Distinct City", "BankList", "Order By City")
End Sub

Private Sub GenerateAccNo()
    
    Dim strBankName As String, strAccName As String
    strBankName = cmbBank.Text
    If strBankName <> "" Then
        If InStr(strBankName, " ") > 0 Then
            strAccName = Left(strBankName, InStr(strBankName, " ") - 1)
        Else
            strAccName = strBankName
        End If
    End If
    strAccName = strAccName & " " & txtfields(7)
    txtfields(8) = GetNextAccno(strAccName, "15", "15-002")
    txtfields(9) = strAccName
    
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

Private Sub cmdDelete_Click()
    
    MsgBox "This option is restricted!!!", vbInformation
    Exit Sub
   On Error GoTo DelErr
 
    If MsgBox("Are You Sure To Delete This Record", vbQuestion + vbYesNo) = vbYes Then
        With rsBank
            .delete
            .MoveNEXT
            If .EOF Then .MoveLast
        End With
    End If

    Call RefreshTV

    Exit Sub

DelErr:
    If err.Number = -2147467259 Then
        MsgBox "Cannot Delete This Record." & vbNewLine & "It Has Some Related Records In Policy Proposals, Delete Them First", vbInformation, "Delete Error"
    Else
        MsgBox "Cannot Delete Record,Following Error Occured" & vbNewLine & err.Description, vbCritical, "Error"
    End If
End Sub

Private Sub cmdEdit_Click()
    Call Ulok
End Sub

Private Sub cmdSave_Click()
    
'    If cmbCompany.MatchFound = False Then
'        MsgBox "Select Company.", vbInformation
'        cmbCompany.SetFocus
'        Exit Sub
    If Trim(cmbBank) = "" Then
        MsgBox "Select The Bank First", vbInformation
        cmbBank.SetFocus
        Exit Sub
    ElseIf Trim(cmbCity) = "" Then
        MsgBox "Select The City First", vbInformation
        cmbBank.SetFocus
        Exit Sub
    ElseIf Trim(txtfields(0)) = "" Then
        MsgBox "You Must Enter The Branch Name", vbInformation
        cmbBank.SetFocus
        Exit Sub
    End If

    On Error GoTo err
    rsBank![Bank] = cmbBank
    rsBank![City] = cmbCity
    If cmbCompany.MatchFound = True Then
        rsBank!CompanyRefID = cmbCompany.List(cmbCompany.ListIndex, 1)
    Else
        rsBank!CompanyRefID = 0
    End If
    Dim bAdding As Boolean
    If rsBank.EditMode = adEditAdd Then
        bAdding = True
    End If
    rsBank.Update
    
    If bAdding Then
        Call GenerateAccNo
        con.Execute "Insert Into BankAccounts(BankID,AccNo,Type,BankAccNo) " & _
         "Values(" & rsBank!BankID & ",'" & txtfields(8) & "','15','" & txtfields(7) & "')"
        
        con.Execute "Insert Into Accounts(AccNo,AccTitle,SubAccof,Type,OpenBal,OpenDate," & _
         "Balance,UserName,ComputerName) Values('" & txtfields(8) & "','" & txtfields(9) & _
         "','15-002','15',0,'" & Date & "',0,'" & CurrentUserName & "','" & strComputerName & "')"
    End If
    
    rsBank.Requery
    Call lok
    Call RefreshTV
    
    Exit Sub
err:
    MsgBox "Cannot Save, Following Error Occured" & vbNewLine & err.Description, vbInformation
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    Me.KeyPreview = True
    
    Call AddToCombo(cmbCompany, "CompanyName", "Companies", "Order By EntryID", , "EntryID")
    
    Call AddToCombo(cmbBank, "Distinct Bank", "BankList", "Order By Bank")
    Set rsBank = New ADODB.Recordset
    rsBank.Open "Select * from BankList Order By BankID", con, adOpenDynamic, adLockOptimistic
 
    For i = 0 To txtfields.count - 1
        Set txtfields(i).DataSource = rsBank
    Next
 
    Set cmbBank.DataSource = rsBank
    Set cmbCity.DataSource = rsBank
 
    Call RefreshTV
    Call lok
    
End Sub

Private Sub RefreshTV()

    Screen.MousePointer = vbHourglass
    TV.Nodes.Clear

    Dim rsBnk As New ADODB.Recordset
    Dim rsCty As New ADODB.Recordset
    Dim rsBch As New ADODB.Recordset

    rsBnk.Open "SELECT DISTINCT Bank FROM BankList ORDER BY Bank", con

    For i = 0 To rsBnk.RecordCount - 1
        TV.Nodes.add , , rsBnk![Bank] & "", rsBnk![Bank] & "", "BNK"
        ' Add Cities
        If rsCty.State = 1 Then rsCty.Close
        rsCty.Open "Select distinct  City from BankList Where Bank='" & rsBnk![Bank] & "'", con
        For j = 0 To rsCty.RecordCount - 1
            TV.Nodes.add rsBnk![Bank] & "", tvwChild, rsCty![City] & "CTY" & rsBnk![Bank], rsCty![City] & "", "CTY"
            If rsBch.State = 1 Then rsBch.Close
            rsBch.Open "Select * From BankList Where Bank='" & rsBnk![Bank] & "' and City='" & rsCty![City] & "'", con
            
            For k = 0 To rsBch.RecordCount - 1
                TV.Nodes.add rsCty![City] & "CTY" & rsBnk![Bank], tvwChild, rsBch![BankID] & "BCH", rsBch![Branch] & "", "BCH"
                rsBch.MoveNEXT
            Next
            rsCty.MoveNEXT
        Next
        rsBnk.MoveNEXT
    Next
    'Tv.Nodes(rsBch![BankID] & "BCH").Selected = True
    Set rsBnk = Nothing
    Set rsCty = Nothing
    Screen.MousePointer = vbDefault
    
End Sub



Private Sub Form_Unload(cancel As Integer)
Set rsBank = Nothing
End Sub

Private Sub lok()
FEdit.Enabled = False
TV.Enabled = True
FSave.Visible = False
FData.Visible = True

End Sub

Private Sub Ulok()
FEdit.Enabled = True
TV.Enabled = False
FSave.Visible = True
FData.Visible = False

End Sub

Private Sub rsBank_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    If rsBank.EOF Or rsBank.BOF Then Exit Sub
    cmbCompany.ListIndex = -1
    For i = 0 To cmbCompany.ListCount - 1
        If Val(cmbCompany.List(i, 1)) = Val(rsBank!CompanyRefID & "") Then
            cmbCompany.ListIndex = i
            Exit For
        End If
    Next
    
End Sub

Private Sub Tv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo err
    
    If Button <> 2 Or TV.Nodes.count = 0 Then Exit Sub
    
    If TV.SelectedItem.Image = "BCH" Then
        'Me.PopUpMenu mnupop
        'Dim Str As String
        'Str = ShowPopUpMenu("Define &New Bank Account in This Branch")
        
        'If Str = "Define &New Bank Account in This Branch" Then
        
        '    Dim f As New frmNewAccount
        '    With f
        '        .chkBankAcc = vbChecked
        '        Call FindInCombo(.cmbBank, Tv.SelectedItem.Parent.Parent)
        '        Call FindInCombo(.cmbBankBranch, Tv.SelectedItem)
        '        .Show 1
        '    End With
        'End If
            
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Tv_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim i As Integer
    If TV.SelectedItem.Image = "BCH" Then
        rsBank.Find "BankID=" & Val(TV.SelectedItem.key), 0, adSearchForward, 1
        Bank = TV.SelectedItem.Parent.Parent
        City = TV.SelectedItem.Parent
        cmbCompany.ListIndex = -1
        For i = 0 To cmbCompany.ListCount - 1
            If Val(cmbCompany.List(i, 1)) = Val(rsBank!CompanyRefID & "") Then
                cmbCompany.ListIndex = i
                Exit For
            End If
        Next
    ElseIf TV.SelectedItem.Image = "BNK" Then
        Bank = TV.SelectedItem
        City = ""
    ElseIf TV.SelectedItem.Image = "CTY" Then
        Bank = TV.SelectedItem.Parent
        City = TV.SelectedItem
    End If
    
End Sub

