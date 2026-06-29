VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmMakerPV_Adv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maker Payment Voucher"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
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
   ScaleHeight     =   8985
   ScaleWidth      =   8625
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbAccNo 
      Height          =   285
      Left            =   1905
      TabIndex        =   34
      Top             =   5400
      Visible         =   0   'False
      Width           =   6555
      _ExtentX        =   11562
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
      TextBackColor   =   12640511
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
   Begin VB.CheckBox chkOtherAccount 
      Caption         =   "Show All Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1905
      TabIndex        =   32
      Top             =   5715
      Width           =   2235
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4245
      TabIndex        =   31
      Top             =   3555
      Visible         =   0   'False
      Width           =   1065
   End
   Begin ComboList.Usercontrol1 cmbMakerName 
      Height          =   360
      Left            =   1410
      TabIndex        =   0
      Top             =   150
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   635
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.CheckBox chkPayByCheque 
      Caption         =   "Pay through Cheque"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3480
      TabIndex        =   29
      Top             =   6150
      Width           =   1785
   End
   Begin VB.Frame FraBank 
      Caption         =   "Cheque Details :"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   1905
      TabIndex        =   16
      Top             =   6135
      Width           =   6465
      Begin VB.CommandButton cmdBank 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6015
         Picture         =   "frmMakerPV_Adv.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   555
         Width           =   360
      End
      Begin VB.TextBox Tbox 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   18
         Top             =   1275
         Width           =   1725
      End
      Begin VB.CheckBox ChkClear 
         Caption         =   "Clear Cheque "
         Enabled         =   0   'False
         Height          =   225
         Left            =   5040
         TabIndex        =   17
         ToolTipText     =   "Clear This Cheque Automatically"
         Top             =   1335
         Value           =   1  'Checked
         Width           =   1350
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   3480
         TabIndex        =   20
         Top             =   1275
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   169213955
         CurrentDate     =   37250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   14
         Left            =   3465
         TabIndex        =   26
         Top             =   1065
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cheque Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   13
         Left            =   1830
         TabIndex        =   25
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cheque No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   12
         Left            =   75
         TabIndex        =   24
         Top             =   1035
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   11
         Left            =   60
         TabIndex        =   23
         Top             =   330
         Width           =   960
      End
      Begin VB.Label lblBankAcc 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   45
         TabIndex        =   22
         Top             =   570
         Width           =   5925
      End
      Begin MSForms.ComboBox cmbCheckType 
         Height          =   315
         Left            =   1815
         TabIndex        =   21
         Top             =   1275
         Width           =   1650
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "2910;556"
         ListWidth       =   2822
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print When Saved"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1905
      TabIndex        =   6
      Top             =   7980
      Width           =   2190
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   1905
      TabIndex        =   7
      Top             =   2820
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dddd dd-MMM-yyyy"
      Format          =   169345027
      CurrentDate     =   37384
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2175
      Left            =   60
      TabIndex        =   30
      Top             =   555
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16777088
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S. #"
         Object.Width           =   873
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vchr #"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Total Value"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Paid"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Balance"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Pay"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSForms.ComboBox cmbCashInHand 
      Height          =   300
      Left            =   1905
      TabIndex        =   35
      Top             =   5400
      Width           =   6555
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "11562;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cash In Hand :"
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
      Left            =   360
      TabIndex        =   33
      Top             =   5415
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Balance :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   7
      Left            =   1065
      TabIndex        =   28
      Top             =   4155
      Visible         =   0   'False
      Width           =   795
   End
   Begin MSForms.TextBox txtBalance 
      Height          =   345
      Left            =   1905
      TabIndex        =   27
      Top             =   4125
      Visible         =   0   'False
      Width           =   1485
      VariousPropertyBits=   679495711
      BackColor       =   12648447
      ForeColor       =   255
      Size            =   "2619;609"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtAlreadyPaid 
      Height          =   345
      Left            =   1905
      TabIndex        =   15
      Top             =   3690
      Visible         =   0   'False
      Width           =   1485
      VariousPropertyBits=   679495711
      BackColor       =   12648447
      ForeColor       =   255
      Size            =   "2619;609"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Already Paid :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   9
      Left            =   660
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   6
      Left            =   780
      TabIndex        =   13
      Top             =   4560
      Width           =   1080
   End
   Begin MSForms.TextBox txtDescription 
      Height          =   345
      Left            =   1905
      TabIndex        =   2
      Top             =   4530
      Width           =   6555
      VariousPropertyBits=   746604571
      Size            =   "11562;609"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtAmount 
      Height          =   345
      Left            =   1905
      TabIndex        =   1
      Top             =   3285
      Visible         =   0   'False
      Width           =   1485
      VariousPropertyBits=   679495711
      BackColor       =   12648447
      ForeColor       =   12582912
      Size            =   "2619;609"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bill Amount :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   8
      Left            =   765
      TabIndex        =   12
      Top             =   3300
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.TextBox txtNetAmt 
      Height          =   345
      Left            =   1905
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5025
      Width           =   1485
      VariousPropertyBits=   679495711
      BackColor       =   16777215
      ForeColor       =   12582912
      Size            =   "2619;609"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1890
      X2              =   8280
      Y1              =   8325
      Y2              =   8325
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Net Amount Paid :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   300
      TabIndex        =   10
      Top             =   5055
      Width           =   1560
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   405
      Left            =   4110
      TabIndex        =   4
      Top             =   8445
      Width           =   2100
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "3704;714"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   405
      Left            =   1905
      TabIndex        =   3
      Top             =   8445
      Visible         =   0   'False
      Width           =   2100
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "3704;714"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      CausesValidation=   0   'False
      Height          =   405
      Left            =   6270
      TabIndex        =   5
      Top             =   8445
      Width           =   2100
      Caption         =   "Cancel    "
      PicturePosition =   327683
      Size            =   "3704;714"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Payment Date :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   540
      TabIndex        =   9
      Top             =   2850
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Maker Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   165
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1890
      X2              =   8280
      Y1              =   8340
      Y2              =   8340
   End
End
Attribute VB_Name = "frmMakerPV_Adv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strBillNo() As String
Dim lPrinterID As Long
Dim lEntryID As Long, bVendorBilling As Boolean

Private Function Saved() As Boolean

    Saved = False
     
    Dim SQL As String
    
    If cmbMakerName.MatchFound = False Then
        MsgBox "Please Select Valid Contractor.", vbInformation
        Exit Function
    ElseIf cmbMakerName.MatchFound = False Then
        MsgBox "Please Select Valid Contractor.", vbInformation
        Exit Function
    ElseIf Val(txtNetAmt) <= 0 Then
        MsgBox "Invlaid Amount.", vbInformation
        Exit Function
'    ElseIf Val(txtNetAmt) > Val(txtBalance) Then
'        MsgBox "Invalid Amount.", vbInformation
'        Exit Function
    End If
    
    If chkOtherAccount.Value = vbChecked And cmbAccNo.MatchFound = False Then
        MsgBox "Invalid Account No. !!!", vbInformation
        Exit Function
    ElseIf chkOtherAccount.Value = vbUnchecked And cmbCashInHand.MatchFound = False Then
        MsgBox "Invalid Account No. !!!", vbInformation
        Exit Function
    End If
     

    Dim strVchrNo As String
    Call StartTrans(con)
    
    'strVchrNo = AddVoucher
    If chkPayByCheque.Value = vbUnchecked Then
        strVchrNo = AddVoucher
    Else
        Dim Issued As Boolean, ChqBookNo As Double
    
        If lblBankAcc = "" Then
            MsgBox "Invalid Bank Selected.", vbInformation
            Exit Function
        ElseIf Not ValidChqNo(GetID(lblBankAcc), Tbox(2), False, ChqBookNo) Then
            If Issued Then
                MsgBox "The Cheque With Number '" & Tbox(2) & "' Has Been Issued.", vbInformation
            Else
                MsgBox "Invalid Cheque No.", vbInformation
            End If
            Tbox(2).SetFocus
            Exit Function
        End If
        Tbox(2).Tag = ChqBookNo
        Dim strDedVchrNo As String
        strVchrNo = SaveTransForCheque(strDedVchrNo)
        If strVchrNo = "" Then
            Exit Function
        End If
    End If
    
'    SQL = "INSERT INTO PrinterPayments(PrinterID,DT,AccNo,BillAmount," & _
'     "FixLoanDeduction,TempLoanDeduction,NetAmountPaid,VchrNo,FixLoanBalance) Values(" & lPrinterID & ",'" & DTPicker1 & "','" & _
'     cmbMakerName.ID & "'," & Val(txtAmount) & "," & Val(txtFixLoan) & "," & _
'     Val(txtDeductShort) & "," & Val(txtNetAmt) & ",'" & strVchrNo & "'," & Val(txtFixLoan.Tag) & ")"
'
'    con.Execute SQL
    
'    If Val(txtFixLoan) > 0 Then
'        con.Execute "INSERT INTO ContractorAmtCleared (AccNo,DT,AmtClrd,VchrNo) Values('" & cmbMakerName.ID & "','" & DTPicker1 & "'," & Val(txtFixLoan) & ",'" & strVchrNo & "')"
'    End If
'
'    If Val(txtDeductShort) > 0 Then
'        con.Execute "UPDATE ContractorAdvancesShort SET AmountCleared=Amount WHERE AccNo='" & cmbMakerName.ID & "'"
'    End If
     
    Dim i As Integer
'    For i = 0 To UBound(strBillNo)
'        con.Execute "UPDATE MakerPostedBills SET Payment_VchrNo='" & strVchrNo & "' WHERE BillNo='" & strBillNo(i) & "' AND PrinterID=" & lPrinterID
'    Next
    If bVendorBilling Then
        'con.Execute "UPDATE VenderPostedBills SET Payment_VchrNo='" & strVchrNo & "' WHERE EntryID=" & lEntryID
        'con.Execute "INSERT INTO VenderPostedBills_Payments(VPB_RefID,VchrNo,PaidAmt) VALUES(" & lEntryID & ",'" & strVchrNo & "'," & Val(txtNetAmt) & ")"
        For i = 1 To LV.ListItems.count
            With LV.ListItems(i)
                If .Checked Then
                    con.Execute "INSERT INTO VenderPostedBills_Payments(VPB_RefID,VchrNo,PaidAmt) VALUES(" & Val(.Tag) & ",'" & strVchrNo & "'," & Val(.SubItems(5)) & ")"
                End If
            End With
        Next
    Else
        'con.Execute "UPDATE MakerPostedBills SET Payment_VchrNo='" & strVchrNo & "' WHERE EntryID=" & lEntryID
        For i = 1 To LV.ListItems.count
            With LV.ListItems(i)
                If .Checked Then
                    con.Execute "INSERT INTO MakerPostedBills_Payments(MPB_RefID,VchrNo,PaidAmt) VALUES(" & Val(.Tag) & ",'" & strVchrNo & "'," & Val(.SubItems(5)) & ")"
                End If
            End With
        Next
    End If
    con.CommitTrans
    
    Saved = True
    If chkPrint = vbChecked Then Call PrintVoucher(MainForm.cr1, strVchrNo, True)
    SQL = ""
    
    Exit Function
err:
    MsgBox err.Description
End Function

Private Function AddVoucher() As String
    
    'Store Null To Return on an Unsuccessful Entry
    Dim VNo As String, m_DTVchr As Date
    m_DTVchr = DTPicker1.Value
    VNo = getNextVchrNo(m_DTVchr, "CPV")
    AddVoucher = ""
    Dim i As Integer
    
    Dim lRefID As Long
    
    Dim dNetAmt As Double, dLongTerm As Double
    dNetAmt = Val(txtNetAmt.Text)
    dLongTerm = 0 'Val(txtFixLoan.Text)
    
    Dim VenderAccNo As String
    
    VenderAccNo = cmbMakerName.ID
     
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double, Desc As String
    Dim strCashInHand As String
    
    'CHANGING FOR ADDIND ALL ACCOUNT OPTION
    'strCashInHand = "15-001-03001"
    
    If chkOtherAccount.Value = vbChecked Then
        strCashInHand = cmbAccNo.ID
    ElseIf chkOtherAccount.Value = vbUnchecked Then
        strCashInHand = cmbCashInHand.List(cmbCashInHand.ListIndex, 1)
    End If

    
    INextSno = getNextSno(m_DTVchr)
    'Remember Sno For Voucher Details
    SnoToUse = INextSno
      
    Bal = getBalance(strCashInHand, m_DTVchr)
    Bal = Bal - dNetAmt
    
    'Desc = Trim(txtVchrDescription) & " " & Trim(txtDesciptionFixLoan) '"Bill Bal. Pay. to F/M " & cmbmakername.Text
    Desc = txtDescription.Text
    'Credit Entry
    con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
         Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & VNo & "','" & strCashInHand & _
        "','" & Desc & "',0," & dNetAmt & "," & Bal & ",'',0)"
          
    Call adjustLedger(strCashInHand, m_DTVchr, INextSno, -dNetAmt)
    
    'Debit Entries
    
    'Maker Debit Entry....
    INextSno = INextSno + 1
    If dNetAmt > 0 Then
        Bal = getBalance(VenderAccNo, m_DTVchr)
                
        Bal = Bal + dNetAmt
        'Desc = strPostingNo
        con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & m_DTVchr & "','" & _
        VNo & "','" & VenderAccNo & "','" & Desc & "'," & _
        dNetAmt & ",0," & Bal & ",'',0)"
        
        Call adjustLedger(VenderAccNo, m_DTVchr, INextSno, dNetAmt)
        
    End If
    Dim dAmount As Double, DTLastReconciliation As Date
'    If getLastPaymentDetails(VenderAccNo, m_DTVchr, dAmount, INextSno, DTLastReconciliation) Then
'        con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT,LastPaymentDT,LastPaymentAmt,ReconciliationDT,ApprovedBy) " & _
'         "Values('" & VNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & _
'         "','" & m_DTVchr & "'," & dAmount & ",'" & DTLastReconciliation & "','" & cmbApprovedBy.Text & "')"
'    Else
'        con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
'    End If
    con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    AddVoucher = VNo
    Exit Function
err:
    AddVoucher = False
    MsgBox "Error In AddVoucher :" & err.Number & vbNewLine & err.Description
    
End Function


Private Sub chkOtherAccount_Click()
    If chkOtherAccount.Value = vbChecked Then
        cmbAccNo.Visible = True
        cmbCashInHand.Visible = False
        cmbAccNo.SetFocus
    Else
        cmbAccNo.Visible = False
        cmbCashInHand.Visible = True
        cmbCashInHand.SetFocus
    End If
End Sub

Private Sub chkPayByCheque_Click()
    If chkPayByCheque.Value = vbChecked Then
        FraBank.Enabled = True
    Else
        FraBank.Enabled = False
    End If
End Sub



Private Sub cmbMakerName_matched()
    
    txtDescription = cmbMakerName.Text
    
    Dim rs As New ADODB.Recordset
    Dim lTotalBillAmt As Long
    Dim ITM As ListItem
    With rs
        
        If bVendorBilling Then
            .Open "SELECT EntryID,VchrNo,NetPayable,TotalAmtPaid AS TotalPaidAmount FROM VVenderPostedBillsList WHERE AccNo='" & cmbMakerName.ID & "' AND Posted=1 AND ROUND(NetPayable-ISNULL(TotalAmtPaid,0),0)>0", con, adOpenForwardOnly, adLockReadOnly
        Else
            '.Open "SELECT EntryID,VchrNo,NetPayable,TotalPaidAmount FROM VMakerPostedBillsList WHERE AccNo='" & cmbMakerName.ID & "' AND Posted=1 AND ROUND(NetPayable-ISNULL(TotalPaidAmount,0),0)>0", con, adOpenForwardOnly, adLockReadOnly
            .Open "SELECT EntryID,VchrNo,NetPayable,TotalPaidAmount FROM VMakerPostedBillsList WHERE AccNo='" & cmbMakerName.ID & "' AND ROUND(NetPayable-ISNULL(TotalPaidAmount,0),0)>0", con, adOpenForwardOnly, adLockReadOnly
        End If
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , !VchrNo & ""
            ITM.ListSubItems.add , , Val(!NetPayable & "")
            ITM.ListSubItems.add , , Val(!TotalPaidAmount & "")
            ITM.ListSubItems.add , , Round(Val(!NetPayable & "") - Val(!TotalPaidAmount & ""), 0)
            ITM.ListSubItems.add , , ""
            If Val(!EntryID & "") = lEntryID Then
                ITM.Checked = True
                ITM.SubItems(5) = Val(ITM.SubItems(4))
            End If
            .MoveNEXT
        Loop
        .Close
        'txtAmount = lTotalBillAmt
       
    End With
    Set rs = Nothing
    Call UpdateTotals
    'txtNetAmt = Val(txtAmount) 'Val(txtAmount) - Val(txtFixLoan) - Val(txtDeductShort)
    
End Sub

Private Sub cmdBank_Click()

    Dim BankID As Double
    lblBankAcc = SelectBankAcc(BankID)
    lblBankAcc.Tag = BankID
    
    'Get ChqNo
    
    Dim ChqBookNo As Double
    Tbox(2) = GetNextChqNo(GetID(lblBankAcc), ChqBookNo)
    Tbox(2).Tag = ChqBookNo
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()

    If Saved Then
        txtAmount = ""
        
        
        txtDescription = ""
        txtNetAmt = ""
        txtAlreadyPaid = ""
        cmbMakerName.Text = ""
        cmbMakerName.Text = ""
        cmbMakerName.SetFocus
        'cmbDept.SetFocus
    End If
    
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    txtBalance.Text = ""
    cmbMakerName.ListHeight = 3000
      
    DTPicker1 = GetServerDate
    DTPicker2 = DTPicker1.Value
    cmbCheckType.AddItem "Open"
    cmbCheckType.AddItem "Payees A/C Only"
    cmbCheckType.ListIndex = 0
    'Call AddToCombo(cmbApprovedBy, "ApprovalAuthority", "PaymentApprovalAuthorities", " ORDER BY EntryID")
    bVendorBilling = False
    With cmbMakerName
        .ClearVals
        .AddVals con, "MakerNo + ' ' + AccTitle", "VVendersAndMakers", "AccNo", " ORDER BY AccNo"
    End With
    
    cmbAccNo.ListHeight = 1600
    cmbAccNo.AddVals con, "'{' + AccNo  + '} ' + AccTitle", "VActiveAccounts", "AccNo"
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT AccNo,AccTitle FROM VActiveAccounts WHERE SubAccOf='" & CashInHand & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCashInHand.Clear
        Do Until .EOF
            cmbCashInHand.AddItem !AccTitle & ""
            cmbCashInHand.List(.AbsolutePosition - 1, 1) = !AccNo & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    If cmbCashInHand.ListCount > 0 Then cmbCashInHand.ListIndex = i

    
End Sub


Private Sub LV_DblClick()
    
    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.Checked = False Then Exit Sub
    
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(6).Left, LV.Top + LV.SelectedItem.Top
        .Text = Val(LV.SelectedItem.SubItems(5))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim i As Integer, dBillAmt As Double, dAlreadyPaid As Double, dBalance As Double
    Dim iCheckedCount As Integer
    iCheckedCount = 0
    
    If Item.Checked Then
        Item.SubItems(5) = Val(Item.SubItems(4))
    Else
        Item.SubItems(5) = ""
    End If
    
'    For i = 1 To LV.ListItems.count
'        With LV.ListItems(i)
'            If .Checked Then
'                dBillAmt = dBillAmt + Val(.ListSubItems(2))
'                dAlreadyPaid = dAlreadyPaid + Val(.ListSubItems(3))
'                dAmt = dAmt + Val(.ListSubItems(4))
'                iCheckedCount = iCheckedCount + 1
'            End If
'        End With
'    Next
'
'    txtAmount = dBillAmt
'    txtAlreadyPaid = dAlreadyPaid
'    txtBalance = dBalance
    Call UpdateTotals
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If Val(txtEdit) < 0 Then
            MsgBox "Invalid Amount.", vbInformation
            Exit Sub
        End If
        If Val(txtEdit) > Val(LV.SelectedItem.SubItems(4)) Then
            MsgBox "Invalid Amount.", vbInformation
            Exit Sub
        End If
            
        LV.SelectedItem.SubItems(5) = Val(txtEdit)
        txtEdit.Visible = False
        
        Call UpdateTotals
        
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub txtNetAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.Value, False)
End Sub

Public Sub ShowMe(p_lEntryID As Long, Optional p_bVendorBilling As Boolean = False)
        
    lEntryID = p_lEntryID
    bVendorBilling = p_bVendorBilling
    cmbMakerName.Enabled = False
    
    Dim lNetPayable As Long
    Dim dTotalPaid As Long, dBalance As Long
    Dim rs As New ADODB.Recordset
    With rs
        If bVendorBilling Then
            .Open "SELECT * FROM VVenderPostedBillsList WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
            cmbMakerName.ID = !AccNo & ""
            lNetPayable = Val(!NetPayable & "")
            .Close
            dTotalPaid = GetSingleDoubleValue("SUM(PaidAmt)", "VenderPostedBills_Payments", " WHERE VPB_RefID=" & lEntryID)
        Else
            .Open "SELECT * FROM VMakerPostedBillsList WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
            cmbMakerName.ID = !AccNo & ""
            lNetPayable = Val(!NetPayable & "")
            .Close
            dTotalPaid = GetSingleDoubleValue("SUM(PaidAmt)", "MakerPostedBills_Payments", " WHERE MPB_RefID=" & lEntryID)
        End If
    End With
            
    txtAmount = lNetPayable
        
    dBalance = lNetPayable - dTotalPaid
    txtAlreadyPaid = dTotalPaid
    txtBalance = dBalance
    txtNetAmt = txtBalance
    
    Me.Show 1
    
End Sub

Private Function SaveTransForCheque(ByRef strDedVchrNo As String) As String

    On Error GoTo err
    'If CheckData = False Then Exit Function

    Dim i As Integer
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    Dim RcptNo As String
    
    RcptNo = "" 'IIf(Len(Trim(Tbox(3))) = 0, "", "Rcpt # :" & Tbox(3))
    
    'Get Serial # and Balances
    INextSno = getNextSno(DTPicker1.Value, "Cheque")
    SnoToUse = INextSno
    
    'Start New Transaction
    'con.BeginTrans
    Dim bankDesc As String
    bankDesc = Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & Tbox(2)
    con.Execute "INSERT INTO Cheque (SNo,CDate,AccNo,Description,Amount,BankID,chequeno,ChqBookNo,chequeType,chequeDate,Posted,Payment,Bounced)  Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
    cmbMakerName.ID & "','" & txtDescription & " (" & bankDesc & ") '," & txtNetAmt & "," & lblBankAcc.Tag & ",'" & _
    Tbox(2) & "'," & Tbox(2).Tag & ",'" & cmbCheckType & "','" & DTPicker2 & "',0,1,0)"
    'Remove the Cheque From Cheque book
    Call IssueChq(Val(Tbox(2).Tag), Tbox(2))
    
'    With LVPymt.ListItems
'        For i = 1 To .Count
'            If .Item(i).Checked Then con.Execute "Insert into ChequeDetail Values('" & SnoToUse & "','" & Val(.Item(i).key) & "'," & Val(LVPymt.ListItems(i).SubItems(4)) & ")"
'        Next
'    End With
    
    'Commit Transaction
    'con.CommitTrans
    
    SaveTransForCheque = SavePaymentTrans(SnoToUse)

    Call PlaySaveSound
    'MsgBox "Transaction Saved Successfuly.", vbInformation
   
    Exit Function
err:
    MsgBox "Transaction Terminated Due to ERROR:" & err.Number & vbNewLine & err.Description, vbCritical
    SaveTransForCheque = ""
    con.RollbackTrans
    
End Function

Private Function SavePaymentTrans(ChqSNo As Double) As String

    Dim i As Integer
    Dim NextVchrNo As String
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim dAmount As Double
    Dim sScNo As String
    Dim sCIH As String
    Dim Bal As Double
    Dim Desc As String
    Dim AccNo As String
    Dim Dated As Date
    'Get Serial # and Balances
    Dim dLongTerm As Double
    dLongTerm = 0 'Val(txtFixLoan.Text)
    
    Dated = DTPicker1
    
    NextVchrNo = getNextVchrNo(Dated, "BPV")
    INextSno = getNextSno(Dated)
    
    SnoToUse = INextSno
    
    'Party Account's Balance
    Bal = getBalance(cmbMakerName.ID, Dated)
    dAmount = Val(txtNetAmt)
    
    sScNo = ChqSNo
    
    
    'Bank Account No.
    sCIH = GetID(lblBankAcc)
    
    Bal = Bal + dAmount
    
    'Debit Entry
    Dim StrDescription As String
    
    StrDescription = "Bill payment  " & txtDescription & " Chq #: " & Tbox(2)
    
    Dim bankDesc As String
    bankDesc = Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & Tbox(2)
    con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & cmbMakerName.ID & _
        "','" & StrDescription & "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
    
    Call adjustLedger(cmbMakerName.ID, Dated, INextSno, Val(dAmount))
    
    Dim ChqNo As String
    INextSno = INextSno + 1
    Bal = getBalance(sCIH, Dated)
    dAmount = Val(txtNetAmt)
    Bal = Bal - dAmount
    'Description
    
    
    ChqNo = Tbox(2)
           
    'Credit Entry
    
    con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & sCIH & _
        "','" & StrDescription & "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
    
    Call adjustLedger(sCIH, Dated, INextSno, -Val(dAmount))
    
'    'Maker Permanent Loan Account.. Credit
'    '\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'    If dLongTerm > 0 Then
'        Dim strMakerLoanAcc As String, strVchrNo As String
'
'        If strMakerLoanAcc = "" Then
'            MsgBox "Can't Continue!!!", vbInformation
'            con.RollbackTrans
'            Exit Function
'        End If
'
'        Dim objVchr As New clsVouchers
'        objVchr.VoucherType = "JV"
'        objVchr.VoucherDate = Dated
'        Desc = "Long Term Loan Deduction:"
'        'Debit Maker
'        objVchr.AddVoucherHead cmbmakername.ID, dLongTerm, Desc
'        'Credit Maker LongTerm
'        objVchr.AddVoucherHead strMakerLoanAcc, -dLongTerm, Desc
'        strDedVchrNo = objVchr.PostVoucher(False, True)
'
'    End If
   '\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
   
    Dim rsCopy As New ADODB.Recordset
    With rsCopy
        .CursorLocation = adUseClient
        .Open "Select * from ChequeDetail Where (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            con.Execute "Insert into VoucherDetail Values('" & ![BalancedSno] & "','" & SnoToUse & "'," & ![Amount] & ")"
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rsCopy = Nothing
    con.Execute "Update Cheque Set Posted=1 Where Sno=" & sScNo
    
    Dim myDT As Date, dpAmount As Double, DTLastReconciliation As Date
    
'    If getLastPaymentDetails(cmbMakerName.Text, myDT, dpAmount, INextSno - 1, DTLastReconciliation) Then
'        con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT,LastPaymentDT,LastPaymentAmt,ReconciliationDT) " & _
'         "Values('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & _
'         "','" & myDT & "'," & dpAmount & ",'" & DTLastReconciliation & "')"
'    Else
'        con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
'    End If
     con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    'SavePaymentTrans = True
    SavePaymentTrans = NextVchrNo
    Exit Function
err:
    'MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
    SavePaymentTrans = ""
    
End Function


Private Sub UpdateTotals()

    Dim i As Integer, dTotalPayment As Double
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                dTotalPayment = dTotalPayment + Val(.SubItems(5))
            End If
        End With
    Next
    txtNetAmt = dTotalPayment
    
End Sub
