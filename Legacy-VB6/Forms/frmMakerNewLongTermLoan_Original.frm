VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmMakerNewLongTermLoan_Original 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maker Long Term Loan"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
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
   ScaleHeight     =   7395
   ScaleWidth      =   6705
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkPost 
      Caption         =   "Post in Financials"
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
      Left            =   4614
      TabIndex        =   36
      Top             =   3885
      Value           =   1  'Checked
      Width           =   1956
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
      TabIndex        =   35
      Top             =   3885
      Width           =   2235
   End
   Begin VB.CheckBox chkPayByCheque 
      Caption         =   "Pay through Cheque"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1905
      TabIndex        =   20
      Top             =   4470
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
      Left            =   105
      TabIndex        =   21
      Top             =   4440
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
         Left            =   6030
         Picture         =   "frmMakerNewLongTermLoan_Original.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   1275
         Width           =   1725
      End
      Begin VB.CheckBox ChkClear 
         Caption         =   "Clear Cheque "
         Enabled         =   0   'False
         Height          =   225
         Left            =   5040
         TabIndex        =   22
         ToolTipText     =   "Clear This Cheque Automatically"
         Top             =   1335
         Value           =   1  'Checked
         Width           =   1350
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   3480
         TabIndex        =   25
         Top             =   1275
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   146210819
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   330
         Width           =   960
      End
      Begin VB.Label lblBankAcc 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   45
         TabIndex        =   27
         Top             =   570
         Width           =   5925
      End
      Begin MSForms.ComboBox cmbCheckType 
         Height          =   315
         Left            =   1815
         TabIndex        =   26
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
   Begin ComboList.Usercontrol1 cmbAccNo 
      Height          =   360
      Left            =   1920
      TabIndex        =   11
      Top             =   150
      Width           =   4650
      _ExtentX        =   8202
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
   Begin ComboList.Usercontrol1 cmbAccName 
      Height          =   360
      Left            =   1920
      TabIndex        =   12
      Top             =   630
      Width           =   4650
      _ExtentX        =   8202
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
      Left            =   105
      TabIndex        =   8
      Top             =   6360
      Width           =   2160
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   635
      _Version        =   393216
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
      Format          =   157155331
      CurrentDate     =   37384
   End
   Begin ComboList.Usercontrol1 cmbAllAccounts 
      Height          =   285
      Left            =   1920
      TabIndex        =   34
      Top             =   3570
      Visible         =   0   'False
      Width           =   4650
      _ExtentX        =   8202
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
   Begin MSForms.ComboBox cmbCashInHand 
      Height          =   300
      Left            =   1920
      TabIndex        =   33
      Top             =   3570
      Width           =   4650
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "8202;529"
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
      Left            =   450
      TabIndex        =   32
      Top             =   3585
      Width           =   1425
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   555
      X2              =   6460
      Y1              =   6810
      Y2              =   6810
   End
   Begin MSForms.TextBox txtDeduct 
      Height          =   360
      Left            =   5070
      TabIndex        =   19
      Top             =   2040
      Width           =   1515
      VariousPropertyBits=   746604571
      Size            =   "2672;635"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtAmt 
      Height          =   360
      Left            =   5070
      TabIndex        =   18
      Top             =   1560
      Width           =   1515
      VariousPropertyBits=   746604571
      Size            =   "2672;635"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   1005
      Left            =   1920
      TabIndex        =   17
      Top             =   2520
      Width           =   4650
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "8202;1773"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtPrevDed 
      Height          =   360
      Left            =   1920
      TabIndex        =   16
      Top             =   2040
      Width           =   1515
      VariousPropertyBits=   746604575
      Size            =   "2672;635"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtBal 
      Height          =   360
      Left            =   1920
      TabIndex        =   15
      Top             =   1560
      Width           =   1515
      VariousPropertyBits=   746604575
      Size            =   "2672;635"
      FontName        =   "Tahoma"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Maker Account # :"
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
      Index           =   4
      Left            =   285
      TabIndex        =   14
      Top             =   195
      Width           =   1590
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
      Left            =   660
      TabIndex        =   13
      Top             =   675
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deduction Amount :"
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
      Left            =   180
      TabIndex        =   10
      Top             =   2085
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New Ded. :"
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
      Index           =   3
      Left            =   4080
      TabIndex        =   9
      Top             =   2085
      Width           =   960
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
      Left            =   795
      TabIndex        =   7
      Top             =   2460
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New Loan :"
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
      Index           =   5
      Left            =   4065
      TabIndex        =   6
      Top             =   1605
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Previous Balance :"
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
      TabIndex        =   5
      Top             =   1605
      Width           =   1575
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   435
      Left            =   2580
      TabIndex        =   2
      Top             =   6885
      Width           =   2100
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "3704;767"
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
      Height          =   435
      Left            =   540
      TabIndex        =   1
      Top             =   6885
      Width           =   1980
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "3492;767"
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
      Height          =   435
      Left            =   4740
      TabIndex        =   3
      Top             =   6885
      Width           =   1770
      Caption         =   "Cancel    "
      PicturePosition =   327683
      Size            =   "3122;767"
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
      Caption         =   "Date :"
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
      Left            =   1350
      TabIndex        =   4
      Top             =   1125
      Width           =   525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   555
      X2              =   6460
      Y1              =   6825
      Y2              =   6825
   End
End
Attribute VB_Name = "frmMakerNewLongTermLoan_Original"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function Saved() As Boolean

    Saved = False
    'On Error GoTo ERR
    Dim SQL As String
    
    If Not cmbAccName.MatchFound Then
        MsgBox "Please Select Valid Vender From The List.", vbInformation
        Exit Function
    ElseIf Val(txtAmt) = 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Function
    ElseIf Val(txtDeduct) > Val(txtAmt) And txtDeduct.Visible Then
        MsgBox "Deduction Cannot Be More Than Loan Amount.", vbInformation
        Exit Function
    ElseIf chkPost.Value = vbChecked Then
        If chkPayByCheque.Value = vbUnchecked And (cmbCashInHand.MatchFound = False And cmbAllAccounts.MatchFound = False) Then
            MsgBox "Please Select Account.", vbInformation
            Exit Function
        End If
    End If
   
    Dim strVchrNo As String
    
    If chkPost.Value = vbChecked Then
        If chkPayByCheque.Value = vbUnchecked Then
            Call StartTrans(con)
            strVchrNo = AddVoucher
        Else
            Dim ChqBookNo As Double
            If lblBankAcc = "" Then
                MsgBox "Invalid Bank Selected.", vbInformation
                Exit Function
            ElseIf Not ValidChqNo(GetID(lblBankAcc), Tbox(2), False, ChqBookNo) Then
                If Issued Then
                    MsgBox "The Cheque With Number '" & txtChqNo & "' Has Been Issued.", vbInformation
                Else
                    MsgBox "Invalid Cheque No.", vbInformation
                End If
                Tbox(2).SetFocus
                Exit Function
            End If
            Tbox(2).Tag = ChqBookNo
            Call StartTrans(con)
            strVchrNo = SaveTransForCheque()
        End If
    End If
    
    SQL = "INSERT INTO MakerAdvances(AccNo,DT,Type,Description," & _
     "Amount,DAmount,AmountCleared,AccVoucherNo) Values('" & cmbAccNo.ID & _
     "','" & DTPicker1 & "',1,'" & txtDesc & _
     "'," & Val(txtAmt) & "," & Val(txtDeduct) & ",0,'" & strVchrNo & "')"
     
    con.Execute SQL
    
    con.CommitTrans
    
    Saved = True
    'If chkPrint.Value = 1 Then Call PrintReport
    SQL = ""
    Exit Function
err:
    MsgBox err.Description

End Function

Private Function SaveTransForCheque() As String

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
    
    AccNo = cmbAccNo.Text '"13-003-16001" 'Employee Permanent Loan Account
    
    theDesc = "Long Term Advance Paid To " & cmbAccName.Text
    'Start New Transaction
    'con.BeginTrans
    bankDesc = Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & Tbox(2)
    
    con.Execute "INSERT INTO Cheque (SNo,CDate,AccNo,Description,Amount,BankID,chequeno,ChqBookNo,chequeType,chequeDate,Posted,Payment,Bounced)  Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
    AccNo & "','" & theDesc & " (" & bankDesc & ") '," & Val(txtAmt) & "," & lblBankAcc.Tag & ",'" & _
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
    'dLongTerm = Val(txtFixLoan.Text)
    AccNo = cmbAccNo.ID '"13-003-16001" 'Employee Permanent Loan Account
    Dated = DTPicker2
    
    NextVchrNo = getNextVchrNo(Dated, "BPV")
    INextSno = getNextSno(Dated)
    
    SnoToUse = INextSno
    
    'Party Account's Balance
    Bal = getBalance(AccNo, Dated)
    dAmount = Val(txtAmt)
    
    sScNo = ChqSNo
    
    theDesc = "Long Term Advance Paid To " & cmbAccName.Text
    
    'Bank Account No.
    sCIH = GetID(lblBankAcc)
    
    Bal = Bal + dAmount
    
    'Debit Entry
    bankDesc = Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & Tbox(2)
    theDesc = theDesc & " " & bankDesc
    
    con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & AccNo & _
        "','" & theDesc & "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
    
    Call adjustLedger(AccNo, Dated, INextSno, Val(dAmount))
    
    Dim ChqNo As String
    INextSno = INextSno + 1
    Bal = getBalance(sCIH, Dated)
    dAmount = Val(txtAmt)
    Bal = Bal - dAmount
    'Description
    
    Desc = "Cheque Paid To " & cmbAccName.Text & " Chq# : " & Tbox(2)
    ChqNo = Tbox(2)
           
    'Credit Entry
    
    con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & NextVchrNo & "','" & sCIH & _
        "','" & theDesc & "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
    
    Call adjustLedger(sCIH, Dated, INextSno, -Val(dAmount))
    
    Dim rsCopy As New ADODB.Recordset
    With rsCopy
        .CursorLocation = adUseClient
        .Open "SELECT * FROM ChequeDetail WHERE (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            con.Execute "INSERT INTO VoucherDetail Values('" & ![BalancedSno] & "','" & SnoToUse & "'," & ![Amount] & ")"
            .MoveNEXT
        Loop
        .Close
    End With
    Set rsCopy = Nothing
    con.Execute "Update Cheque Set Posted=1 Where Sno=" & sScNo
    
    Dim myDT As Date, dpAmount As Double, DTLastReconciliation As Date
    
    con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
'    If getLastPaymentDetails(AccNo, myDT, dpAmount, INextSno - 1, DTLastReconciliation) Then
'        con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT,ApprovedBy,LastPaymentDT,LastPaymentAmt,ReconciliationDT) " & _
'         "Values('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & _
'         "','" & cmbApprovedBy.Text & "','" & myDT & "'," & dpAmount & ",'" & DTLastReconciliation & "')"
'    Else
'        con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT,ApprovedBy) Values('" & Label1(4) & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "','" & cmbApprovedBy.Text & "')"
'    End If
     
    'SavePaymentTrans = True
    SavePaymentTrans = NextVchrNo
    Exit Function
err:
    'MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
    SavePaymentTrans = ""
End Function

Private Function AddVoucher() As String
    
    Dim strDesc As String
    strDesc = "Long Term Advance Paid To " & cmbAccName.Text & " Ded. Amt : " & Val(txtDeduct)
    Dim strDebitAccNo As String, strVenderAccNo As String
    
    If chkOtherAccount.Value = vbChecked Then
        strDebitAccNo = cmbAllAccounts.ID
    Else
        strDebitAccNo = cmbCashInHand.List(cmbCashInHand.ListIndex, 1) 'CashInHand '"15-001" '"41-003-19001"
    End If
    
    
    strVenderAccNo = cmbAccNo.Text
    Dim dValue As Double
    dValue = Val(txtAmt)
    Dim objVchr As New clsVouchers
    objVchr.VoucherDate = DTPicker1.Value
    
    If chkOtherAccount.Value = vbChecked Then
        objVchr.VoucherType = "JV"
    Else
        objVchr.VoucherType = "CPV"
    End If
'    objVchr.ApprovedBy = cmbApprovedBy.Text
'    objVchr.LastPaymentAccNo = cmbAccNo.Text
     
    objVchr.AddVoucherHead strVenderAccNo, dValue, txtDesc
    objVchr.AddVoucherHead strDebitAccNo, -dValue, txtDesc
    
    Dim strVchrNo As String
    strVchrNo = objVchr.PostVoucher(, True)
    AddVoucher = strVchrNo
    
End Function

Private Function AddVoucherOld() As String

    Dim i As Long, INextSno As Double, Bal As Double, Amt As Double
    Dim VchrNo As String, theDesc As String, AccNo As String
    
    Dim Dated As Date
    Dated = DTPicker1
    'Charge All Amount to Cash
    Amt = Val(txtAmt)
    
    If Amt = 0 Then Exit Function
    
    'ConAcc.BeginTrans
     
    'Get The Required Values
     
    INextSno = getNextSno(Dated)
    VchrNo = getNextVchrNo(Dated, "CPV")
     
    'Give The Description
    theDesc = "Long Term Advance Paid To " & cmbAccName.Text
     
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    
    AccNo = CashInHand
    
    Bal = getBalance(AccNo, Dated)
    
    Bal = Bal - Amt
        
    'Credit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
           
    
    Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
     
    INextSno = INextSno + 1
     
    '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
      
      'Determine the Account no for Staff Member
      'Long / Short Tem Advance
      
    AccNo = cmbAccNo.ID '"13-003-16001" 'IIf(LongTerm, LongAdvAccNo, ShortAdvAccNo)
     
    Bal = getBalance(AccNo, Dated)
    
    Amt = Amt
    Bal = Bal + Amt
        
      'Debit To Staff Long/Short Advance Account
      
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
     Dated & "','" & VchrNo & "','" & AccNo & _
     "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
    Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
    con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
        
    AddVoucherOld = VchrNo
    
End Function

Private Sub chkOtherAccount_Click()

    If chkOtherAccount.Value = vbChecked Then
        cmbAllAccounts.Visible = True
        cmbCashInHand.Visible = False
        cmbAllAccounts.SetFocus
    Else
        cmbAllAccounts.Visible = False
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

Private Sub cmbAccName_matched()

    On Error GoTo err
    
    cmbAccNo.ID = cmbAccName.ID
    txtVchrDescription = cmbAccName.Text
    txtDesc = cmbAccName.Text
    '''
    lVenderID = 0
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "SELECT VendID FROM VMakers WHERE AccNo='" & cmbAccName.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            lVenderID = Val(!VendID & "")
        End If
        .Close
    End With
    
    SQL = "(SELECT ISNULL(SUM(AmtClrd),0) Balance FROM MakerAmtCleared WHERE AccNo='" & cmbAccNo.ID & "')"
        
    Balance = Val(con.Execute("SELECT SUM(Amount)-" & SQL & "  FROM MakerAdvances WHERE AccNo='" & cmbAccNo.ID & "'").Fields(0) & "")
    txtBal = Balance
     
    With rs
        .Open "SELECT EntryID,DAmount FROM MakerAdvances WHERE EntryID=(SELECT MAX(EntryID) FROM MakerAdvances WHERE AccNo='" & cmbAccNo.ID & "')", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = True Then
            cmbAccName.Tag = "N.A"
        Else
            cmbAccName.Tag = !EntryID
            txtPrevDed = Val(!dAmount & "")
        End If
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbAccName_UnMatched()
    lVenderID = 0
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
         Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub Form_Load()

    cmbAccNo.ListHeight = 3000
    cmbAccName.ListHeight = 3000
    DTPicker1 = GetServerDate
    DTPicker2 = DTPicker1.Value
    'DTPicker1.Enabled = False
    cmbCheckType.AddItem "Open"
    cmbCheckType.AddItem "Payees A/C Only"
    cmbCheckType.ListIndex = 0
    
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

Public Sub ShowMe()
    
    With cmbAccNo
        .ClearVals
        .AddVals con, "AccNo", "VVendersAndMakers", "AccNo", " ORDER BY VVendersAndMakers.AccNo"
    End With
    
    With cmbAccName
        .ClearVals
        .AddVals con, "'{'+MakerNo+'} '+AccTitle", "VVendersAndMakers", "AccNo", " ORDER BY VVendersAndMakers.AccNo"
    End With
    
    cmbAllAccounts.ListHeight = 1600
    cmbAllAccounts.AddVals con, "'{' + AccNo  + '} ' + AccTitle", "VActiveAccounts", "AccNo"
    
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
    
    Me.Show 1
    
End Sub


