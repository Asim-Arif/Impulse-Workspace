VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmAdvancePaymentPosting 
   BorderStyle     =   0  'None
   Caption         =   "Delete Voucher"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdvancePaymentPosting.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   7905
      Index           =   0
      Left            =   0
      ScaleHeight     =   7905
      ScaleWidth      =   7635
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "ADVANCE PAYMENT POSTING"
      Top             =   0
      Width           =   7635
      Begin VB.TextBox txtNetAmt 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   5700
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   5595
         Width           =   1860
      End
      Begin VB.CheckBox chkShowAllAccounts 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H0000AFF2&
         Height          =   300
         Left            =   30
         TabIndex        =   22
         Top             =   6435
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.TextBox txtAdviceNo 
         Height          =   315
         Left            =   135
         TabIndex        =   20
         Top             =   1530
         Width           =   1770
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   1935
         TabIndex        =   1
         Top             =   1530
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3300
         TabIndex        =   17
         Top             =   5070
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox txtAmtInPakRs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   3885
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2355
         Width           =   1995
      End
      Begin VB.TextBox txtExchRate 
         Height          =   330
         Left            =   2835
         TabIndex        =   4
         Top             =   2355
         Width           =   1020
      End
      Begin VB.TextBox txtAmtRcvd 
         Height          =   330
         Left            =   1515
         TabIndex        =   7
         Top             =   2355
         Width           =   1290
      End
      Begin MSComCtl2.DTPicker DTPosting 
         Height          =   330
         Left            =   135
         TabIndex        =   3
         Top             =   2355
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   157483011
         CurrentDate     =   39644
      End
      Begin MSComctlLib.ListView LV 
         Height          =   2190
         Left            =   30
         TabIndex        =   8
         Top             =   3360
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   3863
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   45042
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S/N"
            Object.Width           =   1085
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Detail of Deduction"
            Object.Width           =   5741
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Debit"
            Object.Width           =   2593
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Credit"
            Object.Width           =   2593
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Curr"
            Object.Width           =   1244
         EndProperty
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   4980
         Picture         =   "frmAdvancePaymentPosting.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6810
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   6300
         Picture         =   "frmAdvancePaymentPosting.frx":09A9
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6810
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AFF2&
         Height          =   195
         Index           =   16
         Left            =   4500
         TabIndex        =   24
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advice No. :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AFF2&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1305
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AFF2&
         Height          =   195
         Index           =   12
         Left            =   30
         TabIndex        =   19
         Top             =   5775
         Width           =   1245
      End
      Begin MSForms.ComboBox cmbBanks 
         Height          =   390
         Left            =   30
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   5985
         Width           =   7530
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "13282;688"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Verdana"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt In Rs. :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AFF2&
         Height          =   195
         Index           =   7
         Left            =   3900
         TabIndex        =   16
         Top             =   2115
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exch. Rate :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AFF2&
         Height          =   195
         Index           =   6
         Left            =   2835
         TabIndex        =   14
         Top             =   2115
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt Rcvd :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AFF2&
         Height          =   195
         Index           =   5
         Left            =   1515
         TabIndex        =   13
         Top             =   2130
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AFF2&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   2115
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9EBCC&
         BackStyle       =   0  'Transparent
         Caption         =   "DEDUCTIONS:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AFF2&
         Height          =   375
         Left            =   30
         TabIndex        =   11
         Top             =   2970
         Width           =   2460
      End
      Begin VB.Image ImgEuro 
         Height          =   480
         Left            =   5940
         Picture         =   "frmAdvancePaymentPosting.frx":1208
         Top             =   1395
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgDollar 
         Height          =   480
         Left            =   5940
         Picture         =   "frmAdvancePaymentPosting.frx":176B
         Top             =   1395
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AFF2&
         Height          =   195
         Index           =   1
         Left            =   4020
         TabIndex        =   10
         Top             =   1305
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AFF2&
         Height          =   195
         Index           =   0
         Left            =   1950
         TabIndex        =   9
         Top             =   1305
         Width           =   915
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   4020
         TabIndex        =   2
         Top             =   1530
         Width           =   1860
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3281;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   150
         Picture         =   "frmAdvancePaymentPosting.frx":1CAB
         Top             =   375
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmAdvancePaymentPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const lCaptionHeight = 600
Dim bMouseDown As Boolean
Dim lDownX As Long, lDownY As Long
Dim iColNo As Integer
Dim iColNoTemp As Integer
Dim lEntryID As Long
Dim strAdviceNo As String
Private Sub DrawBorder(Pic As PictureBox)

    Dim lBorderColor As Long
    lBorderColor = RGB(0, 200, 0)    'RGB(69, 140, 201) 'RGB(69, 140, 201)
    Pic.Line (0, 0)-(Pic.Width, 0), lBorderColor 'Top
    Pic.Line (0, Pic.Height - 85)-(Pic.Width, Pic.Height - 85), lBorderColor 'Bottom
    Pic.Line (0, 0)-(0, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 85, 0)-(Pic.Width - 85, Pic.Height), lBorderColor  'Right
    
End Sub

Private Sub DrawPicCaptions()

    Dim i As Integer
    For i = 0 To PicWC.count - 1
        Call DrawCaptions(PicWC(i))
    Next
    
End Sub

Private Sub DrawCaptions(Pic As PictureBox)

    Dim lBorderColor As Long, lFillColor
    lBorderColor = RGB(70, 70, 181)
    lFillColor = RGB(70, 70, 181)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 350)-(Pic.Width, 350), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 600), lFillColor, BF
        Pic.ForeColor = vbWhite
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 180
        Pic.Print Pic.Tag
    End If
        
    Pic.Line (0, Pic.Height - 20)-(Pic.Width, Pic.Height - 20), lBorderColor 'Bottom

    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 20, 0)-(Pic.Width - 20, Pic.Height), lBorderColor  'Right

End Sub

Private Sub chkShowAllAccounts_Click()

'    Dim iBankID As Integer
'    iBankID = Label2.Tag
'    Call AddBanks(iBankID)
    
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()

    On Error GoTo err
    
    If cmbCountry.MatchFound = False Then Exit Sub
    If cmbCountry.List(cmbCountry.ListIndex, 1) = "Euro €" Then
        ImgEuro.Visible = True
        ImgDollar.Visible = False
    Else
        ImgEuro.Visible = False
        ImgDollar.Visible = True
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "Select Country,Curr,AccNo From ForeignCustomers "
    
    If cmbCust.ID <> "0" Then
        SQL = SQL & " Where CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(.AbsolutePosition - 1, 1) = !Curr & ""
            cmbCountry.List(.AbsolutePosition - 1, 2) = !AccNo & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
    If Val(txtAmtRcvd) <= 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Sub
    ElseIf Val(txtExchRate) <= 0 Then
        MsgBox "Invalid Exchange Rate.", vbInformation
        Exit Sub
    ElseIf Val(txtAmtInPakRs) <= 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Sub
    ElseIf cmbBanks.MatchFound = False Then
        MsgBox "Please Select Bank Account.", vbInformation
        Exit Sub
    End If
    
    Dim strVchrNo As String
    Call StartTrans(con)
    
    strVchrNo = AddVoucher()
    con.Execute "UPDATE FCustAdvancePayments SET VchrNo='" & strVchrNo & "' WHERE EntryID=" & lEntryID

    con.CommitTrans
    
    
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Function AddVoucher() As String

    'Post 2 Vouchers....
    'First Post FULL Amount
    AddVoucher = ""
    
    Dim VNo As String, m_DTVchr As Date
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double, Desc As String, dAmt As Double
    Dim dTotalAmt As Double
    Dim strBankAccNo As String, strCustomerAccNo As String
    Dim strVchrNo As String
    Dim i As Integer
    Dim dNetAmt As Double, dDeductionAmt As Double
    
    strBankAccNo = cmbBanks.List(cmbBanks.ListIndex, 1)
    strCustomerAccNo = cmbCountry.List(cmbCountry.ListIndex, 2)
    
    m_DTVchr = DTPosting
    Dim objVchr As New clsVouchers
    objVchr.VoucherType = "ADV"
    objVchr.VoucherDate = m_DTVchr
    
    Desc = txtAdviceNo.Text & " " & Val(txtAmtRcvd) & " " & cmbCountry.List(cmbCountry.ListIndex, 1) & " @ " & Val(txtExchRate) & " " & cmbCust.ID & "(" & cmbCountry.Text & ")"
    'Debit Bank Account
    dTotalAmt = Val(txtAmtInPakRs)
    dTotalAmt = Round(dTotalAmt, 2)
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .SubItems(4) = "USD" Then
                dDeductionAmt = (Val(.SubItems(2)) * Val(txtExchRate))
            Else
                dDeductionAmt = (Val(.SubItems(2)))
            End If
            dTotalAmt = dTotalAmt - Round(dDeductionAmt, 2)
        End With
    Next
     
      
    objVchr.AddVoucherHead strBankAccNo, dTotalAmt, Desc
    'Credit Customer Account
    objVchr.AddVoucherHead strCustomerAccNo, -Val(txtAmtInPakRs), Desc
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .SubItems(4) = "USD" Then
                dDeductionAmt = (Val(.SubItems(2)) * Val(txtExchRate))
            Else
                dDeductionAmt = (Val(.SubItems(2)))
            End If
            dDeductionAmt = Round(dDeductionAmt, 2)
            If dDeductionAmt > 0 Then
                 'Debit
                objVchr.AddVoucherHead .key, dDeductionAmt, Desc
            End If
        End With
    Next
    
    strVchrNo = objVchr.PostVoucher(False, True, , True)
    Set objVchr = Nothing
    
    con.Execute "INSERT INTO FCustAdvancePaymentsVouchers(RefID,VchrNo) VALUES(" & lEntryID & ",'" & strVchrNo & "')"
    AddVoucher = strVchrNo
    
    
    

    'AddVoucher = VNo
    Exit Function
err:
    AddVoucher = False
    MsgBox "Error In AddVoucher :" & err.Number & vbNewLine & err.Description
    
End Function

Private Sub Form_Load()
    cmbCust.ListHeight = 2600
    cmbCust.AddVals con, "Distinct CustCode", "ForeignCustomers", "CustCode", " Order By ForeignCustomers.CustCode"
    
    DTPosting = Date
     
    Call DrawPicCaptions
    bMouseDown = False
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If LV.SelectedItem.Tag = "Debit" Then
        iColNo = 3
    ElseIf LV.SelectedItem.Tag = "Credit" Then
        iColNo = 4
    Else
        iColNo = iColNoTemp
    End If
    
    If iColNo = 0 Then Exit Sub
    
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
        .Text = Val(LV.SelectedItem.SubItems(iColNo - 1))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x > LV.ColumnHeaders(3).Left And x < LV.ColumnHeaders(3).Left + LV.ColumnHeaders(3).Width Then
        iColNoTemp = 3
    ElseIf x > LV.ColumnHeaders(4).Left And x < LV.ColumnHeaders(4).Left + LV.ColumnHeaders(4).Width Then
        iColNoTemp = 4
    End If
End Sub

Private Sub PicWC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If y < 600 Then
        If Button = 1 Then
            If bMouseDown = False Then
                bMouseDown = True
                lDownX = x
                lDownY = y
            End If
        End If
    End If
End Sub

Private Sub PicWC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If bMouseDown Then
        Me.Move Me.Left + ((x - lDownX) / 15), Me.Top + ((y - lDownY) / 15)
    End If
End Sub

Private Sub PicWC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseDown = False
End Sub

Public Sub ShowMe(p_lEntryID As Long)
    
    lEntryID = p_lEntryID
    Dim strBankAccNo As String
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM FCustAdvancePayments WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        txtAdviceNo = !AdviceNo & ""
        cmbCust.ID = !CustCode & ""
        cmbCountry.Text = !Country & ""
        txtAmtRcvd.Text = Val(!Amount & "")
        txtExchRate.Text = Val(!ExchRate & "")
        txtAmtInPakRs.Text = Val(!Amount & "") * Val(!ExchRate & "")
        DTPosting = !DT
        strBankAccNo = !BankAccNo & ""
        .Close
        
        Label2.Tag = strBankAccNo
        Call AddBanks(strBankAccNo)
        
    End With
    
    Set rs = Nothing
    cmbBanks.ListIndex = 0
                
    Dim ITM As ListItem
    
'    Set ITM = LV.ListItems.add(, "44-002-03001", LV.ListItems.Count + 1)
'    ITM.Tag = "Debit"
'    ITM.SubItems(1) = "Commission"
'    ITM.SubItems(2) = "0"
'    ITM.SubItems(3) = "0"
'    ITM.SubItems(4) = "USD"
    Dim strAccNo As String
    strAccNo = getGeneralDataValue("PRC_FBCharges_AccNo")   '44-06001
    
    If strAccNo = "" Then
        MsgBox "Bank Service Charges Account not found in general data.", vbCritical
        Unload Me
        Exit Sub
    End If
    
    Set ITM = LV.ListItems.add(, strAccNo, LV.ListItems.count + 1)
    ITM.Tag = "Debit"
    ITM.SubItems(1) = "Bank Charges (F B Charges)"
    ITM.SubItems(2) = "0"
    ITM.SubItems(3) = "0"
    ITM.SubItems(4) = "USD"
    
    
    strAccNo = getGeneralDataValue("PRC_WithHolding_AccNo")   '45-23004
    
    If strAccNo = "" Then
        MsgBox "Withholding Tax Account not found in general data.", vbCritical
        Unload Me
        Exit Sub
    End If
    Set ITM = LV.ListItems.add(, strAccNo, LV.ListItems.count + 1)
    ITM.Tag = "Debit"
    ITM.SubItems(1) = "With Holding Tax"
    ITM.SubItems(2) = "0"
    ITM.SubItems(3) = "0"
    ITM.SubItems(4) = "PKR"
    
    strAccNo = getGeneralDataValue("PRC_EDS_AccNo")   '44-001
    
    If strAccNo = "" Then
        MsgBox "EDS Account not found in general data.", vbCritical
        Unload Me
        Exit Sub
    End If
    
    Set ITM = LV.ListItems.add(, strAccNo, LV.ListItems.count + 1)
    ITM.Tag = "Debit"
    ITM.SubItems(1) = "EDS"
    ITM.SubItems(2) = "0"
    ITM.SubItems(3) = "0"
    ITM.SubItems(4) = "PKR"
    
    strAccNo = getGeneralDataValue("PRC_Postage_AccNo")   '"44-02001"
    
    If strAccNo = "" Then
        MsgBox "Postage Account not found in general data.", vbCritical
        Unload Me
        Exit Sub
    End If
    
    Set ITM = LV.ListItems.add(, strAccNo, LV.ListItems.count + 1)
    ITM.Tag = "Debit"
    ITM.SubItems(1) = "Postage"
    ITM.SubItems(2) = "0"
    ITM.SubItems(3) = "0"
    ITM.SubItems(4) = "PKR"
    
    strAccNo = getGeneralDataValue("PRC_BSC_AccNo")   '"44-02001"
    
    If strAccNo = "" Then
        MsgBox "Bank Service Charges Account not found in general data.", vbCritical
        Unload Me
        Exit Sub
    End If
    Set ITM = LV.ListItems.add(, strAccNo, LV.ListItems.count + 1)
    ITM.Tag = "Debit"
    ITM.SubItems(1) = "Bank Service Charges"
    ITM.SubItems(2) = "0"
    ITM.SubItems(3) = "0"
    ITM.SubItems(4) = "PKR"
    
    strAccNo = getGeneralDataValue("PRC_BOC_AccNo")   '"44-02001"
    
    If strAccNo = "" Then
        MsgBox "Bank Other Charges Account not found in general data.", vbCritical
        Unload Me
        Exit Sub
    End If
    
    Set ITM = LV.ListItems.add(, strAccNo, LV.ListItems.count + 1)
    ITM.Tag = "Debit"
    ITM.SubItems(1) = "Bank Other Charges"
    ITM.SubItems(2) = "0"
    ITM.SubItems(3) = "0"
    ITM.SubItems(4) = "PKR"
    
'    strAccNo = getGeneralDataValue("PRC_FloodSurcharge_AccNo")   '"44-02001"
'
'    If strAccNo = "" Then
'        MsgBox "Flood Surcharge Account not found in general data.", vbCritical
'        Unload Me
'        Exit Sub
'    End If
'    Set ITM = LV.ListItems.add(, strAccNo, LV.ListItems.count + 1)
'    ITM.Tag = "Debit"
'    ITM.SubItems(1) = "Flood Surcharge"
'    ITM.SubItems(2) = "0"
'    ITM.SubItems(3) = "0"
'    ITM.SubItems(4) = "PKR"
    
    Me.Show 1
    
End Sub

Private Sub txtAmtInPakRs_Change()
    Call CalculateNetAmt
End Sub

Private Sub txtamtrcvd_Change()
    CalcualteAmounts
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtEdit.Text) < 0 Then
            MsgBox "Invalid Amount.", vbInformation
            Exit Sub
        End If
        If LV.SelectedItem.Tag = "" Then
            LV.SelectedItem.SubItems(2) = 0
            LV.SelectedItem.SubItems(3) = 0
        End If
        LV.SelectedItem.SubItems(iColNo - 1) = Val(txtEdit)
        Call CalculateNetAmt
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub txtExchRate_Change()
    CalcualteAmounts
End Sub

Private Sub CalcualteAmounts()

    txtAmtInPakRs = Round(Val(txtExchRate) * Val(txtAmtRcvd), 4)
    
End Sub


Private Sub AddBanks(strBankAccNo As String)
    
    Dim strCondition As String, strSQL As String
    
'    If chkShowAllAccounts.Value = vbChecked Then
'        If iBankID = 7 Then
'            strSQL = "SELECT AccNo,AccTitle FROM VBankAccounts WHERE BankID=" & iBankID & " AND AccNo NOT IN('15-002-002-18004')"
'        ElseIf iBankID = 8 Then
'            strSQL = "SELECT AccNo,AccTitle FROM VBankAccounts WHERE BankID=" & iBankID
'        Else
'            strSQL = "SELECT AccNo,AccTitle FROM VBankAccounts WHERE BankID=" & iBankID
'        End If
'    Else
'        If iBankID = 7 Then
'            strSQL = "SELECT AccNo,AccTitle FROM VBankAccounts WHERE AccNo IN('15-002-002-18002')"
'        ElseIf iBankID = 8 Then
'            strSQL = "SELECT AccNo,AccTitle FROM VBankAccounts WHERE AccNo='15-002-002-21001'"
'        Else
'            strSQL = "SELECT AccNo,AccTitle FROM VBankAccounts WHERE BankID=" & iBankID
'        End If
'    End If
    strSQL = "SELECT AccNo,AccTitle FROM VBankAccounts WHERE AccNo='" & strBankAccNo & "'"
    Dim rs As New ADODB.Recordset
    With rs
        .Open strSQL, con, adOpenForwardOnly, adLockReadOnly
        cmbBanks.Clear
        Do Until .EOF
            cmbBanks.AddItem !AccTitle & ""
            cmbBanks.List(cmbBanks.ListCount - 1, 1) = !AccNo & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub CalculateNetAmt()

    Dim dTotalDedAmt As Double
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .SubItems(4) = "USD" Then
                dTotalDedAmt = dTotalDedAmt + (Val(.SubItems(2)) * Val(txtExchRate))
            Else
                dTotalDedAmt = dTotalDedAmt + (Val(.SubItems(2)))
            End If
        End With
    Next
    txtNetAmt = Val(txtAmtInPakRs) - dTotalDedAmt
    
End Sub
