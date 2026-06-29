VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmAdvancePayment 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Delete Voucher"
   ClientHeight    =   7260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdvancePayment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   7260
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
      Height          =   7260
      Index           =   0
      Left            =   0
      ScaleHeight     =   7260
      ScaleWidth      =   7260
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "ADVANCE PAYMENTS"
      Top             =   0
      Width           =   7260
      Begin ComboList.Usercontrol1 cmbOrderNo 
         Height          =   285
         Left            =   2895
         TabIndex        =   23
         Top             =   2070
         Width           =   1980
         _ExtentX        =   3493
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
      Begin VB.TextBox txtExchRate 
         Height          =   315
         Left            =   2985
         TabIndex        =   19
         Top             =   2685
         Width           =   1005
      End
      Begin ComboList.Usercontrol1 cmbCountry 
         Height          =   285
         Left            =   1410
         TabIndex        =   1
         Top             =   2070
         Width           =   1485
         _ExtentX        =   2619
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
      Begin ComboList.Usercontrol1 cmbCustCode 
         Height          =   285
         Left            =   165
         TabIndex        =   0
         Top             =   2070
         Width           =   1245
         _ExtentX        =   2196
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
      Begin VB.TextBox txtAddress 
         Height          =   1860
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   3330
         Width           =   6750
      End
      Begin VB.TextBox txtAmount 
         Height          =   315
         Left            =   5895
         TabIndex        =   3
         Top             =   2040
         Width           =   1005
      End
      Begin VB.TextBox txtCurr 
         Height          =   315
         Left            =   4875
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtExportOf 
         Height          =   315
         Left            =   4020
         TabIndex        =   6
         Text            =   "SURGICAL INSTRUMENTS"
         Top             =   2685
         Width           =   2880
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   315
         Left            =   1590
         TabIndex        =   5
         Top             =   2685
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   308936707
         CurrentDate     =   39699
      End
      Begin VB.TextBox txtAdviceNo 
         Height          =   315
         Left            =   165
         TabIndex        =   4
         Top             =   2685
         Width           =   1395
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
         Left            =   4395
         Picture         =   "frmAdvancePayment.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6105
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
         Left            =   5685
         Picture         =   "frmAdvancePayment.frx":09A9
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6090
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order No. :"
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
         Index           =   10
         Left            =   2910
         TabIndex        =   24
         Top             =   1830
         Width           =   870
      End
      Begin MSForms.ComboBox cmbBanks 
         Height          =   390
         Left            =   150
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   5535
         Width           =   6750
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "11906;688"
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
         Index           =   9
         Left            =   135
         TabIndex        =   21
         Top             =   5295
         Width           =   1245
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
         Index           =   8
         Left            =   2970
         TabIndex        =   20
         Top             =   2460
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Importer's Address :"
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
         Left            =   150
         TabIndex        =   18
         Top             =   3090
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount :"
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
         Left            =   5895
         TabIndex        =   17
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency :"
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
         Left            =   4890
         TabIndex        =   15
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Export Of :"
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
         Index           =   3
         Left            =   3990
         TabIndex        =   14
         Top             =   2460
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
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
         Left            =   1605
         TabIndex        =   13
         Top             =   2460
         Width           =   495
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
         Index           =   1
         Left            =   150
         TabIndex        =   12
         Top             =   2460
         Width           =   960
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
         Index           =   0
         Left            =   1440
         TabIndex        =   11
         Top             =   1845
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Code :"
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
         Left            =   150
         TabIndex        =   10
         Top             =   1845
         Width           =   930
      End
      Begin VB.Image Image2 
         Height          =   1200
         Left            =   150
         Picture         =   "frmAdvancePayment.frx":1208
         Top             =   645
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmAdvancePayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bMouseDown As Boolean
Const lCaptionHeight = 600
Dim lDownX As Long, lDownY As Long
Dim iColNo As Integer
Dim iColNoTemp As Integer
Dim bSaved As Boolean
Dim lEntryID As Long

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
        
    'Pic.Line (0, Pic.Height - 30)-(Pic.Width, Pic.Height - 30), lBorderColor 'Bottom
    Pic.Line (0, Pic.Height - 100)-(Pic.Width, Pic.Height), lFillColor, BF
    
    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor                          'Left
    Pic.Line (Pic.Width - 30, 0)-(Pic.Width - 30, Pic.Height), lBorderColor  'Right
    
    'Pic.Line (0, Pic.Height - 500)-(50, Pic.Height), lBorderColor
'    Me.BackColor = lBorderColor
End Sub



Private Sub cmbCountry_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
    
        .Open "SELECT Curr FROM ForeignCustomers WHERE CustCode='" & cmbCustCode.ID & "' AND Country='" & cmbCountry.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        txtCurr = !Curr & ""
        .Close
        
        .Open "Select Address2 From FCustBanks WHERE CustCode='" & cmbCustCode.ID & "' AND Country='" & cmbCountry.ID & "' ORDER BY DefaultBank Desc", con, adOpenForwardOnly, adLockReadOnly

        If .EOF = False Then
            txtAddress = !Address2 & ""
        Else
            txtAddress = ""
        End If

        .Close
    End With
    Set rs = Nothing
    
    cmbOrderNo.ClearVals
    cmbOrderNo.AddVals con, "OrderNo", "FCustomerOrders", "OrderNo", " WHERE CustCode='" & cmbCustCode.ID & "' AND Country='" & cmbCountry.ID & "'"
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbCustCode_matched()

    cmbCountry.ClearVals
    cmbCountry.AddVals con, "Country", "ForeignCustomers", "Country", " WHERE CustCode='" & cmbCustCode.ID & "' ORDER By ForeignCustomers.Country"
    
    Dim strCountry As String
    strCountry = GetSingleStringValue("Country", "ForeignCustomers", " WHERE CustCode='" & cmbCustCode.ID & "'")
    cmbCountry.ID = strCountry
    
    txtCurr = GetSingleStringValue("Curr", "ForeignCustomers", " WHERE CustCode='" & cmbCustCode.ID & "'")
    
End Sub

Private Sub cmbOrderNo_matched()
    txtAmount = GetSingleDoubleValue("OrderAmt", "VFOrderList", " WHERE OrderNo='" & cmbOrderNo.ID & "'")
    txtAmount.Tag = txtAmount
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    If cmbCustCode.MatchFound = False Then
        MsgBox "Please Select Customer.", vbInformation
        Exit Sub
    ElseIf cmbCountry.MatchFound = False Then
        MsgBox "Please Select New Country.", vbInformation
        Exit Sub
    ElseIf Val(txtAmount) <= 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Sub
    ElseIf cmbBanks.MatchFound = False Then
        MsgBox "Please Select Bank.", vbInformation
        Exit Sub
'    ElseIf Val(txtAmount) > Val(txtAmount.Tag) Then
'        MsgBox "Invalid Amount.", vbInformation
'        Exit Sub
    End If
    
    Call StartTrans(con)
    
    If lEntryID = 0 Then
        con.Execute "INSERT INTO FCustAdvancePayments(DT,AdviceNo,CustCode,Country,Amount,ImporterAddress,ExportOf" & _
        ",ExchRate,BankAccNo,UserName,MachineName,OrderNo) VALUES('" & DT & "','" & txtAdviceNo & "','" & cmbCustCode.ID & "','" & cmbCountry.ID & _
        "'," & Val(txtAmount) & ",'" & txtAddress & "','" & txtExportOf & "'," & Val(txtExchRate) & ",'" & _
        cmbBanks.List(cmbBanks.ListIndex, 1) & "','" & CurrentUserName & "','" & _
        strComputerName & "','" & cmbOrderNo.ID & "')"
    Else
        con.Execute "UPDATE FCustAdvancePayments SET AdviceNo='" & txtAdviceNo & "',DT='" & DT & "',CustCode='" & cmbCustCode.ID & "',Country='" & _
         cmbCountry.ID & "',Amount=" & Val(txtAmount) & ",ImporterAddress='" & txtAddress & "',ExportOf='" & _
         txtExportOf & "',ExchRate=" & Val(txtExchRate) & ",BankAccNo='" & cmbBanks.List(cmbBanks.ListIndex, 1) & "',OrderNo='" & cmbOrderNo.ID & "' WHERE EntryID=" & lEntryID
    End If
    
    con.CommitTrans
    bSaved = True
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub Form_Load()

    cmbCustCode.ListHeight = 3200
    cmbCountry.ListHeight = 3200
    
    cmbCustCode.AddVals con, "DISTINCT CustCode", "ForeignCustomers", "CustCode", " Order By ForeignCustomers.CustCode"
    
    Call AddToCombo(cmbBanks, "AccTitle", "VBankAccounts", " ORDER BY AccTitle", , "AccNo")
    'If cmbBanks.ListCount > 0 Then cmbBanks.ListIndex = 0
    
    Call DrawPicCaptions
    bMouseDown = False
    bSaved = False
    DT = Date
    lEntryID = 0
    
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

Public Function ShowMe() As Boolean
    
    'Will hold the handle to the new region.
    Dim hRegion As Long
    '
    'This API can create a Round or Oval Shaped Region.
    hRegion = CreateRoundRectRgn(0, 0, Me.Width / 15, Me.Height / 15, 20, 20)
    '
    'Change the shape of the Form to the newly created Region.
    SetWindowRgn Me.hwnd, hRegion, True
    '
    'Show the new form with the specified shape.
    'frmRegion.Show

    Me.Show 1
    ShowMe = bSaved
    
End Function

Public Function EditMe(p_lEntryID As Long) As Boolean
    On Error GoTo err
    lEntryID = p_lEntryID
    Dim rs As New ADODB.Recordset, i As Integer
    With rs
        .Open "SELECT * FROM FCustAdvancePayments WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        cmbCustCode.ID = !CustCode
        cmbCountry.ID = Val(!CustID & "")
        txtAmount = Val(!Amount & "")
        txtExchRate = Val(!ExchRate & "")
        txtAdviceNo = !AdviceNo & ""
        DT = !DT
        txtExportOf = !ExportOf & ""
        txtAddress = !ImporterAddress & ""
        For i = 0 To cmbBanks.ListCount - 1
            If cmbBanks.List(i, 1) = !BankAccNo & "" Then
                cmbBanks.ListIndex = i
                Exit For
            End If
        Next
        .Close
    End With
    Set rs = Nothing
    Me.Show 1
    EditMe = bSaved
    
    Exit Function
err:
    MsgBox err.Description
End Function
