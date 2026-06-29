VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmVendAddPayDetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Payment Details For The Venders....."
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVendAddPayDetail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6045
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3420
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Height          =   360
      Left            =   4650
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2025
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1215
   End
   Begin VB.PictureBox PIC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Index           =   1
      Left            =   60
      ScaleHeight     =   3330
      ScaleWidth      =   7140
      TabIndex        =   9
      Top             =   30
      Visible         =   0   'False
      Width           =   7200
      Begin VB.TextBox txtChequeNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3435
         TabIndex        =   16
         Top             =   1650
         Width           =   1695
      End
      Begin VB.TextBox txtAmt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1650
         Width           =   1290
      End
      Begin VB.Label lblChequeNo 
         AutoSize        =   -1  'True
         Caption         =   "Cheque No."
         Height          =   195
         Left            =   3435
         TabIndex        =   17
         Top             =   1425
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pay By"
         Height          =   195
         Left            =   1500
         TabIndex        =   15
         Top             =   1425
         Width           =   495
      End
      Begin MSForms.ComboBox cmbPayBy 
         Height          =   285
         Left            =   1500
         TabIndex        =   14
         Top             =   1650
         Width           =   1905
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3360;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   1425
         Width           =   555
      End
      Begin MSForms.ComboBox cmbVend 
         Height          =   285
         Left            =   150
         TabIndex        =   11
         Top             =   1050
         Width           =   4965
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "8758;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblInvNo 
         AutoSize        =   -1  'True
         Caption         =   "Vender Name"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   825
         Width           =   960
      End
   End
   Begin VB.PictureBox PIC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Index           =   0
      Left            =   67
      ScaleHeight     =   3330
      ScaleWidth      =   7140
      TabIndex        =   4
      Top             =   30
      Width           =   7200
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "- Click Next To continue."
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   8
         Top             =   1905
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   $"frmVendAddPayDetail.frx":27A2
         Height          =   585
         Index           =   1
         Left            =   75
         TabIndex        =   7
         Top             =   945
         Width           =   5175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "- This wizard lets you Add payment details for the venders."
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   6
         Top             =   510
         Width           =   4260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Welcome to the Add Payment Details Wizard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   165
         Width           =   3780
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   60
      X2              =   7265
      Y1              =   3510
      Y2              =   3510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   75
      X2              =   7265
      Y1              =   3525
      Y2              =   3525
   End
End
Attribute VB_Name = "frmVendAddPayDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PicIndex As Integer
Dim iMonth As Integer
Dim iYear As Integer
Dim iAdvSal As Integer

Public Sub ShowMe(intMon As Integer, intYear As Integer, AdvSal As Integer)
    
    iMonth = intMon
    iYear = intYear
    iAdvSal = AdvSal
    Me.Show 1
    
End Sub


Private Sub cmbPayBy_Change()
    Call cmbPayBy_Click
End Sub

Private Sub cmbPayBy_Click()

    If cmbPayBy.ListIndex = 0 Then 'Cheque
        lblChequeNo.Visible = True
        txtChequeNo.Visible = True
    Else
        lblChequeNo.Visible = False
        txtChequeNo.Visible = False
    End If
    
End Sub

Private Sub cmbVend_Change()
    Call cmbVend_Click
End Sub

Private Sub cmbVend_Click()

    txtAmt = Val(cmbVend.List(cmbVend.ListIndex, 2))
    cmbPayBy.ListIndex = Val(cmbVend.List(cmbVend.ListIndex, 3))
    
    If cmbVend.List(cmbVend.ListIndex, 4) = "" And cmbPayBy.ListIndex = 0 Then
        txtChequeNo = NextChequeNo(cmbVend.ListIndex - 1) 'Val(cmbVend.List(cmbVend.ListIndex, 4))
    Else
        txtChequeNo = cmbVend.List(cmbVend.ListIndex, 4)
    End If
    
End Sub

Private Function NextChequeNo(iListIndex As Integer) As String

    Dim RetVal As String
    Dim PrevChequeNo As String
    For i = iListIndex To 0 Step -1
        If cmbVend.List(i, 3) = 0 Then
            PrevChequeNo = cmbVend.List(i, 4)
            RetVal = Left(PrevChequeNo, Len(PrevChequeNo) - 3) & Format(Val(Right(PrevChequeNo, 3)) + 1, "000")
            Exit For
        End If
    Next
    NextChequeNo = RetVal
    
End Function
Private Sub cmdBack_Click()

    If cmbVend.ListIndex > 0 Then
        cmbVend.List(cmbVend.ListIndex, 3) = Val(cmbPayBy.ListIndex)
            
        cmbVend.List(cmbVend.ListIndex, 4) = IIf(cmbPayBy.ListIndex = 0, txtChequeNo, "")
        cmbVend.ListIndex = cmbVend.ListIndex - 1
    End If
    Call setButtonStatus
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFinish_Click()

    On Error GoTo err
        
        
    cmbVend.List(cmbVend.ListIndex, 3) = Val(cmbPayBy.ListIndex)
            
    cmbVend.List(cmbVend.ListIndex, 4) = IIf(cmbPayBy.ListIndex = 0, txtChequeNo, "")
            
    'cmbVend.ListIndex = cmbVend.ListIndex + 1
    
    Call StartTrans(con)
    
    For i = 0 To cmbVend.ListCount - 1
        'If iAdvSal = 1 Then
            con.Execute "Update VendMonthlyPayments Set CashPaid=" & Val(cmbVend.List(i, 3)) & _
             ",ChequeNo='" & cmbVend.List(i, 4) & "' Where EntryID=" & Val(cmbVend.List(i, 5))
'        Else
'            con.Execute "Update VendMonthlyPayments Set CashPaid=" & Val(cmbVend.List(i, 3)) & _
'             ",ChequeNo='" & cmbVend.List(i, 4) & "' Where EntryID=" & Val(cmbVend.List(i, 5))
'        End If
    Next
    
    con.CommitTrans
    
    MsgBox "Done Successfuly.", vbInformation
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdNext_Click()
    
    Dim rs As New ADODB.Recordset
    Dim iList As Integer
    
    If PicIndex = 1 Then
        If cmbPayBy.ListIndex = 0 And Len(txtChequeNo) < 3 Then
            MsgBox "You Must Enter A Cheque No. Which is More Than 2 Digits Long.", vbInformation
            Exit Sub
        End If
        If cmbVend.ListIndex < cmbVend.ListCount - 1 Then
            cmbVend.List(cmbVend.ListIndex, 3) = Val(cmbPayBy.ListIndex)
            
            cmbVend.List(cmbVend.ListIndex, 4) = IIf(cmbPayBy.ListIndex = 0, txtChequeNo, "")
            
            cmbVend.ListIndex = cmbVend.ListIndex + 1
        End If
        Call setButtonStatus
        
    ElseIf PicIndex = 0 Then
        With rs
            .Open "Select * From VVendMonthlyPayments Where Year(DT)=" & iYear & " AND Month(DT)=" & iMonth & " AND AdvSalary=" & iAdvSal & " AND PaidAmt>0 AND Payed=1", con, adOpenForwardOnly, adLockReadOnly
            cmbVend.Clear
            Do Until .EOF
                iList = .AbsolutePosition - 1
                cmbVend.AddItem ![VenderName] & ""
                cmbVend.List(iList, 1) = ![VendID] & ""
                cmbVend.List(iList, 2) = ![PaidAmt] & ""
                cmbVend.List(iList, 3) = ![CashPaid] & ""
                cmbVend.List(iList, 4) = ![ChequeNo] & ""
                cmbVend.List(iList, 5) = ![EntryID] & ""
                .MoveNext
            Loop
            If cmbVend.ListCount > 0 Then
                cmbVend.ListIndex = 0
            End If
            .Close
        End With
        PicIndex = PicIndex + 1
        Pic(PicIndex).Visible = True
        Pic(0).Visible = False
    End If
    
End Sub

Private Sub setButtonStatus()
    
    If cmbVend.ListIndex < cmbVend.ListCount - 1 Then
        cmdNext.Enabled = True
    Else
        cmdNext.Enabled = False
        cmdFinish.Enabled = True
    End If
    
    If cmbVend.ListIndex > 0 Then
        cmdBack.Enabled = True
    Else
        cmdBack.Enabled = False
    End If
    
End Sub
Private Sub Form_Load()

    PicIndex = 0
    cmbPayBy.AddItem "Cheque"
    cmbPayBy.AddItem "Cash"
End Sub

Private Sub OptNew_Click(Index As Integer)
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    If Index = 0 Then
        lblAdd.Caption = "Edit Saved Export Package."
        lblInvNo.Caption = "Select Proforma No."
        Sql = "Select PInvoice,SMethod From FPInvoice Where IsNull(PackGen,0)=0"
        Editing = False
        txtBL = ""
        txtFormE = ""
        chkPack.Value = vbUnchecked
        chkPack.Visible = True
    Else
        lblAdd.Caption = "Add New Export Package."
        lblInvNo.Caption = "Select Invoice No."
        Sql = "Select PInvoice,SMethod From FPInvoice Where IsNull(PackGen,0)=1"
        Editing = True
        chkPack.Value = vbUnchecked
        chkPack.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

