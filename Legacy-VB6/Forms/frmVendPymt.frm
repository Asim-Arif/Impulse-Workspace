VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendPymt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment To Vender"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
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
   ScaleHeight     =   3165
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   6255
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment To Vender"
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
         Index           =   1
         Left            =   1485
         TabIndex        =   13
         Top             =   120
         Width           =   3075
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment To Vender"
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
         Index           =   3
         Left            =   1500
         TabIndex        =   14
         Top             =   135
         Width           =   3075
      End
   End
   Begin ComboList.Usercontrol1 cmbMaker 
      Height          =   285
      Left            =   1575
      TabIndex        =   1
      Top             =   1215
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1215
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   57999363
      CurrentDate     =   37384
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00A9BDC2&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pymt ID :"
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
      Index           =   2
      Left            =   3975
      TabIndex        =   17
      Top             =   660
      Width           =   885
   End
   Begin MSForms.TextBox txtPymtID 
      Height          =   285
      Left            =   4890
      TabIndex        =   16
      Top             =   645
      Width           =   1485
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2619;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.CheckBox chkByChq 
      Height          =   300
      Left            =   135
      TabIndex        =   15
      Top             =   645
      Width           =   1875
      BackColor       =   -2147483633
      ForeColor       =   4194304
      DisplayStyle    =   4
      Size            =   "3307;529"
      Value           =   "0"
      Caption         =   "Pay By Cheque"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CheckBox chkPrint 
      Height          =   255
      Left            =   90
      TabIndex        =   6
      Top             =   2745
      Width           =   1830
      BackColor       =   -2147483633
      ForeColor       =   4194304
      DisplayStyle    =   4
      Size            =   "3228;450"
      Value           =   "0"
      Caption         =   "Print When Saved"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   3420
      TabIndex        =   8
      Top             =   2700
      Width           =   1455
      Caption         =   "Save & Close "
      PicturePosition =   327683
      Size            =   "2566;635"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label lblBalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   4875
      TabIndex        =   2
      Top             =   1215
      Width           =   1500
   End
   Begin MSForms.TextBox txtAmt 
      Height          =   285
      Left            =   105
      TabIndex        =   5
      Top             =   1800
      Width           =   1485
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2619;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.ComboBox cmbBanks 
      Height          =   285
      Left            =   1620
      TabIndex        =   3
      Top             =   1800
      Width           =   3285
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5794;503"
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
      BackColor       =   &H00A9BDC2&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Amount                  Bank  Name                                                   Chq #                     "
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
      Left            =   105
      TabIndex        =   11
      Top             =   1530
      Width           =   6255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00A9BDC2&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Date                       Vender                                                           Balance"
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
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   945
      Width           =   6255
   End
   Begin MSForms.TextBox txtChqNo 
      Height          =   285
      Left            =   4920
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2566;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   360
      Left            =   1935
      TabIndex        =   7
      Top             =   2700
      Width           =   1455
      Caption         =   "Save & New  "
      PicturePosition =   327683
      Size            =   "2566;635"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   360
      Left            =   4920
      TabIndex        =   9
      Top             =   2700
      Width           =   1455
      Caption         =   "Cancel        "
      PicturePosition =   327683
      Size            =   "2566;635"
      Picture         =   "frmVendPymt.frx":0000
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   150
      X2              =   6345
      Y1              =   2610
      Y2              =   2610
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   6345
      Y1              =   2625
      Y2              =   2625
   End
End
Attribute VB_Name = "frmVendPymt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Function Saved() As Boolean

    On Error GoTo err
    Dim Sql As String
    
    If Not cmbMaker.MatchFound Then
        MsgBox "Select Valid Maker Name", vbInformation
        Exit Function
    End If
    
    If Val(txtAmt) <= 0 Then
        MsgBox "Invalid Amount", vbInformation
        txtAmt.SetFocus
        Exit Function
    End If
    
    If chkByChq.value Then
        If cmbBanks.MatchFound = False Then
            MsgBox "Select Valid Bank", vbInformation
            cmbBanks.SetFocus
            Exit Function
        ElseIf txtChqNo = "" Then
            MsgBox "Plz Enter Valid Cheque #.", vbInformation
            txtChqNo.SetFocus
            Exit Function
        End If
    End If
    
    Dim Bal As Double
    Bal = Val(con.Execute("Select Balance From VenderLedger Where EntryID=(Select Max(EntryID) From VenderLedger Where VendID=" & cmbMaker.ID & ")").Fields(0).value & "")
    If chkByChq.value Then
        Sql = "Insert Into VenderLedger(VendID,DT,BankName,ChqNo,Amt,IsCheque,PymtID,Balance) Values('" & cmbMaker.ID & "','" & _
         Format(DTPicker1, "dd-MMM-yyyy") & "','" & cmbBanks & "','" & txtChqNo & "'," & Val(txtAmt) & ",1,'" & txtPymtID & "'," & Bal + Val(txtAmt) & ")"
    Else
        Sql = "Insert Into VenderLedger(VendID,DT,Amt,PymtID,Balance) Values('" & cmbMaker.ID & "','" & _
         Format(DTPicker1, "dd-MMM-yyyy") & "'," & Val(txtAmt) & ",'" & txtPymtID & "'," & Bal + Val(txtAmt) & ")"
    End If
    
    con.Execute Sql
    'Dim iEntryID As Long
    'iEntryID = Val(con.Execute("Select Max(EntryID) From VenderLedger").Fields(0).value & "")
    
    Saved = True
    
    'If chkPrint Then
    '    With MainForm.cr1
    '        .Connect = con
    '        .Destination = crptToWindow
    '        .ReportFileName = RptPath & "\MakerChq.rpt"
    '        .SelectionFormula = "{VMakerLedger.EntryID}=" & iEntryID
    '        .Action = 1
    '        .PageZoomNext
    '     End With
    ' End If
    
    Sql = ""
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub chkByChq_Click()

    Dim PymtID As String
    
    If chkByChq.value Then
        cmbBanks.Enabled = True
        txtChqNo.Enabled = True
        
        PymtID = Val(con.Execute("Select Max(right(PymtID,3)) From VenderLedger Where DT='" & Format(DTPicker1, "dd-MMM-yyyy") & "' and UPPER(Left(PymtID,3))='BPV'").Fields(0).value & "")
    
        PymtID = "BPV-" & _
         Format(DTPicker1, "ddMMyy") & "-" & Format(Val(PymtID) + 1, "000")
         
    Else
        cmbBanks.Enabled = False
        txtChqNo.Enabled = False
        PymtID = Val(con.Execute("Select Max(right(PymtID,3)) From VenderLedger Where DT='" & Format(DTPicker1, "dd-MMM-yyyy") & "' and UPPER(Left(PymtID,3))='CPV'").Fields(0).value & "")
    
        PymtID = "CPV-" & _
         Format(DTPicker1, "ddMMyy") & "-" & Format(Val(PymtID) + 1, "000")
         
    End If
    
    txtPymtID = PymtID
    
End Sub

Private Sub cmbMaker_matched()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select CurrentBal From VVendCurrentBal Where VendID=" & cmbMaker.ID, con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then lblBalance = Format(.Fields(0), "0.00 DR ;0.00 CR ;0.00 ")
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

Private Sub cmdNext_Click()
    If Saved Then
        Unload Me
        Me.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    
    DTPicker1 = ServerDate
    chkByChq.value = True
    
    cmbMaker.ListHeight = 1200
    Call cmbMaker.AddVals(con, "VenderName", "VVenders", "VendID")
    
    Call AddToCombo(cmbBanks, "Bank", "BankList", , , "BankID")
         
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub
