VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmAddBankNo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assign FDBC No. To Invoices"
   ClientHeight    =   2715
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFormE 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4275
      MaxLength       =   150
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   375
      Width           =   1740
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   300
      Left            =   2280
      TabIndex        =   1
      Top             =   1095
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
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
      Format          =   27918339
      CurrentDate     =   37928
   End
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   375
      Width           =   2220
      _ExtentX        =   3916
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
   Begin VB.TextBox txtFDBC 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   45
      MaxLength       =   150
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1095
      Width           =   2220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   405
      X2              =   5969
      Y1              =   2160
      Y2              =   2160
   End
   Begin MSForms.ComboBox cmbCustomInvoice 
      Height          =   285
      Left            =   2295
      TabIndex        =   2
      Top             =   375
      Width           =   1965
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3466;503"
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
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FDBC No.                                 Date"
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
      Left            =   45
      TabIndex        =   7
      Top             =   825
      Width           =   4170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Customer                               Invoice                              Form E No."
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
      Left            =   45
      TabIndex        =   6
      Top             =   105
      Width           =   5970
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   4320
      Picture         =   "frmAddBankNo.frx":0000
      Top             =   930
      Width           =   480
   End
   Begin MSForms.CommandButton cmdClose 
      CausesValidation=   0   'False
      Height          =   390
      Left            =   4245
      TabIndex        =   5
      Top             =   2295
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Cancel          "
      PicturePosition =   327683
      Size            =   "3096;688"
      Picture         =   "frmAddBankNo.frx":05EF
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdGenerate 
      Height          =   390
      Left            =   2445
      TabIndex        =   4
      Top             =   2295
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Save            "
      PicturePosition =   327683
      Size            =   "3096;688"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   420
      X2              =   5969
      Y1              =   2175
      Y2              =   2175
   End
End
Attribute VB_Name = "frmAddBankNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbCust_matched()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select CustomInvoice,FormE From CustomInvoice Where CustCode='" & cmbCust.ID & "' And CustomInvoice Not In(Select CustomInvoice From BankNoForEE) Order By CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
        cmbCustomInvoice.Clear
        Do Until .EOF
            cmbCustomInvoice.AddItem .Fields(0) & ""
            cmbCustomInvoice.List(cmbCustomInvoice.ListCount - 1, 1) = .Fields(1) & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbCustomInvoice_Change()
    Call cmbCustomInvoice_Click
End Sub

Private Sub cmbCustomInvoice_Click()
    txtFormE.Text = cmbCustomInvoice.List(cmbCustomInvoice.ListIndex, 1)
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    
    On Error GoTo err
    
    Dim RecAffect As Long
    
    If cmbCustomInvoice.MatchFound = False Then
        MsgBox "Select Valid Invoice.", vbInformation
        Exit Sub
    End If
    
    If txtFDBC = "" Then
        MsgBox "Must Enter FDBC No.", vbInformation
        txtFDBC.SetFocus
        Exit Sub
    End If
    
    Call StartTrans(con)
    
    con.Execute "Insert Into BankNoForEE(CustomInvoice,BankNo,ShipDT) Values('" & _
     cmbCustomInvoice.Text & "','" & txtFDBC & "','" & DT1 & "')"
    
     
    con.CommitTrans
    
    Unload Me
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Load()
        
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'HICO'")
    
    'Call AddToCombo(cmbBankName, "Distinct Bank", "BankList", "Order By Bank")
    
    cmbCust.ListHeight = 2500
    DT1 = ServerDate
    
End Sub

Public Sub ShowForm(CustCode As String, InvoiceNo As String)

    On Error GoTo err
    
    cmbCust.ID = CustCode
    cmbCustomInvoice = InvoiceNo
    txtFDBC.TabIndex = 0
    Me.Show 1
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

