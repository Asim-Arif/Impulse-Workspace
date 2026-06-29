VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRcvRebates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receive Rebate Payments..."
   ClientHeight    =   2355
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtChq 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1425
      MaxLength       =   150
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   1605
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   300
      Left            =   2430
      TabIndex        =   0
      Top             =   345
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
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
      Format          =   55443459
      CurrentDate     =   37928
   End
   Begin VB.TextBox txtAmt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3045
      MaxLength       =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   1170
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1605
      Width           =   4170
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   45
      TabIndex        =   9
      Top             =   960
      Width           =   1380
      _ExtentX        =   2434
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
      Format          =   55443459
      CurrentDate     =   37928
   End
   Begin MSForms.ComboBox cmbCustomInvoice 
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   345
      Width           =   2370
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4180;503"
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
      Caption         =   " Description"
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
      Index           =   1
      Left            =   45
      TabIndex        =   8
      Top             =   1335
      Width           =   4170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Rcv Date             Cheque No.               Amount"
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
      Top             =   720
      Width           =   4170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Invoice No.                               Inv. Date"
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
      Width           =   4170
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   4320
      Picture         =   "frmRcvRebates.frx":0000
      Top             =   930
      Width           =   480
   End
   Begin MSForms.CommandButton cmdClose 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4275
      TabIndex        =   5
      Top             =   1950
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Cancel          "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmRcvRebates.frx":05EF
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
      Height          =   360
      Left            =   4275
      TabIndex        =   4
      Top             =   1500
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Receive     "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmRcvRebates.frx":0701
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmRcvRebates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbCustomInvoice_Change()
    Call cmbCustomInvoice_Click
End Sub

Private Sub cmbCustomInvoice_Click()
    
    txtAmt.Tag = cmbCustomInvoice.List(cmbCustomInvoice.ListIndex, 1)
    txtAmt.Text = txtAmt.Tag
    DT1 = cmbCustomInvoice.List(cmbCustomInvoice.ListIndex, 2)
    
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
    
    If Val(txtAmt) > Val(txtAmt.Tag) Or Val(txtAmt) <= 0 Then
        MsgBox "Invalid Amount Entered.", vbInformation
        Cancel = True
        Exit Sub
    End If
    
    Call StartTrans(con)
    
    
    
    con.Execute "Insert Into RebateStatus(CustomInvoice,DT,ChqNo,Amt,Description) Values('" & _
      cmbCustomInvoice.Text & "','" & DTPicker1.value & "','" & txtChq & "'," & Val(txtAmt) & ",'" & _
      txtDescription & "')"
    
    
    
    con.CommitTrans
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Load()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select CustomInvoice,DT,Balance From VRebateStatus Where Balance>0 And EntryType=0 ORDER BY CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
        cmbCustomInvoice.Clear
        Do Until .EOF
            cmbCustomInvoice.AddItem ![CustomInvoice] & ""
            cmbCustomInvoice.List(cmbCustomInvoice.ListCount - 1, 1) = ![Balance] & ""
            cmbCustomInvoice.List(cmbCustomInvoice.ListCount - 1, 2) = Format(![DT] & "", "dd-MMM-yyyy")
            .MoveNext
        Loop
        .Close
    End With
        
    DTPicker1.value = Date
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub txtAmt_Validate(Cancel As Boolean)

    If Val(txtAmt) > Val(txtAmt.Tag) Or Val(txtAmt) <= 0 Then
        MsgBox "Invalid Amount Entered.", vbInformation
        Cancel = True
        Exit Sub
    End If
    
End Sub

Private Sub txtDescription_Validate(Cancel As Boolean)
    txtDescription = Replace(txtDescription, "|", "")
End Sub
Public Sub ShowForm(InvoiceNo As String)

    On Error GoTo err
    
    cmbCustomInvoice = InvoiceNo
    
    txtAmt.TabIndex = 0
    Me.Show 1
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
