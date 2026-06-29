VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmNewCustomPListA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Custom Packing List"
   ClientHeight    =   1620
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Top             =   375
      Width           =   1740
      _ExtentX        =   3069
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
   Begin VB.TextBox txtCurrency 
      Appearance      =   0  'Flat
      DataField       =   "RptTime"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   45
      MaxLength       =   150
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Currency                      Invoice No"
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
      Top             =   810
      Width           =   3495
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   285
      Left            =   1815
      TabIndex        =   6
      Top             =   375
      Width           =   1725
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3043;503"
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
      Caption         =   " Customer                     Country"
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
      TabIndex        =   5
      Top             =   105
      Width           =   3495
   End
   Begin MSForms.ComboBox cmbCustomInvoice 
      Height          =   285
      Left            =   1815
      TabIndex        =   4
      Top             =   1080
      Width           =   1725
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3043;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   3750
      Picture         =   "frmNewCustomPListA.frx":0000
      Top             =   1005
      Width           =   480
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   3630
      TabIndex        =   1
      Top             =   555
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Cancel       "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmNewCustomPListA.frx":05EF
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
      Left            =   3630
      TabIndex        =   0
      Top             =   105
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Generate"
      PicturePosition =   327683
      Size            =   "3096;635"
      Accelerator     =   71
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmNewCustomPListA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReGenerate As Boolean
Public Sub ShowForm(ReGen As Boolean)
    ReGenerate = ReGen
    If ReGenerate Then
        cmdGenerate.Caption = "ReGenerate"
    Else
        cmdGenerate.Caption = "Generate"
    End If
    Me.Show 1
End Sub
Private Sub cmbCust_matched()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Country,Curr From ForeignCustomers Where CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    
    Dim RecAffect As Long
    If cmbCustomInvoice.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select PackingMode From VrptCustomInvoiceDetail Where CustomInvoice='" & cmbCustomInvoice.Text & "' and isnull(PackingMode,0)=0", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            MsgBox "One Or More Items In This Invoice Has No PackingMode Defined.", vbInformation
            Exit Sub
        End If
        .Close
    End With
    Set rs = Nothing
    
    Dim cmd As New ADODB.Command
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_CreatePList"
        .Parameters("@CustomInvoice").Value = cmbCustomInvoice.Text
        .Execute RecAffect
    End With
    
    'If RecAffect > 0 Then
        MsgBox "Packing List Generated Succesfully.", vbInformation
        Unload Me
    'End If
    
End Sub

Private Sub Form_Load()
        
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'HICO'")
 
    cmbCust.ListHeight = 2500
    
End Sub
Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    cmbCountry.Tag = cmbCust.Text & cmbCountry
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_CustomInvoiceForPList"
        
        .Parameters("@Customer").Value = cmbCountry.Tag
        
        .Parameters("@ReGen").Value = ReGenerate
        
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
    End With
    
    cmbCustomInvoice.Clear
    With rs
        Do Until .EOF
            cmbCustomInvoice.AddItem .Fields(0) & ""
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


