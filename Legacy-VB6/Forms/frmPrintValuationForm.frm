VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmPrintValuationForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Valuation Form"
   ClientHeight    =   2220
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPrintValue 
      Caption         =   "Print Value"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FC1603&
      Height          =   285
      Left            =   2100
      TabIndex        =   10
      Top             =   1440
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1440
   End
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
   Begin VB.OptionButton OptPrint 
      Caption         =   "Front Side"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   9
      Top             =   1470
      Value           =   -1  'True
      Width           =   2280
   End
   Begin VB.OptionButton OptPrint 
      Caption         =   "Back Side"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   8
      Top             =   1830
      Width           =   2280
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
      Picture         =   "frmPrintValuationForm.frx":0000
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
      Picture         =   "frmPrintValuationForm.frx":05EF
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   3630
      TabIndex        =   0
      Top             =   105
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Print"
      PicturePosition =   327683
      Size            =   "3096;635"
      Accelerator     =   80
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmPrintValuationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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


Private Sub cmdPrint_Click()

    On Error GoTo err
    'Exit Sub
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    If OptPrint(0).Value Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptValuationFormA_1.rpt") 'rptValuationFormA_1
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptValuationFormA_2.rpt") 'rptValuationFormA_2
    End If
     
    If strCompany = "IAA" Or strCompany = "Kami" Then
        If chkPrintValue.Value = vbChecked Then
            rpt.FormulaFields.GetItemByName("HideValue").Text = False
        Else
            rpt.FormulaFields.GetItemByName("HideValue").Text = True
        End If
    End If
    
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptValuationFormA.CustomInvoice}='" & cmbCustomInvoice & "' ", rpt
    
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
        
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'HICO'")
 
    cmbCust.ListHeight = 2500
        
    If strCompany = "IAA" Or strCompany = "Kami" Then
        chkPrintValue.Visible = True
    Else
        chkPrintValue.Visible = False
    End If
    
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
    
    cmbCustomInvoice.Clear
    With rs
        .Open "Select Distinct CustomInvoice From VrptCustomInvoice Where CustCode='" & cmbCust.ID & "' And Country='" & cmbCountry & "' ORDER BY CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
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


