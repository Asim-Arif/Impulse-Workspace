VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmPrintFormE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print E Form."
   ClientHeight    =   4140
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   45
      TabIndex        =   0
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
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      Height          =   1110
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2415
      Width           =   2490
   End
   Begin VB.TextBox txtCustom 
      Appearance      =   0  'Flat
      Height          =   1110
      Left            =   2565
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2415
      Width           =   2490
   End
   Begin VB.TextBox txtShipMode 
      Appearance      =   0  'Flat
      DataField       =   "RptTime"
      ForeColor       =   &H00800000&
      Height          =   1110
      Left            =   5070
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   990
      Width           =   2490
   End
   Begin VB.TextBox txtNIC 
      Appearance      =   0  'Flat
      DataField       =   "RptTime"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5070
      MaxLength       =   150
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2415
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.TextBox txtTerms 
      Appearance      =   0  'Flat
      Height          =   1110
      Left            =   2550
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   990
      Width           =   2490
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      Height          =   1110
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   990
      Width           =   2490
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
      Left            =   5115
      TabIndex        =   15
      Top             =   2865
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1440
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
      Left            =   5115
      TabIndex        =   14
      Top             =   3225
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox txtCurrency 
      Appearance      =   0  'Flat
      DataField       =   "RptTime"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3585
      MaxLength       =   150
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   375
      Width           =   1740
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Quantity                                       Terms of Sale                              Shipping Method"
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
      Left            =   30
      TabIndex        =   17
      Top             =   720
      Width           =   7530
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Port Of Shipment                         Land Of Custom                          NIC #"
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
      Left            =   30
      TabIndex        =   16
      Top             =   2130
      Width           =   5040
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
      Left            =   3585
      TabIndex        =   13
      Top             =   105
      Width           =   3990
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   285
      Left            =   1815
      TabIndex        =   1
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
      TabIndex        =   12
      Top             =   105
      Width           =   3495
   End
   Begin MSForms.ComboBox cmbCustomInvoice 
      Height          =   285
      Left            =   5355
      TabIndex        =   2
      Top             =   375
      Width           =   2220
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3916;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   5805
      TabIndex        =   10
      Top             =   3705
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Cancel       "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmPrintFormE.frx":0000
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
      Left            =   3990
      TabIndex        =   9
      Top             =   3705
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
Attribute VB_Name = "frmPrintFormE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ShowForm(CustCode As String, Country As String, InvNo As String)

    cmbCust.ID = CustCode
    cmbCountry = Country
    cmbCustomInvoice = InvNo
'    Dim rs As New ADODB.Recordset
'    With rs
'        .Open "Select CustomInvoice From GSP Where CustomInvoice='" & InvNo & "'", con, adOpenForwardOnly, adLockReadOnly
'        If .RecordCount = 0 Then
'            cmdAdd_Click
'            cmbInvNoNew = InvNo
'        Else
'            cmbInvNo = InvNo
'        End If
'        .Close
'    End With

    Me.Show , frmFlash
    
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
            .MoveNext
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbCustomInvoice_Change()
    Call cmbCustomInvoice_Click
End Sub

Private Sub cmbCustomInvoice_Click()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From EForm Where CustomInvoice='" & cmbCustomInvoice & "'", con, adOpenForwardOnly, adLockReadOnly
        
        If .RecordCount = 0 Then
            .Close
            .Open "Select TotalCartons,TotalQty,NICNo,PaymentTerms,ShippingMode From VCustomInvoice1 Where CustomInvoice='" & cmbCustomInvoice & "'", con, adOpenForwardOnly, adLockReadOnly
            
'            txtNIC = ![NICNo] & "" '"34603-3094131-0"
            txtQty = ![TotalCartons] & "- Cartons " '& ![TotalQty] & " Pcs"
            txtTerms = ![PaymentTerms] & ""
            txtShipMode = ![ShippingMode] & ""
            
            txtPort = "Sialkot"
            txtCustom = "Sambrial" & " Sialkot"
            
        Else
            txtNIC = ![NICNo] & ""
            txtQty = ![Qty] & ""
            txtTerms = ![TermsOfSale] & ""
            txtShipMode = ![ShippingMode] & ""
            txtPort = ![ShipPort] & ""
            txtCustom = ![LandCustom] & ""
            
        End If
        .Close
    End With
    
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
    
    Screen.MousePointer = vbHourglass
    If cmbCustomInvoice.MatchFound = False Then Exit Sub
    
    If Val(con.Execute("Select Count(CustomInvoice) From EForm Where CustomInvoice='" & cmbCustomInvoice & "'").Fields(0).value & "") = 0 Then
        con.Execute "Insert Into EForm (CustomInvoice,NICNo,Qty,TermsOfSale,ShippingMode,ShipPort,LandCustom) " & _
         "Values('" & cmbCustomInvoice & "','" & txtNIC & "','" & txtQty & _
         "','" & txtTerms & "','" & txtShipMode & "','" & txtPort & "','" & txtCustom & "')"
    Else
        con.Execute "Update EForm Set NICNo='" & txtNIC & "',Qty='" & _
         txtQty & "',TermsOfSale='" & txtTerms & "',ShippingMode='" & txtShipMode & _
         "',ShipPort='" & txtPort & "',LandCustom='" & txtCustom & "' Where CustomInvoice='" & _
         cmbCustomInvoice & "'"
    End If
      
    Dim rpt As craxddrt.Report
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select OurBankID From CustomInvoice Where CustomInvoice='" & cmbCustomInvoice & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
'            If Val(!OurBankID & "") = 7 Then
'                Set rpt = rptApp.OpenReport(rptExportPath & "\rptFormEFront.rpt")
'            ElseIf Val(!OurBankID & "") = 8 Or Val(!OurBankID & "") = 9 Then
'                Set rpt = rptApp.OpenReport(rptExportPath & "\rptFormEFrontUBL.rpt")
'            Else
'                MsgBox "Not Settings saved for Selected Bank.", vbInformation
'                Exit Sub
'            End If
            Set rpt = rptApp.OpenReport(RptPath & "\rptFormEFront.rpt")
        Else
            MsgBox "Not Settings saved for Selected Bank.", vbInformation
            Exit Sub
        End If
        .Close
    End With
    Set rs = Nothing
    
    'Prime Commercial Bank
    'United Bank Limited
    
'    If OptPrint(0).value Then
'        Set rpt = rptApp.OpenReport(rptExportPath & "\rptFormEFront.rpt")
'    Else
'        Set rpt = rptApp.OpenReport(rptExportPath & "\rptFormEFront.rpt")
'    End If
    
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptFormE.CustomInvoice}='" & cmbCustomInvoice & "' ", rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Load()
        
    Call cmbCust.AddVals(con, "Distinct CustCode", "ForeignCustomers", "CustCode", " Where Active=1")
 
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
    
    cmbCustomInvoice.Clear
    With rs
        .Open "Select CustomInvoice From CustomInvoice Where CustCode='" & cmbCust.ID & "' And Country='" & cmbCountry & "' Order By CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            cmbCustomInvoice.AddItem .Fields(0) & ""
            
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


