VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmShippingInstructions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SHIPPING INSTRUCTIONS FORM"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShippingInstructions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   5985
      Left            =   75
      TabIndex        =   3
      Top             =   645
      Width           =   7200
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   480
         Width           =   2280
         _ExtentX        =   4022
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
      Begin VB.TextBox txtEncloseres 
         Appearance      =   0  'Flat
         Height          =   1110
         Left            =   4500
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   3015
         Width           =   2520
      End
      Begin VB.TextBox txtNonKind 
         Appearance      =   0  'Flat
         Height          =   1110
         Left            =   1935
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   3015
         Width           =   2520
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1935
         TabIndex        =   21
         Top             =   945
         Width           =   5070
      End
      Begin VB.TextBox txtGW 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1935
         TabIndex        =   20
         Top             =   1290
         Width           =   1950
      End
      Begin VB.TextBox txtMarks 
         Appearance      =   0  'Flat
         Height          =   1110
         Left            =   1935
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1635
         Width           =   5085
      End
      Begin VB.TextBox txtMeasurement 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5085
         TabIndex        =   18
         Top             =   1290
         Width           =   1950
      End
      Begin VB.TextBox txtSpecial 
         Appearance      =   0  'Flat
         Height          =   1110
         Left            =   1935
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   4170
         Width           =   5100
      End
      Begin VB.Frame FraSave 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3540
         TabIndex        =   13
         Top             =   5280
         Visible         =   0   'False
         Width           =   3555
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   90
            TabIndex        =   15
            Top             =   180
            Width           =   1665
            Caption         =   "Save        "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmShippingInstructions.frx":014A
            Accelerator     =   83
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
            Left            =   1800
            TabIndex        =   14
            Top             =   180
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FraAdd 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   150
         TabIndex        =   4
         Top             =   5280
         Width           =   6945
         Begin MSForms.CommandButton cmdEdit 
            Height          =   360
            Left            =   1785
            TabIndex        =   12
            Top             =   180
            Width           =   1665
            Caption         =   "Edit Selected   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmShippingInstructions.frx":025C
            Accelerator     =   69
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAdd 
            Height          =   360
            Left            =   75
            TabIndex        =   11
            Top             =   180
            Width           =   1665
            Caption         =   "Add New     "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmShippingInstructions.frx":036E
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   5175
            TabIndex        =   6
            Top             =   180
            Width           =   1665
            Caption         =   "Close            "
            PicturePosition =   327683
            Size            =   "2937;635"
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
            Left            =   3480
            TabIndex        =   5
            Top             =   180
            Width           =   1665
            Caption         =   "Print             "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmShippingInstructions.frx":0480
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Encloseres"
         Height          =   195
         Index           =   1
         Left            =   4515
         TabIndex        =   30
         Top             =   2790
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No. && Kind Of Packages"
         Height          =   195
         Index           =   0
         Left            =   1965
         TabIndex        =   28
         Top             =   2790
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Description Of Goods :"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Gross Weight (Kg) :"
         Height          =   195
         Index           =   6
         Left            =   450
         TabIndex        =   25
         Top             =   1305
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Marks && Number :"
         Height          =   195
         Index           =   7
         Left            =   585
         TabIndex        =   24
         Top             =   1635
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Measurement :"
         Height          =   195
         Index           =   8
         Left            =   3990
         TabIndex        =   23
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Special Instruction :"
         Height          =   195
         Index           =   9
         Left            =   435
         TabIndex        =   22
         Top             =   4140
         Width           =   1425
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2400
         TabIndex        =   10
         Top             =   480
         Width           =   2265
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3995;503"
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
         Caption         =   " Customer                                 Country                                   Invoice No"
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
         Index           =   15
         Left            =   90
         TabIndex        =   9
         Top             =   240
         Width           =   7035
      End
      Begin MSForms.ComboBox cmbInvNo 
         Height          =   285
         Left            =   4695
         TabIndex        =   8
         Tag             =   "AutoCatID"
         Top             =   480
         Width           =   2430
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4286;503"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbInvNoNew 
         Height          =   285
         Left            =   4695
         TabIndex        =   16
         Tag             =   "AutoCatID"
         Top             =   480
         Visible         =   0   'False
         Width           =   2430
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4286;503"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   195
      TabIndex        =   0
      Top             =   0
      Width           =   6960
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping Instructions Form"
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
         Index           =   8
         Left            =   30
         TabIndex        =   1
         Top             =   150
         Width           =   6765
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Shipping Instructions Form"
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
         Height          =   465
         Index           =   9
         Left            =   45
         TabIndex        =   2
         Top             =   165
         Width           =   6750
      End
   End
End
Attribute VB_Name = "frmShippingInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PInvoice As String
Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    
    cmbCust.Tag = cmbCust.Text & cmbCountry
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    
    If strProforma <> "" Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select CustomInvoice,ShippingL From VShippingInstructions1 Where CustCode+Country='" & cmbCust.ID & cmbCountry.Text & "' ORDER BY CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
        cmbInvNo.Clear
        Do Until .EOF
            cmbInvNo.AddItem .Fields(0) & ""
            cmbInvNo.List(cmbInvNo.ListCount - 1, 1) = .Fields(1) & ""
            .MoveNEXT
        Loop
        .Close
        
        .Open "Select CustomInvoice From CustomInvoice Where CustomInvoice Not In (Select CustomInvoice From ShippingInstructions) and CustCode+Country='" & cmbCust.ID & cmbCountry.Text & "' ORDER BY CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
        cmbInvNoNew.Clear
        Do Until .EOF
            cmbInvNoNew.AddItem .Fields(0) & ""
            
            .MoveNEXT
        Loop
        .Close
        
    End With
    Set rs = Nothing
    
    
    
    Exit Sub
err:
    MsgBox err.Description
    
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

Private Sub cmbinvno_Change()
    Call cmbinvno_click
End Sub

Private Sub cmbInvNoNew_Click()

    On Error GoTo err
    If cmbInvNoNew.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    Dim TotalCartons As Long
    
    With rs
        .Open "Select * From VTotalCartons Where CustomInvoice='" & cmbInvNoNew & "'", con, adOpenForwardOnly, adLockReadOnly
        
        
        txtEncloseres = "1.   INVOICES 5 COPIES." & vbNewLine & _
                        "2.   PACKING LIST 2 COPIES." & vbNewLine & _
                        "3.   VALUATION FORM (A)." & vbNewLine & _
                        "4.   FORM E ONE SET." & vbNewLine & _
                        "5.   UNDERTAKING" & vbNewLine & _
                        "6.   COMMERCIAL EXPORTER DECLARATION"
        
        txtMarks = Left(cmbInvNoNew, InStr(5, cmbInvNoNew, "-") + 4)
        txtDescription = "SURGICAL INSTRUMENTS."
        If .RecordCount > 0 Then
        
            txtMarks = txtMarks & "/1 TO " & Val(![TotalCartons] & "")
                        
            txtNonKind = "(" & Val(![TotalCartons] & ")") & " CARTONS" & vbNewLine & _
                            "CONTAINING DISPOSABLE" & vbNewLine & _
                            "SURGICAL INSTRUMENTS"
                            
            TotalCartons = Val(![TotalCartons] & "")
            
        Else
            txtMarks = txtMarks & "/1 TO "
                        
            txtNonKind = " CARTONS" & vbNewLine & _
                         "CONTAINING DISPOSABLE" & vbNewLine & _
                         "SURGICAL INSTRUMENTS"
        End If
        .Close
        
'        .Open "Select Sum(Qty*(FinishedWeight/1000))/2 From VrptSteelCalculation Where PInvoice='" & cmbInvNoNew & "'", con, adOpenForwardOnly, adLockReadOnly
'        If .RecordCount > 0 Then
'            txtGW = Val(.Fields(0) & " Kg")
'        End If
        
        
        .Open "Select * From BoxSize", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            txtMeasurement = Val(.Fields(0) & "") * Val(.Fields(1) & "") * Val(.Fields(2) & "") * TotalCartons
        End If
        .Close
        .Open "Select InvoiceOf From CustomInvoice Where CustomInvoice='" & cmbInvNoNew & "'", con, adOpenForwardOnly, adLockReadOnly
        txtDescription = !InvoiceOf & ""
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbinvno_click()

    On Error GoTo err
    If cmbInvNo.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "Select * From VShippingInstructions Where CustomInvoice='" & cmbInvNo & "'", con, adOpenForwardOnly, adLockReadOnly
        txtDescription = ![GoodsDesc] & ""
        txtGW = ![GWeight] & ""
        txtMeasurement = ![Measurement] & ""
        txtMarks = ![MarksnNumbers] & ""
        txtNonKind = ![NonKind] & ""
        txtEncloseres = ![Encloseres] & ""
        
        txtSpecial = ![SpecialInstructions] & ""
        
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmdAdd_Click()
        
        
    Call lok(False)
    PInvoice = ""
    
End Sub
Private Sub lok(Locked As Boolean)
    
    If Locked Then PInvoice = ""
    
    FraAdd.Visible = Locked
    FraSave.Visible = Not Locked
    
    cmbInvNo.Visible = Locked
    cmbInvNoNew.Visible = Not Locked
    
    cmbInvNo.Locked = False
    
    
    For Each c In Controls
        If TypeOf c Is TextBox Then c.Locked = Locked
    Next
    
End Sub


Private Sub cmdCancel_Click()
    Call cmbinvno_click
    Call lok(True)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    
    
    If cmbInvNo.MatchFound = False Then Exit Sub
'    If cmbInvNo.List(cmbInvNo.ListIndex, 1) Then
'        MsgBox "This Record Has Been Locked.", vbInformation
'        Exit Sub
'    End If
        
    PInvoice = cmbInvNo
    
    FraSave.Visible = True
    FraAdd.Visible = False
    
    cmbInvNo.Locked = True
    
    Call lok(False)
    
End Sub

Private Sub cmdPrint_Click()
    
    Dim strRet As String
    strRet = ShowPopUpMenu1("Shipping Instructions (With Value)", "Shipping Instructions (Without Value)")
    
    If strRet = "Shipping Instructions (With Value)" Then
        Call PrintShippingInstructions(False)
    ElseIf strRet = "Shipping Instructions (Without Value)" Then
        Call PrintShippingInstructions(True)
    End If
    
End Sub

Private Sub PrintShippingInstructions(Optional bHideValue As Boolean = False)
    
    On Error GoTo err
    If cmbInvNo.MatchFound = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report 'New rptShipping
    Set rpt = rptApp.OpenReport(RptPath & "\rptShipping.rpt") '
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@HideValue}" Then
            FormulaField.Text = bHideValue
            Exit For
        End If
    Next
            
    Dim f As New frmPrevRpt
    
    Load f
    f.ShowReport "{VShippingInstructions.CustomInvoice}='" & cmbInvNo.Text & "'", rpt
    
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
    Call StartTrans(con)
    
    If PInvoice <> "" Then
        con.Execute "Delete From ShippingInstructions Where CustomInvoice='" & PInvoice & "'"
    End If
    
    
    con.Execute "Insert Into ShippingInstructions(CustomInvoice,GoodsDesc" & _
     ",GWeight,Measurement,MarksnNumbers,NonKind,Encloseres," & _
     "SpecialInstructions) Values('" & _
     IIf(PInvoice = "", cmbInvNoNew, cmbInvNo) & "','" & txtDescription & _
     "','" & txtGW & "','" & txtMeasurement & "','" & txtMarks & "','" & _
     txtNonKind & "','" & txtEncloseres & "','" & txtSpecial & "')"
    
    con.CommitTrans
    
    Call lok(True)
    Call cmbCountry_Click
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub Form_Load()

    On Error GoTo err
    
    cmbCust.ListHeight = 2100
    
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where Active=1")
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
