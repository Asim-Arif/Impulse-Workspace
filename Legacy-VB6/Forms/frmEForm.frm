VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmEForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form E"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   10635
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7155
      Left            =   45
      TabIndex        =   3
      Top             =   30
      Width           =   10560
      Begin VB.Frame FraSave 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6750
         TabIndex        =   31
         Top             =   6345
         Visible         =   0   'False
         Width           =   3555
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   1800
            TabIndex        =   33
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
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   90
            TabIndex        =   32
            Top             =   180
            Width           =   1665
            Caption         =   "Save        "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmEForm.frx":0442
            Accelerator     =   83
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3360
         TabIndex        =   26
         Top             =   6345
         Width           =   6945
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   3480
            TabIndex        =   30
            Top             =   180
            Width           =   1665
            Caption         =   "Print             "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmEForm.frx":0554
            Accelerator     =   80
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
            TabIndex        =   29
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
         Begin MSForms.CommandButton cmdAdd 
            Height          =   360
            Left            =   75
            TabIndex        =   28
            Top             =   180
            Width           =   1665
            Caption         =   "Add New     "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmEForm.frx":0666
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdEdit 
            Height          =   360
            Left            =   1785
            TabIndex        =   27
            Top             =   180
            Width           =   1665
            Caption         =   "Edit Selected   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmEForm.frx":0778
            Accelerator     =   69
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7560
         MaxLength       =   150
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   990
         Width           =   1245
      End
      Begin VB.TextBox txtDT 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7260
         MaxLength       =   150
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1635
         Width           =   1500
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   1635
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
         Left            =   3810
         MaxLength       =   150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1635
         Width           =   1005
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   10560
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Form 'E'"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   390
            Index           =   8
            Left            =   105
            TabIndex        =   6
            Top             =   180
            Width           =   10605
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Form 'E'"
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
            Left            =   90
            TabIndex        =   7
            Top             =   180
            Width           =   10590
         End
      End
      Begin MSForms.ComboBox cmbInvNo 
         Height          =   285
         Left            =   4830
         TabIndex        =   35
         Tag             =   "AutoCatID"
         Top             =   1635
         Width           =   2445
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4313;503"
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
         Left            =   4830
         TabIndex        =   34
         Tag             =   "AutoCatID"
         Top             =   1635
         Visible         =   0   'False
         Width           =   2445
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4313;503"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
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
         Caption         =   " Date"
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
         Index           =   4
         Left            =   7560
         TabIndex        =   25
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Date"
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
         Index           =   3
         Left            =   7260
         TabIndex        =   22
         Top             =   1395
         Width           =   1500
      End
      Begin MSForms.TextBox txtAWBNo 
         Height          =   570
         Left            =   2130
         TabIndex        =   21
         Top             =   4725
         Width           =   1890
         VariousPropertyBits=   -1467987941
         BorderStyle     =   1
         Size            =   "3334;1005"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtValOnDoc 
         Height          =   570
         Left            =   285
         TabIndex        =   20
         Top             =   4725
         Width           =   1875
         VariousPropertyBits=   -1467987941
         BorderStyle     =   1
         Size            =   "3307;1005"
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
         Caption         =   "        Value On Doc.                  AWB No."
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
         Left            =   285
         TabIndex        =   19
         Top             =   4485
         Width           =   3735
      End
      Begin MSForms.TextBox txtCustom 
         Height          =   885
         Left            =   7710
         TabIndex        =   18
         Top             =   3540
         Width           =   1890
         VariousPropertyBits=   -1467987941
         BorderStyle     =   1
         Size            =   "3334;1561"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtShipPort 
         Height          =   885
         Left            =   5865
         TabIndex        =   17
         Top             =   3540
         Width           =   1875
         VariousPropertyBits=   -1467987941
         BorderStyle     =   1
         Size            =   "3307;1561"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtStreamer 
         Height          =   885
         Left            =   2145
         TabIndex        =   16
         Top             =   3540
         Width           =   3735
         VariousPropertyBits=   -1467987941
         BorderStyle     =   1
         Size            =   "6588;1561"
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
         Caption         =   $"frmEForm.frx":088A
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
         Left            =   285
         TabIndex        =   15
         Top             =   3300
         Width           =   9315
      End
      Begin MSForms.TextBox txtConsignee 
         Height          =   885
         Left            =   285
         TabIndex        =   14
         Top             =   3540
         Width           =   1875
         VariousPropertyBits=   -1467987941
         BorderStyle     =   1
         Size            =   "3307;1561"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDestination 
         Height          =   885
         Left            =   7725
         TabIndex        =   13
         Top             =   2355
         Width           =   1875
         VariousPropertyBits=   -1467987941
         BorderStyle     =   1
         Size            =   "3307;1561"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtTOS 
         Height          =   885
         Left            =   5865
         TabIndex        =   12
         Top             =   2355
         Width           =   1875
         VariousPropertyBits=   -1467987941
         BorderStyle     =   1
         Size            =   "3307;1561"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtInvValue 
         Height          =   885
         Left            =   4005
         TabIndex        =   11
         Top             =   2355
         Width           =   1875
         VariousPropertyBits=   -1467987941
         BorderStyle     =   1
         Size            =   "3307;1561"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtQty 
         Height          =   885
         Left            =   2145
         TabIndex        =   10
         Top             =   2355
         Width           =   1875
         VariousPropertyBits=   -1467987941
         BorderStyle     =   1
         Size            =   "3307;1561"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtGoods 
         Height          =   885
         Left            =   285
         TabIndex        =   9
         Top             =   2355
         Width           =   1875
         VariousPropertyBits=   -1467987941
         BorderStyle     =   1
         Size            =   "3307;1561"
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
         Caption         =   $"frmEForm.frx":0912
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
         Left            =   285
         TabIndex        =   8
         Top             =   2115
         Width           =   9315
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Customer                     Country                       Currency       Invoice No."
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
         Left            =   285
         TabIndex        =   4
         Top             =   1395
         Width           =   6990
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2055
         TabIndex        =   1
         Top             =   1635
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
   End
End
Attribute VB_Name = "frmEForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String

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
        .Open "Select Distinct PInvoice From VShippingInstructions Where CustCode+Country='" & cmbCust.ID & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbInvNo.Clear
        Do Until .EOF
            cmbInvNo.AddItem .Fields(0) & ""
            
            .MoveNEXT
        Loop
        .Close
        
        .Open "Select PInvoice From FPInvoice Where PInvoice Not In (Select PInvoice From ShippingInstructions) and CustCode+Country='" & cmbCust.ID & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
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

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub


Private Sub cmbinvno_Change()
    Call cmbinvno_click
End Sub

Private Sub cmbinvno_click()

    On Error GoTo err
    If cmbInvNo.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From VrptProformas Where PInvoice='" & cmbInvNo.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        
        txtDT = Format(![DT] & "", "dd-MMM-yyyy")
        
        'txtGoods = ![Goods] & ""
        txtQty = ![Qty] & " -Pcs"
        'txtInvValue = ![InvValue] & ""
        txtTOS = ![TradeTerms] & ""
        txtDestination = ![DisPort] & ""
        txtConsignee = ![Address] & ""
        'txtStreamer = ![Carrier] & ""
        txtShipPort = ![DisPort] & ""
        txtCustom = ![LoadPort] & ""
        'txtValOnDoc = ![ValueOnDoc] & ""
        txtAWBNo = ![AWBNo] & ""
        .Close
    End With
    
    Exit Sub
err:
    MsgBox err.Description
End Sub





Private Sub cmdAdd_Click()
    frmNewEForm.Show
End Sub

Private Sub lok(Locked As Boolean)
    
    If Locked Then PInvoice = ""
    
    FraAdd.Visible = Locked
    FraSave.Visible = Not Locked
    
    cmbInvNo.Visible = Locked
    cmbInvNoNew.Visible = Not Locked
    
    cmbInvNo.Locked = False
    
    
    For Each C In Controls
        If TypeOf C Is TextBox Then C.Locked = Locked
    Next
    
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass

    Load frmPrevRpt
    'frmPrevRpt.ShowReport "{VEForm.CustomInvoice}='" & cmbCustomInvoice & "' ", rptFormE1
    
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
       
    On Error GoTo err
      
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'HICO'")
    OrderDT = Date
    DeliveryDT = DateAdd("m", 1, Date)
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.Move 1365, 1845
    FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub

Private Function Saved() As Boolean
    
    Saved = False
    Call StartTrans(con)
    '----------------------------------------------------------------
    If strOrderNo <> "" Then
        con.Execute "Delete from EForm Where CustomInvoice='" & _
         strOrderNo & "'"
    End If
    
    con.Execute "Insert Into EForm(CustomInvoice,Goods,Qty,InvValue,TermsOfSale" & _
     ",Destination,Consignee,Carrier,PortOfShpment,CustomsPort,ValueOnDoc,AWBNo)" & _
     " Values('" & cmbCustomInvoice & "','" & txtGoods & "','" & txtQty & "','" & _
     txtInvValue & "','" & txtTOS & "','" & txtDestination & "','" & _
     txtConsignee & "','" & txtStreamer & "','" & txtShipPort & "','" & _
     txtCustom & "','" & txtValOnDoc & "','" & txtAWBNo & "')"
      
    con.CommitTrans
    
    Saved = True
    Exit Function

err:
    MsgBox err.Description
End Function

