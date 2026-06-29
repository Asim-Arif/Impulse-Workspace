VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmNewEForm 
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10635
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewEForm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   10635
   WindowState     =   2  'Maximized
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
      TabIndex        =   7
      Top             =   30
      Width           =   10560
      Begin VB.TextBox txtDT 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6570
         MaxLength       =   150
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1035
         Width           =   1245
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   1035
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
         Top             =   1035
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
         TabIndex        =   10
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
            TabIndex        =   11
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
            TabIndex        =   12
            Top             =   180
            Width           =   10590
         End
      End
      Begin VB.Frame Frame6 
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
         Height          =   660
         Left            =   4665
         TabIndex        =   8
         Top             =   6405
         Width           =   5280
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   3510
            TabIndex        =   6
            Top             =   195
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
         Begin MSForms.CommandButton cmdSClear 
            Height          =   360
            Left            =   90
            TabIndex        =   4
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSClose 
            Height          =   360
            Left            =   1800
            TabIndex        =   5
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;635"
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
         Left            =   6570
         TabIndex        =   28
         Top             =   795
         Width           =   1245
      End
      Begin MSForms.TextBox txtAWBNo 
         Height          =   570
         Left            =   2130
         TabIndex        =   27
         Top             =   4125
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
         TabIndex        =   26
         Top             =   4125
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
         TabIndex        =   25
         Top             =   3885
         Width           =   3735
      End
      Begin MSForms.TextBox txtCustom 
         Height          =   885
         Left            =   7710
         TabIndex        =   24
         Top             =   2940
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
         TabIndex        =   23
         Top             =   2940
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
         TabIndex        =   22
         Top             =   2940
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
         Caption         =   $"frmNewEForm.frx":0442
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
         TabIndex        =   21
         Top             =   2700
         Width           =   9315
      End
      Begin MSForms.TextBox txtConsignee 
         Height          =   885
         Left            =   285
         TabIndex        =   20
         Top             =   2940
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
         TabIndex        =   19
         Top             =   1755
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
         TabIndex        =   18
         Top             =   1755
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
         TabIndex        =   17
         Top             =   1755
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
         TabIndex        =   16
         Top             =   1755
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
         TabIndex        =   15
         Top             =   1755
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
         Caption         =   $"frmNewEForm.frx":04CA
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
         TabIndex        =   14
         Top             =   1515
         Width           =   9315
      End
      Begin MSForms.ComboBox cmbCustomInvoice 
         Height          =   285
         Left            =   4830
         TabIndex        =   3
         Top             =   1035
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
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   270
         TabIndex        =   13
         Top             =   6720
         Width           =   1545
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "2725;609"
         Value           =   "0"
         Caption         =   "Print On Save"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
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
         TabIndex        =   9
         Top             =   795
         Width           =   6270
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2055
         TabIndex        =   1
         Top             =   1035
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
Attribute VB_Name = "frmNewEForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String
Public Sub EditOrder(TempOrdNo As String)
    
    On Error GoTo Err
    strOrderNo = TempOrdNo
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From VEForm Where CustomInvoice='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        
        
        With cmbCustomInvoice
            .Clear
            .AddItem strOrderNo
            .ListIndex = 0
            .Locked = True
        End With
    
        
        txtGoods = ![Goods] & ""
        txtQty = ![Qty] & ""
        txtInvValue = ![InvValue] & ""
        txtTOS = ![TermsOfSale] & ""
        txtDestination = ![Destination] & ""
        txtConsignee = ![Consignee] & ""
        txtStreamer = ![Carrier] & ""
        txtShipPort = ![PortOfShpment] & ""
        txtCustom = ![CustomsPort] & ""
        txtValOnDoc = ![ValueOnDoc] & ""
        txtAWBNo = ![AwbNo] & ""
        .Close
        
        
    End With
    
    Set rs = Nothing
    Me.Show
    Exit Sub
Err:
    MsgBox Err.Description
    
End Sub


Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo Err
    If cmbCountry.MatchFound = False Then Exit Sub
    cmbCountry.Tag = cmbCust.Text & cmbCountry
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_CustomInvoiceForEForm"
        
        .Parameters("@Customer").value = cmbCountry.Tag
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = cmd.Execute
    End With
    
    cmbCustomInvoice.Clear
    With rs
        
        Do Until .EOF
            cmbCustomInvoice.AddItem .Fields(0) & ""
            
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
Err:
    MsgBox Err.Description
    
End Sub

Private Sub cmbCust_matched()
    
    On Error GoTo Err
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
Err:
    MsgBox Err.Description
    
End Sub

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub

Private Sub cmbCustomInvoice_Change()
    Call cmbCustomInvoice_Click
End Sub

Private Sub cmbCustomInvoice_Click()

    On Error GoTo Err
    If cmbCustomInvoice.MatchFound = False Then Exit Sub
    

    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From VCustomPList Where CustomInvoice='" & cmbCustomInvoice.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        
        txtDT = Format(![DT], "dd-MMM-yyyy")
        
        txtQty = ![TotalCartons] & " -Cartons" & vbNewLine & ![TotalQty] & " " & ![Unit]
        txtInvValue = ![Curr] & " : " & ![totamt] & vbNewLine & ![TradeTerms]
        txtTOS = ![PayTerms] & ""
        txtDestination = ![DisPort] & vbNewLine & ![DisCity] & vbNewLine & ![DisCountry]
        txtConsignee = ![Address2] & ""
        txtStreamer = UCase(![ShippingMode] & "")
        txtAWBNo = "AWB # : " & ![AwbNo] & vbNewLine & Format(![AWBNoDT] & "", "dd-MMM-yyyy")
        txtValOnDoc = ![Curr] & " : " & ![totamt]
        txtShipPort = "Sialkot."
        txtCustom = "SAMBRIAL" & vbNewLine & "Siakot."
        
        .Close
    End With
    
    Exit Sub
Err:
    MsgBox Err.Description
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdSClear_Click()
    If Saved Then
        Unload Me
        Me.Show
    End If
End Sub

Private Sub cmdSClose_Click()
    If Saved Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
       
    On Error GoTo Err
      
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'HICO'")
    OrderDT = Date
    DeliveryDT = DateAdd("m", 1, Date)
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    txtGoods = "PLIERS/" & vbNewLine & "H.TOOLS"
    
    Exit Sub
Err:
    MsgBox Err.Description
   
End Sub

Private Sub Form_Resize()
On Error Resume Next
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
    
    con.Execute "Insert Into EForm(CustomInvoice,DT,Goods,Qty,InvValue,TermsOfSale" & _
     ",Destination,Consignee,Carrier,PortOfShpment,CustomsPort,ValueOnDoc,AWBNo)" & _
     " Values('" & cmbCustomInvoice & "','" & txtDT & "','" & txtGoods & "','" & txtQty & "','" & _
     txtInvValue & "','" & txtTOS & "','" & txtDestination & "','" & _
     txtConsignee & "','" & txtStreamer & "','" & txtShipPort & "','" & _
     txtCustom & "','" & txtValOnDoc & "','" & txtAWBNo & "')"
      
    con.CommitTrans
    
    Saved = True
    Exit Function

Err:
    MsgBox Err.Description
End Function

