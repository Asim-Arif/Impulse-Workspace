VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmInspection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inspection Certificate"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInspection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   7155
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   90
      TabIndex        =   24
      Top             =   945
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
   Begin VB.Frame Frame2 
      Height          =   3645
      Left            =   75
      TabIndex        =   3
      Top             =   1335
      Width           =   7035
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   315
         Left            =   1785
         TabIndex        =   18
         Top             =   1530
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   50266115
         CurrentDate     =   37250
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
         Height          =   1125
         Left            =   5085
         TabIndex        =   20
         Top             =   2430
         Width           =   1860
         Begin MSForms.CommandButton cmdPrint 
            Height          =   390
            Left            =   75
            TabIndex        =   22
            Top             =   180
            Width           =   1665
            Caption         =   "Print "
            PicturePosition =   327683
            Size            =   "2937;688"
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
            Height          =   390
            Left            =   60
            TabIndex        =   21
            Top             =   630
            Width           =   1665
            Caption         =   "Close            "
            PicturePosition =   327683
            Size            =   "2937;688"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   7
         Left            =   1800
         TabIndex        =   19
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bank"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   6
         Left            =   150
         TabIndex        =   17
         Top             =   3030
         Width           =   345
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Tag             =   "Attributes"
         Top             =   3240
         Width           =   4875
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "8599;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "L/C No."
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   5
         Left            =   3390
         TabIndex        =   15
         Top             =   1320
         Width           =   540
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   4
         Left            =   3360
         TabIndex        =   14
         Tag             =   "ItemName"
         Top             =   1530
         Width           =   1635
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2884;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Tag             =   "ItemUsage"
         Top             =   2130
         Width           =   4875
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "8599;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "City, Country"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   12
         Top             =   2475
         Width           =   975
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Tag             =   "Attributes"
         Top             =   2685
         Width           =   4875
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "8599;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   10
         Top             =   1920
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proformas"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   9
         Top             =   180
         Width           =   735
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Tag             =   "ItemName"
         Top             =   405
         Width           =   4875
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "8599;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Tag             =   "ItemUsage"
         Top             =   975
         Width           =   4875
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "8599;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AWB/B.L No."
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   1320
         Width           =   930
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Tag             =   "Attributes"
         Top             =   1530
         Width           =   1635
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2884;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orders"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   765
         Width           =   495
      End
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
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   6960
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Inspection Certificate"
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
         Left            =   15
         TabIndex        =   1
         Top             =   135
         Width           =   6765
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Inspection Certificate"
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
         Top             =   150
         Width           =   6750
      End
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   285
      Left            =   1815
      TabIndex        =   26
      Top             =   945
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
      Caption         =   " Customer                     Country                      Invoice No"
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
      Index           =   10
      Left            =   90
      TabIndex        =   25
      Top             =   705
      Width           =   5325
   End
   Begin MSForms.ComboBox cmbInvNo 
      Height          =   285
      Left            =   3525
      TabIndex        =   23
      Tag             =   "AutoCatID"
      Top             =   945
      Width           =   1890
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3334;503"
      ListWidth       =   3527
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
Attribute VB_Name = "frmInspection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ComInvoice As String

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
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_FPInvoices"
        
        .Parameters("@Customer").value = cmbCust.Tag
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
    End With
    
    cmbInvNo.Clear
    With rs
        Do Until .EOF
            cmbInvNo.AddItem .Fields(0) & ""
            
            .MoveNext
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
            
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbinvno_Change()
    Call cmbinvno_click
End Sub

Private Sub cmbInvNoq_Click()

    On Error GoTo err
    If cmbInvNo.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From InspectionCertificate Where ComInvoice='" & cmbInvNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        TBox(1) = ![BL] & ""
        OrderDT = ![DT] & ""
        
        TBox(4) = ![LC] & ""
        TBox(3) = ![Customer] & ""
        
        TBox(2) = ![City] & ""
        
        TBox(6) = ![Bank] & ""
        
        
        TBox(5) = ![Orders] & ""
        TBox(0) = ![Proformas] & ""
        
        .Close
        
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbinvno_click()

    On Error GoTo err
    If cmbInvNo.MatchFound = False Then Exit Sub
    Dim Orders As String
    Dim Proformas As String
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * From VrptProformas Where PInvoice='" & cmbInvNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        TBox(1) = ![AwbNo] & ""
        OrderDT = ![DT] & ""
        
        TBox(4) = ![PaymentTerms] & ""
        TBox(3) = ![CustCode] & ""
        
        TBox(2) = ![Country] & ""
        
        'TBox(6) = Replace(![Notify] & "", vbCrLf, " ")
        TBox(6) = Replace(![Address1] & "", vbCrLf, " ")
        
        For i = 0 To .RecordCount - 1
            If InStr(Orders, ![OrderNo] & "") = 0 Then
                Orders = Orders & ![OrderNo] & ","
            End If
            
            If Not InStr(Proformas, ![PInvoice] & "") Then
                Proformas = Proformas & ![PInvoice] & ","
            End If
            .MoveNext
        Next i
        
        TBox(5) = Left(Orders, Len(Orders) - 1)
        TBox(0) = Left(Proformas, Len(Proformas) - 1)
        
        .Close
        
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdAdd_Click()
        
        
    Call lok(False)
    ComInvoice = ""
    For i = 0 To TBox.Count - 1
        TBox(i) = ""
    Next i
    
End Sub
Private Sub lok(Locked As Boolean)

    FraAdd.Visible = Locked
    FraSave.Visible = Not Locked
    
    cmbInvNo.Visible = Locked
    cmbInvNoNew.Visible = Not Locked
    
    cmbInvNo.Locked = False
    
End Sub


Private Sub cmdCancel_Click()
    Call lok(True)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdedit_Click()

    ComInvoice = cmbInvNo
    FraSave.Visible = True
    FraAdd.Visible = False
    cmbInvNo.Locked = True
    'Call Lok(False)
    
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
    Call StartTrans(con)
    
    Dim InvNo As String
    If ComInvoice <> "" Then
        con.Execute "Delete From InspectionCertificate Where ComInvoice='" & ComInvoice & "'"
        InvNo = ComInvoice
    Else
        InvNo = cmbInvNoNew
    End If
    
    
    con.Execute "Insert Into InspectionCertificate(ComInvoice," & _
     "Proformas,Orders,BL,DT,LC,Customer,City,Bank) Values(" & _
     "'" & InvNo & "','" & TBox(0) & "','" & TBox(5) & _
     "','" & TBox(1) & "','" & OrderDT & "','" & TBox(4) & _
     "','" & TBox(3) & "','" & TBox(2) & "','" & TBox(6) & _
     "')"
    
    con.CommitTrans
    
    Call lok(True)
    Call FillCmbs
        
    If ComInvoice <> "" Then
        cmbInvNo = ComInvoice
        ComInvoice = ""
    End If
        
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()

    On Error GoTo err
    
    cmbCust.ListHeight = 2100
    
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where custCode<>'Hico'")
    Call FillCmbs
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub FillCmbs()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_FPInvoices"
        
        .Parameters("@Customer").value = cmbCust.Tag
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
    End With
    
    With rs
    
        cmbInvNo.Clear
        Do Until .EOF
            cmbInvNo.AddItem .Fields(0) & ""
            .MoveNext
        Loop
        .Close
        
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
