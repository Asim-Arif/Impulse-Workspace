VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmCovering 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Covering..."
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCovering.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7860
   Begin VB.Frame Frame2 
      Height          =   6255
      Left            =   75
      TabIndex        =   3
      Top             =   645
      Width           =   7725
      Begin VB.Frame Frame4 
         Caption         =   "Information On Lines :"
         ForeColor       =   &H00800000&
         Height          =   2880
         Left            =   105
         TabIndex        =   17
         Top             =   2640
         Width           =   7515
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Line 7 :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   13
            Left            =   150
            TabIndex        =   31
            Top             =   2490
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Line 6 :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   12
            Left            =   150
            TabIndex        =   30
            Top             =   2115
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Line 5 :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   11
            Left            =   150
            TabIndex        =   29
            Top             =   1755
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Line 4 :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   10
            Left            =   150
            TabIndex        =   28
            Top             =   1380
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Line 3 :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   7
            Left            =   150
            TabIndex        =   27
            Top             =   1020
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Line 2 :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   6
            Left            =   150
            TabIndex        =   26
            Top             =   645
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Line 1 :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   25
            Top             =   285
            Width           =   525
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   11
            Left            =   720
            TabIndex        =   24
            Tag             =   "ItemUsage"
            Top             =   2460
            Width           =   6660
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "11747;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   8
            Left            =   720
            TabIndex        =   23
            Tag             =   "ItemUsage"
            Top             =   1350
            Width           =   6660
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "11747;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   6
            Left            =   720
            TabIndex        =   22
            Tag             =   "ItemUsage"
            Top             =   610
            Width           =   6660
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "11747;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   5
            Left            =   720
            TabIndex        =   21
            Tag             =   "ItemUsage"
            Top             =   240
            Width           =   6660
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "11747;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   7
            Left            =   720
            TabIndex        =   20
            Tag             =   "ItemUsage"
            Top             =   980
            Width           =   6660
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "11747;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   9
            Left            =   720
            TabIndex        =   19
            Tag             =   "ItemUsage"
            Top             =   1720
            Width           =   6660
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "11747;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   10
            Left            =   720
            TabIndex        =   18
            Tag             =   "ItemUsage"
            Top             =   2090
            Width           =   6660
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "11747;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
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
         Height          =   615
         Left            =   4110
         TabIndex        =   10
         Top             =   5550
         Width           =   3510
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1755
            TabIndex        =   12
            Top             =   165
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
            Left            =   60
            TabIndex        =   11
            Top             =   165
            Width           =   1665
            Caption         =   "Print "
            PicturePosition =   327683
            Size            =   "2937;635"
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
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   75
         TabIndex        =   34
         Top             =   480
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
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   1800
         TabIndex        =   37
         Top             =   480
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
         Index           =   15
         Left            =   75
         TabIndex        =   36
         Top             =   240
         Width           =   5325
      End
      Begin MSForms.ComboBox cmbInvNo 
         Height          =   285
         Left            =   3510
         TabIndex        =   35
         Tag             =   "AutoCatID"
         Top             =   480
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "For Name"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   14
         Left            =   2625
         TabIndex        =   33
         Top             =   810
         Width           =   690
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   12
         Left            =   2610
         TabIndex        =   32
         Tag             =   "ItemUsage"
         Top             =   1035
         Width           =   2475
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4366;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "For"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   4
         Left            =   5160
         TabIndex        =   16
         Top             =   2055
         Width           =   240
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   4
         Left            =   5145
         TabIndex        =   15
         Tag             =   "ItemUsage"
         Top             =   2265
         Width           =   2475
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4366;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Form E"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   3
         Left            =   5160
         TabIndex        =   14
         Top             =   1365
         Width           =   495
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   3
         Left            =   5145
         TabIndex        =   13
         Tag             =   "ItemUsage"
         Top             =   1575
         Width           =   2475
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4366;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bank Address"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1365
         Width           =   975
      End
      Begin MSForms.TextBox TBox 
         Height          =   1005
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Tag             =   "ItemName"
         Top             =   1575
         Width           =   2475
         VariousPropertyBits=   -1467987941
         BorderStyle     =   1
         Size            =   "4366;1773"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   7
         Tag             =   "ItemUsage"
         Top             =   1035
         Width           =   2475
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4366;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer Bank"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   2
         Left            =   2655
         TabIndex        =   6
         Top             =   1365
         Width           =   1080
      End
      Begin MSForms.TextBox TBox 
         Height          =   1005
         Index           =   1
         Left            =   2625
         TabIndex        =   5
         Tag             =   "Attributes"
         Top             =   1575
         Width           =   2475
         VariousPropertyBits=   -1467987941
         BorderStyle     =   1
         Size            =   "4366;1773"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "L/C #"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   4
         Top             =   810
         Width           =   405
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
      Left            =   435
      TabIndex        =   0
      Top             =   0
      Width           =   6960
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Covering"
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
         Caption         =   "Covering"
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
Attribute VB_Name = "frmCovering"
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

Private Sub cmbInvNo1_Click()

    On Error GoTo err
    If cmbInvNo.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From V Where ComInvoice='" & cmbInvNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        TBox(0) = ![Bank] & ""
        TBox(1) = ![CustomerBank] & ""
        TBox(2) = ![LCNo] & ""
        TBox(3) = ![FormE] & ""
        TBox(4) = ![For] & ""
        TBox(5) = ![l1] & ""
        TBox(6) = ![l2] & ""
        TBox(7) = ![l3] & ""
        
        TBox(8) = ![l4] & ""
        TBox(9) = ![l5] & ""
        TBox(10) = ![l6] & ""
        TBox(11) = ![l7] & ""
        TBox(12) = ![NameOf] & ""
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
        
        
        TBox(1) = ![Address1] & ""
        'TBox(2) = ![LC] & ""
        TBox(3) = ![FormENo] & ""
        'TBox(4) = ![totamt] & ""
        TBox(5) = "Commercial Invoice in 4 Copies"
        TBox(6) = "Us Customs Invoice in 5 copies"
        TBox(7) = "Packing List in 4 copies"
        
        TBox(8) = "Original copy of Bill of Lading"
        TBox(9) = "Copy of Inspection Certificate"
        TBox(10) = "Form 'E' "
        TBox(11) = ""
        TBox(12) = "" 'For Name...
        .Close
        
        TBox(0) = con.Execute("Select Bank From Banks").Fields(0).value & ""
        
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

Private Sub cmdPrint_Click()

    On Error GoTo err
    'Dim rpt As New rptCovering
    
    'Dim f As New frmPrevRpt
    'f.ShowReport " {Covering.ComInvoice}='" & cmbInvNo & "' ", rpt
    
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
    Call StartTrans(con)
    
    Dim InvNo As String
    If ComInvoice <> "" Then
        con.Execute "Delete From Covering Where ComInvoice='" & ComInvoice & "'"
        InvNo = ComInvoice
    Else
        InvNo = cmbInvNoNew
    End If
    
    
    con.Execute "Insert Into Covering(ComInvoice," & _
     "Bank,CustomerBank,[For],LCNo,FormE,l1,l2,l3,l4" & _
     ",l5,l6,l7,NameOf) Values('" & InvNo & "','" & _
     TBox(0) & "','" & TBox(1) & "','" & TBox(4) & _
     "','" & TBox(2) & "','" & TBox(3) & "','" & _
     TBox(5) & "','" & TBox(6) & "','" & TBox(7) & _
     "','" & TBox(8) & "','" & TBox(9) & _
     "','" & Replace(TBox(10), "'", "''") & "','" & TBox(11) & "','" & TBox(12) & "')"
    
    con.CommitTrans
    
    Call lok(True)
    
        
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
    
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
