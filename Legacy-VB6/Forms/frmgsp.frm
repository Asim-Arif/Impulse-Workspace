VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmGSP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "G S P"
   ClientHeight    =   6165
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
   Icon            =   "frmgsp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   5490
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
      Begin VB.TextBox txtTransport 
         Appearance      =   0  'Flat
         Height          =   1470
         Left            =   4290
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   1650
         Width           =   2835
      End
      Begin VB.TextBox txtCustAddress 
         Appearance      =   0  'Flat
         Height          =   1470
         Left            =   75
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   1650
         Width           =   4155
      End
      Begin VB.TextBox txtOrigin 
         Appearance      =   0  'Flat
         Height          =   840
         Left            =   4290
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   3915
         Width           =   2835
      End
      Begin VB.TextBox txtPlace 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4290
         TabIndex        =   23
         Text            =   "Sialkot"
         Top             =   1110
         Width           =   2835
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         Height          =   1410
         Left            =   75
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   3345
         Width           =   4155
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
         Top             =   4770
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
            Picture         =   "frmgsp.frx":014A
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
            Picture         =   "frmgsp.frx":025C
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
      Begin MSComCtl2.DTPicker DT 
         Height          =   315
         Left            =   75
         TabIndex        =   25
         Top             =   1110
         Width           =   4140
         _ExtentX        =   7303
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
         CustomFormat    =   "dddd dd-MMM-yyyy"
         Format          =   122290179
         CurrentDate     =   37055
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
         Top             =   4770
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
            Picture         =   "frmgsp.frx":036E
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Origin Criterion"
         Height          =   195
         Index           =   3
         Left            =   4320
         TabIndex        =   28
         Top             =   3705
         Width           =   1080
      End
      Begin MSForms.ComboBox cmbOCountry 
         Height          =   285
         Left            =   4290
         TabIndex        =   26
         Top             =   3345
         Width           =   2835
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5001;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Place"
         Height          =   195
         Index           =   2
         Left            =   4320
         TabIndex        =   24
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Country Criteria"
         Height          =   195
         Index           =   1
         Left            =   4320
         TabIndex        =   22
         Top             =   3165
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Description Of Goods"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   21
         Top             =   3165
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   19
         Top             =   900
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Customer Address"
         Height          =   195
         Index           =   7
         Left            =   105
         TabIndex        =   18
         Top             =   1440
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Transport && Route"
         Height          =   195
         Index           =   8
         Left            =   4320
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
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
         BackStyle       =   0  'Transparent
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
         Caption         =   "G S P"
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
         Caption         =   "G S P"
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
Attribute VB_Name = "frmGSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PInvoice As String

Public Sub ShowForm(CustCode As String, Country As String, InvNo As String)

    cmbCust.ID = CustCode
    cmbCountry = Country
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select CustomInvoice From GSP Where CustomInvoice='" & InvNo & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount = 0 Then
            cmdAdd_Click
            cmbInvNoNew = InvNo
        Else
            cmbInvNo = InvNo
        End If
        .Close
    End With
    Set rs = Nothing
    Me.Show , frmFlash
    
End Sub
    
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
        .Open "Select CustomInvoice From VGSP Where CustCode+Country='" & cmbCust.ID & cmbCountry.Text & "' ORDER BY CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
        cmbInvNo.Clear
        Do Until .EOF
            cmbInvNo.AddItem .Fields(0) & ""
            'cmbInvNo.List(cmbInvNo.ListCount - 1, 1) = .Fields(1) & ""
            .MoveNEXT
        Loop
        .Close
        
        .Open "Select CustomInvoice From CustomInvoice Where CustomInvoice Not In (Select CustomInvoice From GSP) and CustCode+Country='" & cmbCust.ID & cmbCountry.Text & "' ORDER BY CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
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
    Dim strDesc As String
    Dim strDischarge As String
    Dim strBLDetail As String
    With rs
    
        If cmbCust.ID <> "DECATHLON" Then
            .Open "Select * From VTotalCartons Where CustomInvoice='" & cmbInvNoNew & "'", con, adOpenForwardOnly, adLockReadOnly
            If .RecordCount > 0 Then
                txtDescription = ![TotalCartons] & " -Cartons Containing " & ![TotalQty] & " Pcs"
            End If
            .Close
        End If
        
        .Open "Select ShippingMode,AWBNo,AWBNoDT,ComConsignee,Vessel,ManDischarge,DisCity,DisCountry,LoadCity,LoadCountry From VCustomInvoice1 Where CustomInvoice='" & cmbInvNoNew & "'", con, adOpenForwardOnly, adLockReadOnly
        txtCustAddress = ![ComConsignee] & ""
        If !ManDischarge & "" = "" Then
            strDischarge = !DisCity & " " & !DisCountry
        Else
            strDischarge = !ManDischarge & ""
        End If
        If !ShippingMode = "By Air" Then
            strBLDetail = "under AWB No. " & !AWBNo & " Dated:" & Format(!AWBNoDT, "dd-MMM-yyyy")
        Else
            strBLDetail = "under B/L No. " & !AWBNo & " Dated:" & Format(!AWBNoDT, "dd-MMM-yyyy")
        End If
        'txtTransport = "Under AWB No. : " & ![AWBNo] & " Dated : " & Format(![AWBNoDT], "dd-MMM-yyyy") & vbNewLine & " Through"
        txtTransport = "Through (" & !Vessel & " " & !ShippingMode & vbNewLine & _
                        "from (" & !LoadCity & " " & !LoadCountry & " to (" & strDischarge & vbNewLine & _
                        strBLDetail
                        
        .Close
        
        If cmbCust.ID = "DECATHLON" Then
            .Open "Select OrderNo,ItemName,Qty,Unit From VCustomInvoice Where CustomInvoice='" & cmbInvNoNew & "'", con, adOpenForwardOnly, adLockReadOnly
            txtDescription = ""
            strDesc = ""
            Do Until .EOF
                If strDesc <> "" Then strDesc = strDesc & vbNewLine & vbNewLine
                strDesc = strDesc & "Total " & !Qty & " " & !Unit & " OF " & !ItemName & " (ORDER # " & !OrderNo & ")"
                .MoveNEXT
            Loop
            .Close
            txtDescription = strDesc
        ElseIf UCase(cmbCust.ID) = "UHLSPORT" Then
            .Open "Select DISTINCT OrderNo From VCustomInvoice Where CustomInvoice='" & cmbInvNoNew & "'", con, adOpenForwardOnly, adLockReadOnly
            txtDescription = ""
            strDesc = ""
            If strDesc <> "" Then strDesc = strDesc & vbNewLine & vbNewLine
            strDesc = strDesc & " Orders : "
            Do Until .EOF
                strDesc = strDesc & !OrderNo & ","
                .MoveNEXT
            Loop
            .Close
            txtDescription = strDesc
        End If
    End With
    
    Set rs = Nothing
    txtPlace = "Sialkot"
    
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
        '.Open "Select * From VShippingInstructions Where CustomInvoice='" & cmbInvNo & "'", con, adOpenForwardOnly, adLockReadOnly
        .Open "Select * From GSP Where CustomInvoice='" & cmbInvNo & "'", con, adOpenForwardOnly, adLockReadOnly
        DT = ![GSPDT]
        txtPlace = ![Place] & ""
        txtCustAddress = ![CustAdd] & ""
        txtTransport = ![Transport] & ""
        txtDescription = ![Description] & ""
        cmbOCountry = ![OriginCountry] & ""
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbOCountry_Change()
    Call cmbOCountry_Click
End Sub

Private Sub cmbOCountry_Click()
    If cmbOCountry.MatchFound = False Then Exit Sub
    txtOrigin = cmbOCountry.List(cmbOCountry.ListIndex, 1)
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
    If cmbInvNo.List(cmbInvNo.ListIndex, 1) Then
        MsgBox "This Record Has Been Locked.", vbInformation
        Exit Sub
    End If
        
    PInvoice = cmbInvNo
    
    FraSave.Visible = True
    FraAdd.Visible = False
    
    cmbInvNo.Locked = True
    
    Call lok(False)
    cmbInvNo.Visible = True
    cmbInvNoNew.Visible = False
    
    
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    If cmbInvNo.MatchFound = False Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
'    If UCase(cmbCust.ID) = "DECATHLON" And UCase(cmbCountry.Text) = "CHINA" Then
'        Set rpt = rptApp.OpenReport(rptExportPath & "\rptGSPChina.rpt")
'    Else
        Dim strRet As String
        strRet = ShowPopUpMenu1("Print &White GSP", "-", "Print &Green GSP")
        Select Case strRet
            Case "Print &White GSP"
                Set rpt = rptApp.OpenReport(RptPath & "\rptGSPWhite.rpt")
            Case "Print &Green GSP"
                Set rpt = rptApp.OpenReport(RptPath & "\rptGSPGreen.rpt")
            Case Else
                Exit Sub
        End Select
    
    
    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt
    f.ShowReport "{VGSP.CustomInvoice}='" & cmbInvNo.Text & "' ", rpt
    
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    Dim sTempInv As String
    Call StartTrans(con)
    
    If PInvoice <> "" Then
        con.Execute "Delete From GSP Where CustomInvoice='" & PInvoice & "'"
    End If
    sTempInv = cmbInvNoNew
    
    con.Execute "Insert Into GSP(CustomInvoice,GSPDT" & _
     ",Place,CustAdd,Transport,Description,OriginCountry)" & _
     " Values('" & _
     IIf(PInvoice = "", cmbInvNoNew, cmbInvNo) & "','" & DT & _
     "','" & txtPlace & "','" & txtCustAddress & "','" & txtTransport & "','" & _
     txtDescription & "','" & cmbCountry & "')"
     
    con.CommitTrans
    
    If PInvoice = "" Then
        Call cmbCountry_Click
        cmbInvNo = sTempInv
    End If
    
    Call lok(True)
        
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub Form_Load()

    On Error GoTo err
    
    cmbCust.ListHeight = 2100
    DT = GetServerDate
    Call cmbCust.AddVals(con, "Distinct CustCode", "ForeignCustomers", "CustCode")
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * From GSPCodes Order By Country", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            cmbOCountry.AddItem ![Country] & ""
            cmbOCountry.List(cmbOCountry.ListCount - 1, 1) = ![code] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

