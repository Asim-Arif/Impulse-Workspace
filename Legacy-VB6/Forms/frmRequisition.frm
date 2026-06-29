VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmRequisition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Requisition form "
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form37"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   60
      TabIndex        =   3
      Top             =   570
      Width           =   7605
      Begin ComboList.Usercontrol1 cmbLab 
         Height          =   285
         Left            =   3945
         TabIndex        =   4
         Top             =   960
         Width           =   2430
         _ExtentX        =   4286
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
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   90
         TabIndex        =   12
         Top             =   465
         Width           =   3120
         _ExtentX        =   5503
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
      Begin ComboList.Usercontrol1 cmbVendor 
         Height          =   285
         Left            =   1485
         TabIndex        =   13
         Top             =   960
         Width           =   2430
         _ExtentX        =   4286
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
      Begin MSComctlLib.ListView LV1 
         Height          =   4095
         Left            =   90
         TabIndex        =   14
         Top             =   1290
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Steel"
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Kgs"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSForms.TextBox TBox 
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   17
         Top             =   960
         Width           =   1365
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         ScrollBars      =   2
         Size            =   "2408;503"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   240
         Left            =   90
         TabIndex        =   16
         Top             =   240
         Width           =   7410
         BackColor       =   11517387
         Caption         =   "  Customer                                                  Country                                  Invoice"
         Size            =   "13070;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label3 
         Height          =   240
         Left            =   90
         TabIndex        =   15
         Top             =   735
         Width           =   7410
         BackColor       =   11517387
         Caption         =   " Pch Ord No.         Steel Vendor                               Lab                                               %Increase"
         Size            =   "13070;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdOk 
         Height          =   360
         Left            =   4245
         TabIndex        =   11
         Top             =   5430
         Width           =   1605
         Caption         =   " Save & Close"
         PicturePosition =   327683
         Size            =   "2831;635"
         Picture         =   "frmRequisition.frx":0000
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdNext 
         Height          =   360
         Left            =   2610
         TabIndex        =   10
         Top             =   5430
         Width           =   1605
         Caption         =   "  Save & New"
         PicturePosition =   327683
         Size            =   "2831;635"
         Accelerator     =   78
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbOrders 
         Height          =   285
         Left            =   5385
         TabIndex        =   9
         Top             =   450
         Width           =   2115
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         ForeColor       =   6244673
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3731;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   165
         TabIndex        =   8
         Top             =   5415
         Visible         =   0   'False
         Width           =   1605
         Caption         =   "         Print"
         PicturePosition =   131072
         Size            =   "2831;635"
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
         Left            =   5895
         TabIndex        =   7
         Top             =   5430
         Width           =   1605
         Caption         =   "         Close"
         PicturePosition =   131072
         Size            =   "2831;635"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.TextBox TBox 
         Height          =   285
         Index           =   0
         Left            =   6390
         TabIndex        =   6
         Top             =   960
         Width           =   1110
         VariousPropertyBits=   679495707
         MaxLength       =   2
         BorderStyle     =   1
         ScrollBars      =   2
         Size            =   "1958;503"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   3225
         TabIndex        =   5
         Top             =   465
         Width           =   2145
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         ForeColor       =   6244673
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3784;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7605
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisition Form"
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
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   7335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisition Form"
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
         Height          =   390
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   135
         Width           =   7335
      End
   End
End
Attribute VB_Name = "frmRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String
Dim FromDHR As Boolean
Dim Success As Boolean
Public Sub EditOrder(TempOrd As String)
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    strOrderNo = TempOrd
    
    With rs
        .Open "Select * From VFSteelPurchOrd Where PInvoice='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        cmbOrders.Locked = True
        cmbOrders = ![PInvoice] & ""
        cmbVendor.ID = ![VendID]
        cmbLab.ID = ![LabID]
        TBox(0) = ![ExtraPer] & ""
        TBox(1) = ![PurchOrdNo] & ""
        TBox(1).Tag = ![DT] & ""
        .Close
    End With
    Set rs = Nothing
    Me.Show 1
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select PInvoice,DT From VFProformaList Where CustCode+Country='" & cmbCust.Text & cmbCountry & "' and PInvoice Not In(Select PInvoice From FSteelPurchaseOrder)", con, adOpenForwardOnly, adLockReadOnly
        cmbOrders.Clear
        Do Until .EOF
            cmbOrders.AddItem ![PInvoice] & ""
            cmbOrders.List(cmbOrders.ListCount - 1, 1) = Format(![DT], "dd-MMM-yyyy")
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
        .Open "Select Distinct Country From ForeignCustomers Where CustCode='" & cmbCust.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem .Fields(0) & ""
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbOrders_Change()
    Call cmbOrders_Click
End Sub

Private Sub cmbOrders_Click()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "Select * From VFOrderSteelWeight Where PInvoice='" & cmbOrders.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        LV1.ListItems.Clear
        Do Until .EOF
            Set ITM = LV1.ListItems.Add(, ![SteelUsed] & "'" & .AbsolutePosition, ![SteelType] & "")
            'ITM.ListSubItems.Add , ![Gage], ![SteelGage] & ""
            'itm.SubItems(1) = ![Gage] & ""
            ITM.SubItems(1) = ![TotalWeight] & ""
            'ITM.SubItems(3) = ![UnitQty] & ""
            .MoveNext
        Loop
        
        .Close
    End With
    'If Not Editing
    Dim FakeDT As Date
    'Dim rs As New ADODB.Recordset
    Dim iNextNo As String
    
    If strOrderNo = "" Then
        FakeDT = DateAdd("d", 2, cmbOrders.List(cmbOrders.ListIndex, 1))
        If Format(FakeDT, "dddd") = "Sunday" Then
            FakeDT = DateAdd("d", 2, FakeDT)
        End If
        'Now Get This Month's Max #.
        With rs
            .Open "Select Max(SubString(PurchOrdNo,4,4)) From FSteelPurchaseOrder Where Month(DT)='" & Month(FakeDT) & "' and Year(DT)='" & Year(FakeDT) & "'"
            iNextNo = Format(Val(.Fields(0) & "") + 1, "0000")
            .Close
        End With
        Set rs = Nothing
        
        TBox(1) = "HS-" & iNextNo & Format(FakeDT, "-MM-yy")
        TBox(1).Tag = FakeDT
    End If
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub Combo1_Click()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "select distinct(orderno),Odate from orders where custcode='" & Combo1.Text & "' order by Odate desc", con, adOpenForwardOnly, adLockReadOnly
        List1.Clear
        Do Until .EOF
            List1.AddItem Trim(.Fields(0).value)
            .MoveNext
        Loop
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

Private Sub cmdNext_Click()
    If Saved Then
        Unload Me
        Me.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then
        Unload Me
    End If
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Load frmPrevRpt
    
    rptPurchOrdForSteel.FormulaFields(5).Text = Val(TBox(0))
    'rptPurchOrdForSteel.FormulaFields(2).Text = "'" & strFromTo & "'"
    
    'frmPrevRpt.ShowReport "{VFOrderSteelWeight.OrderNo}='" & cmbOrders.Text & "'", rptPurchOrdForSteel
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
         
    On Error GoTo err
    FromDHR = False
    Success = False
    
    Dim rs As New ADODB.Recordset
        
    cmbCust.ListHeight = 2000
    cmbVendor.ListHeight = 2000
    cmbLab.ListHeight = 2000
    cmbOrders.Locked = False
    
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where custCode<>'Hico'")
    Call cmbVendor.AddVals(con, "Name", "SteelVendors", "ID")
    Call cmbLab.AddVals(con, "Name", "TestLabs", "ID")
    
    Exit Sub
err:
    MsgBox err.Description
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If FromDHR And Not Success Then
        If MsgBox("This Will Also Close DHR Wizard." & vbNewLine & "Do You Want To Continue ?", vbQuestion + vbYesNo) = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Resize()
    'Me.Left = (MainForm.Width - Me.Width) / 2
    'Me.Top = (MainForm.Height - Me.Height) / 10
End Sub
Private Function Saved() As Boolean

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim DTRec As Date
    
    With rs
        .Open "Select DTForReq From DHRDetail1 Where PInvoice='" & cmbOrders.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount = 0 Then
            MsgBox "Create DHR First.", vbInformation
            Exit Function
        Else
            DTRec = ![DTForReq] & ""
        End If
        .Close
    End With
    
    Dim strMillBatchNo As String
    
    Me.MousePointer = vbHourglass
    Saved = False

    Randomize
    
    strMillBatchNo = Int(Rnd * (3000 - (500 - 1)) + (500))
    strMillBatchNo = "INM-" & strMillBatchNo & "/" & Right(Year(TBox(1).Tag), 2)
    
    
    If Not FromDHR Then Call StartTrans(con)
    
    If strOrderNo = "" Then
        con.Execute "Insert Into FSteelPurchaseOrder(PInvoice,LabID,VendID," & _
         "ExtraPer,PurchOrdNo,DT,MillBatchNo,RecDT) Values('" & cmbOrders.Text & "'," & _
         cmbLab.ID & "," & cmbVendor.ID & "," & Val(TBox(0)) & _
         ",'" & TBox(1) & "','" & TBox(1).Tag & "','" & strMillBatchNo & "','" & _
         DTRec & "')"
    Else
        con.Execute "Delete from FSteelPurchDetail Where PInvoice='" & strOrderNo & "'"
        con.Execute "Delete From FPurchOrderComps1 Where RefID In(Select EntryID From FSteelPurchDetail1 Where PInvoice='" & strOrderNo & "')"
        
        con.Execute "Update FSteelPurchaseOrder Set LabID=" & cmbLab.ID & _
         ",VendID=" & cmbVendor.ID & ",ExtraPer=" & Val(TBox(0)) & _
         ",PurchOrdNo='" & TBox(1) & "' " & _
         " Where PInvoice='" & strOrderNo & "'"
    End If
    
    '--------------------------------------
    Call DoForSaleInvoice
    '--------------------------------------
    
    
    Dim iChNo As Integer
    
    
    Dim dWeight As Double, dWeight1 As Double
    Dim dWUsed As Double
    Dim dBal As Double
    Dim iRnd As Double
    
    Dim dSW As Double
    
    Dim dRndVal As Double
    
    Dim iRefID As Double
    
    dWUsed = 0
    'Get Charge No.
    Randomize
    iChNo = Int(Rnd * (8000 - (2000 - 1)) + (2000))
    
    For i = 1 To LV1.ListItems.Count
        
        
    'rs.Open "Select * From VSteelTypeComps Where SteelID=" & Val(LV1.ListItems(i).Key), con, adOpenForwardOnly, adLockReadOnly
        rs.Open "Select * From VSteelTypeComps Where SteelID=" & Val(LV1.ListItems(i).Key), con, adOpenForwardOnly, adLockReadOnly
        
        dWeight = Val(LV1.ListItems(i).SubItems(1)) 'Total Steel
        
        'dWeight1 = dWeight + (dWeight * (Val(TBox(0)) / 100)) 'Steel With %
        dWeight = dWeight + (dWeight * (Val(TBox(0)) / 100)) 'Steel With %
        'dWeight = Int(Rnd * (dWeight1 - (dWeight - 1)) + (dWeight))
        
        dWUsed = 0 'Reset To 0.
        
        Do While dWUsed < dWeight
            If rs.RecordCount > 0 Then rs.MoveFirst
            dBal = dWeight - dWUsed
            
            Randomize
            'iRnd = Int(Rnd * (Val(rs![StripeMaxW]) - (Val(rs![StripeMinW]) - 1)) + (Val(rs![StripeMinW])))
            iRnd = Int(Rnd * (Val(400) - (Val(385) - 1)) + (Val(385)))
            
            'The Following Line is commented out To keep the weight of Last Stripe Consistent with other stripes
            'If iRnd > dBal Then iRnd = dBal 'If Random Value is More Than Bal
            dSW = iRnd
            
            con.Execute "Insert Into FSteelPurchDetail(PInvoice,SteelUsed" & _
             ",Guage,StripeWeight,ChargeNo)" & _
             " Values('" & cmbOrders & "','" & Val(LV1.ListItems(i).Key) & _
             "','" & LV1.ListItems(i).ListSubItems(1).Key & "'," & dSW & "," & _
             iChNo & ")"
            
            iRefID = con.Execute("Select Max(EntryID) From FSteelPurchDetail").Fields(0).value
            
            Do Until rs.EOF
                dRndVal = Format((Rnd * (Val(rs![UVal] & "") - (Val(rs![lVal] & ""))) + (Val(rs![lVal] & ""))), "#.00")
                
                con.Execute "Insert Into FPurchOrderComps(RefID,CompID" & _
                 ",CompVal) Values(" & iRefID & "," & rs![CompID] & "," & _
                 dRndVal & ")"
    
                rs.MoveNext
            Loop
            
            'Increment Charge No. By One.
            iChNo = iChNo + 1
            
            dWUsed = dWUsed + dSW
        Loop
        
        rs.Close
        
    Next i
    
    
    ''''''''''''''
    
    Dim TestRefNo As String
    Dim TestDT As Date
    Dim MaterialTestDT As Date
    Dim strBNo As String
    Dim iRndRow As Integer
    
    
    TestDT = DateAdd("d", 1, DTRec)
    
    TestDT = CheckHoliday(TestDT)
    
    MaterialTestDT = DateAdd("d", 2, TestDT)
    MaterialTestDT = CheckHoliday(MaterialTestDT)
    
    'If Format(TestDT, "dddd") = "Sunday" Then TestDT = DateAdd("d", 1, TestDT)
    TestRefNo = Int(Rnd * (999 - (100 - 1)) + (100)) & "/" & Format(TestDT, "yyyy")
    
    strBNo = Format(TestDT, "ddMMyy")
    
    con.Execute "Insert Into FSteelPurchDetail1(PInvoice,OurRefNo,DT,BatchNo,MaterialTestDT) Values('" & _
     cmbOrders & "','" & TestRefNo & "','" & TestDT & "','" & strBNo & "','" & MaterialTestDT & "')"
     
    Randomize
    iRndRow = Int(Rnd * (LV1.ListItems.Count - (1 - 1)) + (1))
    
    rs.Open "Select * From VSteelTypeComps Where SteelID=" & Val(LV1.ListItems(iRndRow).Key), con, adOpenForwardOnly, adLockReadOnly
    
    iRefID = con.Execute("Select Max(EntryID) From FSteelPurchDetail1").Fields(0).value
    
    Do Until rs.EOF
        Randomize
        dRndVal = Format((Rnd * (Val(rs![UVal] & "") - (Val(rs![lVal] & ""))) + (Val(rs![lVal] & ""))), "#.00")
                
        con.Execute "Insert Into FPurchOrderComps1(RefID,CompID" & _
         ",CompVal) Values(" & iRefID & "," & rs![CompID] & "," & _
         dRndVal & ")"

        rs.MoveNext
    Loop
    
    rs.Close
    
    con.CommitTrans
    Success = True
    Set rs = Nothing
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    strOrderNo = ""
    Saved = True
    Me.MousePointer = vbDefault
    Exit Function
    
err:
    Me.MousePointer = vbDefault
    MsgBox err.Description
    
End Function
Private Sub DoForSaleInvoice()

    Dim rs As New ADODB.Recordset
    Dim rndVal As Integer
    Dim dW As Double
    With rs
        .Open "Select * From VFSteelPurchOrdGageWise Where PInvoice='" & cmbOrders.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
        
            'rndVal = Int(Rnd * (Val(2) - (Val(1) - 1)) + (Val(1)))
            'If rndVal = 1 Then
            'Else
            'End If
            'dW = ![TotalWeight] + (![TotalWeight] * (Val(TBox(0)) / 100))
            If IsNull(![TotalWeightEP]) Then
                dW = ![TotalWeight] + (![TotalWeight] * (Val(TBox(0)) / 100))
            Else
                dW = ![TotalWeightEP] + (![TotalWeightEP] * (Val(TBox(0)) / 100))
            End If
            'dW = ![TotalWeightEP] + (![TotalWeightEP] * (Val(TBox(0)) / 100))
            
            con.Execute "Insert Into FSteelPurchDetail2(PInvoice,SteelID,GageID,RealWeight,SaleInvWeight) Values('" & cmbOrders.Text & "'," & _
             ![SteelUsed] & ",'" & ![Gage] & "'," & dW & "," & dW + 1 & ")"
             
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub

Public Function ShowForm() As Boolean
    FromDHR = True
    Me.Show 1
    ShowForm = Success
End Function
