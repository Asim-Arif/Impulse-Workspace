VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmRequisitionOld 
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7725
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Steel"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Guage"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
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
         MaxLength       =   2
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
         Caption         =   "  Customer                                                  Country                                         Order"
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
         Picture         =   "frmRequisitionOld.frx":0000
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
         Picture         =   "frmRequisitionOld.frx":015E
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
         Left            =   5760
         TabIndex        =   9
         Top             =   465
         Width           =   1740
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         ForeColor       =   6244673
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3069;503"
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
         Picture         =   "frmRequisitionOld.frx":0270
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
         Picture         =   "frmRequisitionOld.frx":0382
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
         Left            =   3240
         TabIndex        =   5
         Top             =   465
         Width           =   2505
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         ForeColor       =   6244673
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "4419;503"
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
Attribute VB_Name = "frmRequisitionOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String
Public Sub EditOrder(TempOrd As String)
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    strOrderNo = TempOrd
    
    With rs
        .Open "Select * From VFSteelPurchOrd Where OrderNo='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        cmbOrders.Locked = True
        cmbOrders = ![OrderNo] & ""
        cmbVendor.ID = ![VendID]
        cmbLab.ID = ![LabID]
        TBox(0) = ![ExtraPer] & ""
        TBox(1) = ![PurchOrdNo] & ""
        TBox(1).Tag = ![Dt] & ""
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
        .Open "Select OrderNo,DT From FCustomerOrders Where CustCode+Country='" & cmbCust.Text & cmbCountry & "' and OrderNo Not In(Select OrderNo From FSteelPurchaseOrder)", con, adOpenForwardOnly, adLockReadOnly
        cmbOrders.Clear
        Do Until .EOF
            cmbOrders.AddItem ![OrderNo] & ""
            cmbOrders.List(cmbOrders.ListCount - 1, 1) = Format(![Dt], "dd-MMM-yyyy")
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
    Dim itm As ListItem
    With rs
        .Open "Select * From VFOrderSteelWeight Where OrderNo='" & cmbOrders.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        LV1.ListItems.Clear
        Do Until .EOF
            Set itm = LV1.ListItems.Add(, ![SteelUsed] & "'", ![SteelType] & "")
            itm.ListSubItems.Add , ![Gage] & "'", ![SteelGage] & ""
            'itm.SubItems(1) = ![Gage] & ""
            itm.SubItems(2) = ![TotalWeight] & ""
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
            .Open "Select Max(SubString(PurchOrdNo,6,4)) From FSteelPurchaseOrder Where Month(DT)='" & Month(FakeDT) & "' and Year(DT)='" & Year(FakeDT) & "'"
            iNextNo = Format(Val(.Fields(0) & "") + 1, "0000")
            .Close
        End With
        Set rs = Nothing
        
        TBox(1) = "HSPL-" & iNextNo & Format(FakeDT, "-MM-yy")
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
    
    frmPrevRpt.ShowReport "{VFOrderSteelWeight.OrderNo}='" & cmbOrders.Text & "'", rptPurchOrdForSteel
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
         
    On Error GoTo err
    Dim rs As New ADODB.Recordset
        
    cmbCust.ListHeight = 2000
    cmbVendor.ListHeight = 2000
    cmbLab.ListHeight = 2000
    cmbOrders.Locked = False
    
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
    Call cmbVendor.AddVals(con, "Name", "SteelVendors", "ID")
    Call cmbLab.AddVals(con, "Name", "TestLabs", "ID")
    
    Exit Sub
err:
    MsgBox err.Description
        
End Sub


Private Sub Form_Resize()
    Me.Left = (MainForm.Width - Me.Width) / 2
    Me.Top = (MainForm.Height - Me.Height) / 10
End Sub
Private Function Saved() As Boolean

    On Error GoTo err
    Dim strMillBatchNo As String
    Me.MousePointer = vbHourglass
    Saved = False
    
    
    
    strMillBatchNo = Int(Rnd * (3000 - (500 - 1)) + (500))
    strMillBatchNo = "HSL-" & strMillBatchNo & "/" & Right(Year(TBox(1).Tag), 2)
    Call StartTrans(con)
    
    If strOrderNo = "" Then
        con.Execute "Insert Into FSteelPurchaseOrder(OrderNo,LabID,VendID," & _
         "ExtraPer,PurchOrdNo,DT,MillBatchNo) Values('" & cmbOrders.Text & "'," & _
         cmbLab.ID & "," & cmbVendor.ID & "," & Val(TBox(0)) & _
         ",'" & TBox(1) & "','" & TBox(1).Tag & "','" & strMillBatchNo & "')"
    Else
        con.Execute "Update FSteelPurchaseOrder Set LabID=" & cmbLab.ID & _
         ",VendID=" & cmbVendor.ID & ",ExtraPer=" & Val(TBox(0)) & _
         " Where OrderNo='" & strOrderNo & "'"
    End If
    
    
    Dim iChNo As Integer
    
    
    Dim dWeight As Double
    Dim dWUsed As Double
    Dim dBal As Double
    Dim iRnd As Integer
    
    Dim dSW As Double
    'Dim dC As Double
    'Dim dSi As Double
    'Dim dMn As Double
    'Dim dP As Double
    'Dim dS As Double
    'Dim dNi As Double
    'Dim dCr As Double
    Dim dRndVal As Collection
    
    Dim rs As New ADODB.Recordset
    
    dWUsed = 0
    'Get Charge No.
    iChNo = Int(Rnd * (8000 - (2000 - 1)) + (2000))
    
    For i = 1 To LV1.ListItems.Count
    
        rs.Open "Select * From SteelTypeComps Where SteelID=" & Val(LV1.ListItems(i).Key), con, adOpenForwardOnly, adLockReadOnly
            
        dWeight = Val(LV1.ListItems(i).SubItems(2)) 'Total Steel
        dWeight = dWeight + (dWeight * (Val(TBox(0)) / 100)) 'Steel With %
        dWUsed = 0 'Reset To 0.
        
        Do While dWUsed < dWeight
            dBal = dWeight - dWUsed
            iRnd = Int(Rnd * (755 - (725 - 1)) + (725))
            If iRnd > dBal Then iRnd = dBal 'If Random Value is More Than Bal
            dSW = iRnd
            dC = Format((Rnd * (0.14 - (0.11)) + (0.11)), "#.00")
            dSi = Format((Rnd * (0.68 - (0.31)) + (0.31)), "#.00")
            dMn = Format((Rnd * (0.78 - (0.35)) + (0.35)), "#.00")
            
            dP = Format((Rnd * (0.4 - (0.3)) + (0.3)), "#.00")
            dS = 0.01 'ds= Format((Rnd * (0.14 - (0.11)) + (0.11)), "#.00")
            dNi = Format((Rnd * (0.7 - (0.3)) + (0.3)), "#.00")
            dCr = Format((Rnd * (13.1 - (12.1)) + (12.1)), "#.00")
            
            con.Execute "Insert Into FSteelPurchDetail(OrderNo,SteelUsed" & _
             ",Guage,StripeWeight,ChargeNo,C,Si,Mn,P,S,Ni,Cr)" & _
             " Values('" & cmbOrders & "','" & LV1.ListItems(i).Text & _
             "','" & LV1.ListItems(i).SubItems(1) & "'," & dSW & "," & _
             iChNo & "," & dC & "," & dSi & "," & dMn & "," & dP & _
             "," & dS & "," & dNi & "," & dCr & ")"
             
            'Increment Charge No. By One.
            iChNo = iChNo + 1
            
            dWUsed = dWUsed + dSW
        Loop
        
    Next i
    
    
    ''''''''''''''
    Dim TestRefNo As String
    Dim TestDT As Date
    Dim strBNo As String
    
    TestDT = DateAdd("d", 14, TBox(1).Tag)
    If Format(TestDT, "dddd") = "Sunday" Then TestDT = DateAdd("d", 1, TestDT)
    TestRefNo = Int(Rnd * (999 - (100 - 1)) + (100)) & "/" & Format(TestDT, "yyyy")
    
    strBNo = Format(TestDT, "ddMMyy")
    
    dC = Format((Rnd * (0.14 - (0.11)) + (0.11)), "#.00")
    dSi = Format((Rnd * (0.68 - (0.31)) + (0.31)), "#.00")
    dMn = Format((Rnd * (0.78 - (0.35)) + (0.35)), "#.00")
            
    dP = Format((Rnd * (0.4 - (0.3)) + (0.3)), "#.00")
    dS = 0.01 'ds= Format((Rnd * (0.14 - (0.11)) + (0.11)), "#.00")
    dNi = Format((Rnd * (0.7 - (0.3)) + (0.3)), "#.00")
    dCr = Format((Rnd * (13.1 - (12.1)) + (12.1)), "#.00")
            
    con.Execute "Insert Into FSteelPurchDetail1(OrderNo,OurRefNo,DT,BatchNo,C,Si,Mn,P,S,Ni,Cr) Values('" & _
     cmbOrders & "','" & TestRefNo & "','" & TestDT & "','" & strBNo & "'," & dC & "," & dSi & _
     "," & dMn & "," & dP & "," & dS & "," & dNi & "," & dCr & ")"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    con.CommitTrans
    
    strOrderNo = ""
    Saved = True
    Me.MousePointer = vbDefault
    Exit Function
    
err:
    Me.MousePointer = vbDefault
    MsgBox err.Description
    
End Function

