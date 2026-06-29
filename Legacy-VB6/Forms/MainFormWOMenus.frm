VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm MainFormWOMenus 
   BackColor       =   &H8000000C&
   Caption         =   "Contractor Package"
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   1020
   ClientWidth     =   8880
   Icon            =   "MainFormWOMenus.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4110
      Top             =   3210
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   4650
      Top             =   2295
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1950
      Top             =   2340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormWOMenus.frx":0ABA
            Key             =   "Cont"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormWOMenus.frx":0FBC
            Key             =   "PaySheet"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormWOMenus.frx":105B
            Key             =   "PayEntry"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormWOMenus.frx":1139
            Key             =   "Issue"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormWOMenus.frx":120E
            Key             =   "Recieve"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormWOMenus.frx":12B0
            Key             =   "Items"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainFormWOMenus.frx":1393
            Key             =   "ContLedger"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1005
      ButtonWidth     =   1984
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Contractors"
            Key             =   "Cont"
            Object.ToolTipText     =   "Contractors List"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Items Info"
            Key             =   "Items"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Issue Items"
            Key             =   "Issue"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rec. Items"
            Key             =   "Recieve"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Adv. Payment"
            Key             =   "PayEntry"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pay. Sheet"
            Key             =   "PaySheet"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cont Ledger"
            Key             =   "ContLedger"
            Object.ToolTipText     =   "Contractor Ledger"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4515
      Top             =   3195
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Please Select Payroll Database to Connect."
      Filter          =   """Payroll Database|Payroll.mdb"""
   End
End
Attribute VB_Name = "MainFormWOMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dashBankLIst_Click()
    frmBankList.Show
End Sub

Private Sub MDIForm_Load()
    
    On Error GoTo err
    'mnucont.Visible = CONTRACTORS_ENABLED
    'Toolbar1.Buttons("Cont").Visible = CONTRACTORS_ENABLED
    'Toolbar1.Buttons("ContLedger").Visible = CONTRACTORS_ENABLED
    Dim i As Long
    With Toolbar1
        Set Toolbar1.ImageList = ImageList1
        For i = 1 To .Buttons.Count
            If Not .Buttons(i).Style = tbrSeparator Then
                .Buttons(i).Image = ImageList1.ListImages(.Buttons(i).Key).Key
            End If
        Next i
    End With
    
    mnupopemp.Visible = False
    mnupopdept.Visible = False
    mnupopPorts.Visible = False
    
    Call SetProcMenuTags
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub SetProcMenuTags()

    
    
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set con = Nothing
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
   ' Con.RollbackTrans
    'If Con.State = 1 Then Con.Close
    Set con = Nothing
End Sub

Private Sub mnuAssHRrpt_Click()
    Call ShowAssHRrpt
End Sub

Private Sub mnuAssItems_Click()
    Call ContractorList.showItems
End Sub

Private Sub mnuAuth_Click()
    frmAuthorizeIssRcv.Show
End Sub

Private Sub mnuCalc_Click()

    On Error GoTo err
    Call Shell("calc.exe")
    Exit Sub
err:
    MsgBox "Cannot Run Calculater Due To Error: " & vbNewLine & err.Description, vbInformation

End Sub

Private Sub mnuComp_Click()
    frmComponents.Show 1
End Sub

Private Sub mnuCompCatal_Click()
    frmItems.Show
End Sub

Private Sub mnuCompositions_Click()
    frmCompositions.Show
End Sub

Private Sub mnuCustCatal_Click()
    frmCustCatal.Show
End Sub

Private Sub mnuCustList_Click()
    frmFCustomerList.Show
End Sub

Private Sub mnuCustListrpt_Click()
    Call ShowCustListrpt
End Sub

Private Sub mnuCustOrdList_Click()
    frmFCustOrderList.Show
End Sub

Private Sub mnuCustProformaList_Click()
    frmFProformaList.Show
End Sub

Private Sub mnuDailyRpts_Click()
    frmDailyIssRcvRpt.Show
End Sub

Private Sub mnuDeptEdit_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.mnuEdit(mnuDeptEdit.Tag)
    Else
        Call ContDeptList.mnuEdit(mnuDeptEdit.Tag)
    End If
End Sub

Private Sub mnuDeptInact_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.MakeInactive
    Else
        Call ContDeptList.MakeInactive
    End If
End Sub

Private Sub mnuDeptList_Click()
    DeptList.Show
End Sub

Private Sub mnuDeptnew_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.mnuAddDept
    Else
        Call ContDeptList.mnuAddDept
    End If
End Sub

Private Sub mnuDeptPrint_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.PrintDeptList
    Else
        Call ContDeptList.PrintDeptList
    End If
End Sub

Private Sub mnuDeptRefresh_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.RefreshIt
    Else
        Call ContDeptList.RefreshIt
    End If
End Sub

Private Sub mnuDrilSS_Click()
    Call ShowStockStatus(3)
End Sub

Private Sub mnuEmpEdit_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.mnuEdit(mnuEmpEdit.Tag)
    Else
        Call ContractorList.mnuEdit(mnuEmpEdit.Tag)
    End If
End Sub

Private Sub mnuempNewDept_Click()
    If Me.mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.mnuAddDept
    Else
        Call ContractorList.mnuAddDept
    End If
End Sub
Private Sub mnuEmpNewEmp_Click()
    If Me.mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.mnuAdd
    Else
        Call ContractorList.mnuAdd
    End If
End Sub


Private Sub mnuEmpPrint_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.PrintEmpList
    Else
          Call ContractorList.PrintEmpList
    End If
End Sub

Private Sub mnuempRefresh_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.RefreshIt
    Else
        Call ContractorList.RefreshIt
    End If
End Sub
Private Sub mnuExit_Click()
    If con.State = 1 Then con.Close
    Set con = Nothing
    End
End Sub

Private Sub mnuFDAProcess_Click()
    frmFDAProcesses.Show 1
End Sub

Private Sub mnuFillSS_Click()
    Call ShowStockStatus(4)
End Sub

Private Sub mnuForeignCatal_Click()
    frmForeignCatalog.Show 1
End Sub

Private Sub mnuForgSS_Click()
    Call ShowStockStatus(1)
End Sub

Private Sub mnuFuncSS_Click()
    Call ShowStockStatus(5)
End Sub

Private Sub mnuGuage_Click()
    frmGages.Show 1
End Sub

Private Sub mnuHICOResources_Click()
    frmHICOR.Show
End Sub

Private Sub mnuHolidays_Click()
    Load Holidays
    'Holidays.Move 3000, 500
    With Holidays
        .Move (Me.Width - .Width) / 2, Me.Top + 500
        .Show
    End With
End Sub

Private Sub mnuHRGroups_Click()
    Dim f As New frmItemGroups
    Load f
    Call f.ShowForm("HRGroups", "Hico Resources Groups", "Hico Resource Name")
End Sub

Private Sub mnuHRrpt_Click()
    Call ShowHRrpt
End Sub

Private Sub mnuIssDril_Click()
    Call ShowIssForm(3)
    
End Sub

Private Sub mnuIssFill_Click()
    Call ShowIssForm(4)
End Sub

Private Sub mnuIssForg_Click()
    Call ShowIssForm(1)
End Sub

Private Sub mnuIssFunc_Click()
    Call ShowIssForm(5)
End Sub

Private Sub mnuIssFunt_Click()
    Call ShowIssForm(9)
End Sub

Private Sub mnuIssHeat_Click()
    Call ShowIssForm(6)
End Sub

Private Sub mnuIssHold_Click()
    Call ShowIssForm(13)
End Sub

Private Sub mnuIssMach_Click()
    Call ShowIssForm(2)
End Sub

Private Sub mnuIssPack_Click()
    Call ShowIssForm(12)
End Sub

Private Sub mnuIsspass_Click()
    Call ShowIssForm(8)
End Sub

Private Sub mnuIssPoli_Click()
    Call ShowIssForm(7)
End Sub

Private Sub mnuIssPoRe_Click()
    Call ShowIssForm(10)
End Sub

Private Sub mnuIssRepair_Click()
    frmVendIssRepair.Show
End Sub

Private Sub mnuIssRepairPoli_Click()
    frmVendIssRepair.Show
End Sub

Private Sub mnuIssRM_Click()
    frmRMIss.Show
End Sub

Private Sub mnuIssRpts_Click()
    Dim f As New frmRptIssRcv
    Call f.ShowForIss
End Sub

Private Sub mnuIssUltr_Click()
    Call ShowIssForm(11)
End Sub

Private Sub mnuItemsFinQuality_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("ItemsFinQuality", "Items Fin. Quality", "Finishing Quality")
End Sub

Private Sub mnuItemTypes_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("ItemTypes", "Item Types", "Item Types")
End Sub

Private Sub mnuItmGrps_Click()
    Dim f As New frmItemGroups
    Load f
    Call f.ShowForm("ItemGroups", "Item Groups", "Group Name")
End Sub

Private Sub mnuLabs_Click()
    frmTestLabs.Show
End Sub

Private Sub mnuLogoff_Click()

    If MsgBox("Are You Sure To Log Off ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Unload MainForm
    LoggingOff = True
    Call Main
    
End Sub

Private Sub mnuLotProc_Click()
    frmInitialLotProcess.Show
End Sub

Private Sub mnuMachSS_Click()
    Call ShowStockStatus(2)
End Sub

Private Sub mnuMake_Click()

    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.MakeInactive
    Else
        Call ContractorList.MakeInactive
    End If
    
End Sub

Private Sub mnuMatList_Click()
    frmMaterial.Show
End Sub

Private Sub mnuNewOrd_Click()
    frmFNewCustOrder.Show
End Sub

Private Sub mnuNewProforma_Click()
    frmFNewProforma.Show
End Sub

Private Sub mnuNewStockOrder_Click()
    frmFNewStockOrder.Show
End Sub

Private Sub mnuOrdStatus_Click()
    frmFCustOrdStatus.Show
End Sub

Private Sub mnuOrdStatus1_Click()
    frmFCustOrdStatus1.Show
End Sub

Private Sub mnuOrdStatusComp_Click()
    frmFCustOrdStatusComp.Show
End Sub

Private Sub mnuPassword_Click()
    
    With frmUserInfo
        .Show_Option = Opt_ChangeUserPassword
        .Tag = CurrentUserName
        .Show 1
    End With
    
End Sub

Private Sub mnuPer_Click()
    frmPercentRanges.Show
End Sub

Private Sub mnuPOrderList_Click()
    frmRMPOrdList.Show
End Sub

Private Sub mnuPORpt_Click()
    On Error GoTo err
    
    Dim f As New frmPrint
    Load f
    With f
        Set .rpt = New rptPOrder
        .cmbCust.ClearVals
        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
        .cmbCust.AddItem "<All Venders>", "0"
        .cmbCust.AddVals con, "VendID1", "Venders", "VendID"
        .SelFor = "{VPOrdersrpt.RecieptID}='"
        .JobCondField = " VendID "
        .JobField = " Distinct RecieptID "
        .JobIDField = " RecieptID "
        .JobTable = " VPOrdersrpt "
        .TheEnd = "'"
        .Show , MainForm
    End With
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPorts_Click()
    frmPorts.Show
End Sub

Private Sub mnuProcess_Click()
    frmProcesses.Show 1
End Sub


Private Sub mnuReqForm_Click()
    'frmRequisition.Show
    frmGMPs.Show
End Sub

Private Sub mnuProdPlan_Click()
    frmProdPlan.Show
End Sub

Private Sub mnuProdPlanrpt_Click()

    On Error GoTo err
    
    Dim f As New frmPrint
    Load f
    With f
        Set .rpt = New rptProdPlan
        .cmbCust.ClearVals
        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
        .cmbCust.AddItem "<All Customers>", "0"
        .cmbCust.AddVals con, "CustCode", "ForeignCustomers", "CustCode"
        .SelFor = "{VOrderPlanning.OrderNo}='"
        .JobCondField = " CustCode "
        .JobField = " Distinct OrderNo "
        .JobIDField = " OrderNo "
        .JobTable = " VFItemsQtyToPlan "
        .TheEnd = "'"
        .Show , MainForm
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPymtTerms_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("PaymentTerms", "Payment Terms", "Payment Terms")
End Sub

Private Sub mnuRcvRepair_Click()
    frmVendRcvRepair.Show
End Sub

Private Sub mnuRcvRepairPoli_Click()
    frmVendRcvRepair.Show
End Sub

Private Sub mnuRcvRepairWastage_Click()
    frmVendRcvRepairWaste.Show
End Sub

Private Sub mnuRcvRpts_Click()
    Dim f As New frmRptIssRcv
    Call f.ShowForIss(True)
End Sub

Private Sub mnuRecRM_Click()
    frmRMRcvItems.Show
End Sub

Private Sub mnuRecSDril_Click()
    Call ShowRcvSForm(3)
End Sub

Private Sub mnuRecSFill_Click()
    Call ShowRcvSForm(4)
End Sub

Private Sub mnuRecSForg_Click()
    Call ShowRcvSForm(1)
End Sub

Private Sub mnuRecSFunc_Click()
    Call ShowRcvSForm(5)
End Sub

Private Sub mnuRecSFunt_Click()
    Call ShowRcvSForm(9)
End Sub

Private Sub mnuRecSHeat_Click()
    Call ShowRcvSForm(6)
End Sub

Private Sub mnuRecSHold_Click()
    Call ShowRcvSForm(13)
End Sub

Private Sub mnuRecSMach_Click()
    Call ShowRcvSForm(2)
End Sub

Private Sub mnuRecSPack_Click()
    Call ShowRcvSForm(12)
End Sub

Private Sub mnuRecSPass_Click()
    Call ShowRcvSForm(8)
End Sub

Private Sub mnuRecSPoli_Click()
    Call ShowRcvSForm(7)
End Sub

Private Sub mnuRecSPoRe_Click()
    Call ShowRcvSForm(10)
End Sub

Private Sub mnuRecSUltr_Click()
    Call ShowRcvSForm(11)
End Sub

Private Sub mnuRecVDril_Click()
    Call ShowRcvVForm(3)
End Sub

Private Sub mnuRecVFill_Click()
    Call ShowRcvVForm(4)
End Sub

Private Sub mnuRecVForg_Click()
    Call ShowRcvVForm(1)
End Sub

Private Sub mnuRecVFunc_Click()
    Call ShowRcvVForm(5)
End Sub

Private Sub mnuRecVFunt_Click()
    Call ShowRcvVForm(9)
End Sub

Private Sub mnuRecVHeat_Click()
    Call ShowRcvVForm(6)
End Sub

Private Sub mnuRecVHold_Click()
    Call ShowRcvVForm(13)
End Sub

Private Sub mnuRecVMach_Click()
    Call ShowRcvVForm(2)
End Sub

Private Sub mnuRecVPack_Click()
    Call ShowRcvVForm(12)
End Sub

Private Sub mnuRecVPass_Click()
    Call ShowRcvVForm(8)
End Sub

Private Sub mnuRecVPoli_Click()
    Call ShowRcvVForm(7)
End Sub

Private Sub mnuRecVPoRe_Click()
    Call ShowRcvVForm(10)
End Sub

Private Sub mnuRecVUltr_Click()
    Call ShowRcvVForm(11)
End Sub

Private Sub mnuRecWastDril_Click()
    Call ShowWastRcvForm(3)
End Sub

Private Sub mnuRecWastFill_Click()
    Call ShowWastRcvForm(4)
End Sub

Private Sub mnuRecWastForg_Click()
    Call ShowWastRcvForm(1)
End Sub

Private Sub mnuRecWastFunc_Click()
    Call ShowWastRcvForm(5)
End Sub

Private Sub mnuRecWastFunt_Click()
    Call ShowWastRcvForm(9)
End Sub

Private Sub mnuRecWastHeat_Click()
    Call ShowWastRcvForm(6)
End Sub

Private Sub mnuRecWastHold_Click()
    Call ShowWastRcvForm(13)
End Sub

Private Sub mnuRecWastMach_Click()
    Call ShowWastRcvForm(2)
End Sub

Private Sub mnuRecWastPack_Click()
    Call ShowWastRcvForm(12)
End Sub

Private Sub mnuRecWastPass_Click()
    Call ShowWastRcvForm(8)
End Sub

Private Sub mnuRecWastPoli_Click()
    Call ShowWastRcvForm(7)
End Sub

Private Sub mnuRecWastPoRe_Click()
    Call ShowWastRcvForm(10)
End Sub

Private Sub mnuRecWastUltr_Click()
    Call ShowWastRcvForm(11)
End Sub

Private Sub mnuReq_Click()
    'frmRequisition.Show
    frmGMPs.Show
End Sub

Private Sub mnuReqRpt_Click()

    On Error GoTo err
    
    Dim f As New frmPrint
    Load f
    With f
        Set .rpt = New rptRequsition
        .cmbCust.ClearVals
        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
        .cmbCust.AddItem "<All Processes>", "0"
        .cmbCust.AddVals con, "Description", "Processes", "ProcessID"
        .SelFor = "{VRequsitionrpt.RecieptID}='"
        .JobCondField = " ProcessID "
        .JobField = " Distinct RecieptID "
        .JobIDField = " RecieptID "
        .JobTable = " VRequsitionrpt "
        .TheEnd = "'"
        .Show , MainForm
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuRM_Click()
    frmRM.Show
End Sub

Private Sub mnuRMAssItems_Click()
    frmRMVendAssItems.Show
End Sub

Private Sub mnuRMGroups_Click()
    Dim f As New frmItemGroups
    Load f
    Call f.ShowForm("RMGroups", "RM Groups", "RM Name")
End Sub

Private Sub mnuRMListrpt_Click()
    Call ShowRMListrpt
End Sub

Private Sub mnuRMPO_Click()
    frmRMPurchOrd.Show
End Sub

Private Sub mnuRMVendList_Click()
    frmRMVenderList.Show
End Sub

Private Sub mnuRMVendListrpt_Click()
    Call ShowRMVendListrpt
End Sub

Private Sub mnuShowDept_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.ShowInactive(mnuShowDept)
    Else
        Call ContDeptList.ShowInactive(mnuShowDept)
    End If
End Sub

Private Sub mnuShowInactive_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.ShowInactive(mnuShowInactive)
    Else
        Call ContractorList.ShowInactive(mnuShowInactive)
    End If
End Sub

Private Sub mnuSteel_Click()
    frmSteelList.Show
End Sub

Private Sub mnuSteelVends_Click()
    frmSteelVends.Show
End Sub

Private Sub mnuTaxRanges_Click()
    Load TaxRanges
    'Holidays.Move 3000, 500
    With TaxRanges
        .Move (Me.Width - .Width) / 2, Me.Top + 500
        .Show
    End With
End Sub

Private Sub mnuUnits_Click()
    frmUnits.Show
End Sub

Private Sub mnuUserManager_Click()
    frmUsers.Show 1
End Sub

Public Sub ApplySecuritySettings(UserID As Long)
   
    On Error GoTo err
    
    'General Settings...
    mnuLotProc.Enabled = UserHasAccess("SetInitLotProc", UserID)
    mnuPer.Enabled = UserHasAccess("PercentRanges")
    mnuComp.Enabled = UserHasAccess("Components")
    mnuProcess.Enabled = UserHasAccess("Processes")
    mnuFDAProcess.Enabled = UserHasAccess("FDAProcesses")
    mnuSteelVends.Enabled = UserHasAccess("SteelVenders")
    mnuLabs.Enabled = UserHasAccess("TestLabs")
    mnuUnits.Enabled = UserHasAccess("MeasuringUnits")
    mnuCompositions.Enabled = UserHasAccess("Compositions")
    mnuPorts.Enabled = UserHasAccess("CitiesNPorts")
    mnuPymtTerms.Enabled = UserHasAccess("PymtTerms")
    
    'Venders
    mnuVendList.Enabled = UserHasAccess("VenderList")
    mnuVendAssItems.Enabled = UserHasAccess("VendAssItems")
    mnuVendPymt.Enabled = UserHasAccess("VendAdvPymt")
    mnuVendPaySheet.Enabled = UserHasAccess("VendPymtSheet")
    mnuVendAdvShort.Enabled = UserHasAccess("NewShortTerm")
    mnuVendAdvLong.Enabled = UserHasAccess("NewLongTerm")
    mnuVendShortLoanLedger.Enabled = UserHasAccess("ShortTermLedger")
    mnuVendLongLoanLedger.Enabled = UserHasAccess("LongTermLedger")
    mnuVendLedger.Enabled = UserHasAccess("VendLedger")
    
    'Raw Material
    mnuRMGroups.Enabled = UserHasAccess("RMGroups")
    mnuRM.Enabled = UserHasAccess("RawMaterial")
    mnuRMVendList.Enabled = UserHasAccess("RMVendList")
    mnuRMAssItems.Enabled = UserHasAccess("RMVAssItems")
    mnuRMPO.Enabled = UserHasAccess("RMVNewPO")
    mnuPOrderList.Enabled = UserHasAccess("RMVPOList")
    mnuRecRM.Enabled = UserHasAccess("RMRcvRM")
    mnuIssRM.Enabled = UserHasAccess("RMIssRM")
    
    'HICO Resources
    mnuHRGroups.Enabled = UserHasAccess("HICOResGroups")
    mnuHICOResources.Enabled = UserHasAccess("HICOResources")
    mnuVendAssR.Enabled = UserHasAccess("AssHICORes")
        
    'Company
    mnuCompDeatil.Enabled = UserHasAccess("CompanyDetail")
    mnuCompCatal.Enabled = UserHasAccess("CompanyCatalog")
    mnuItmGrps.Enabled = UserHasAccess("ItemGroups")
    mnuItemsFinQuality.Enabled = UserHasAccess("ItemFinQuality")
    mnuItemTypes.Enabled = UserHasAccess("ItemTypes")
    mnuForeignCatal.Enabled = UserHasAccess("ForeignCatalog")
    mnuCustList.Enabled = UserHasAccess("CustomerList")
    mnuCustCatal.Enabled = UserHasAccess("CustomerCatalog")
    mnuNewOrd.Enabled = UserHasAccess("CustNewOrder")
    mnuNewStockOrder.Enabled = UserHasAccess("StockNewOrder")
    mnuCustOrdList.Enabled = UserHasAccess("CustOrdList")
    mnuOrdStatus1.Enabled = UserHasAccess("CustOrdStatus")
    mnuNewProforma.Enabled = UserHasAccess("CustNewProforma")
    mnuCustProformaList.Enabled = UserHasAccess("CustProList")
        
    mnuReq.Enabled = UserHasAccess("Requsition")
    mnuProdPlan.Enabled = UserHasAccess("AddProdPlan")
    mnuAuth.Enabled = UserHasAccess("Authorize")
    
    'Reports
    
    'Users
    mnuUserManager.Enabled = UserHasAccess("UserManagement")
    mnuPassword.Enabled = UserHasAccess("ChangePassword")
    
    
    HicoVisible = UserHasAccess("HicoVisible")
    'If UserID = 1 Then mnuAuth.Enabled = True Else mnuAuth.Enabled = False
    
    If UserID <> 1 Then Call ApplySecurityForIssRcv(UserID)
    '
    'With Toolbar1.Buttons
    '   '.Item("TakeAtt").Enabled = mnutake.Enabled
    '   '.Item("NewShort").Enabled = mnuShort.Enabled
    '   '.Item("NewLong").Enabled = mnuLong.Enabled
    '   '.Item("EmpLedger").Enabled = mnuEmpLedger.Enabled
    '   '.Item("SalSheet").Enabled = mnuSalSheet.Enabled
    'End With
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub ApplySecurityForIssRcv(UserID As Long)
    
    For i = 1 To mnuIssRcvProcs.Count
        mnuIssRcvProcs(i).Enabled = False
    Next i
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select ProcessID From UserAssProcesses Where UserID=" & UserID, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            mnuIssRcvProcs(![ProcessID]).Enabled = True
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub
Public Sub mnuVendAdvLong_Click()

    With frmVendNewAdv
        .LongTerm = True
        .Show 1
    End With
    
End Sub

Public Sub mnuVendAdvShort_Click()
    
    With frmVendNewAdv
        .LongTerm = False
        .Show 1
    End With
    
End Sub

Private Sub mnuVendAssItems_Click()
    frmVendAssItems.Show
End Sub

Private Sub mnuVendAssR_Click()
    frmVendAssHR1.Show
End Sub

Private Sub mnuVendIssItems_Click()
    frmVendIssItems.Show
End Sub

Private Sub mnuVendIssItemsByItems_Click()
    frmVendIssItemsByItem.Show
End Sub

Private Sub mnuVendLedger_Click()
    frmVendLedger.Show
End Sub

Private Sub mnuVendList_Click()
    frmVenderList.Show
End Sub

Private Sub mnuVendListrpt_Click()
    Call ShowVendListrpt
End Sub

Private Sub mnuVendLongLoanLedger_Click()
    frmVendLoanLedger.Show
End Sub

Private Sub mnuVendPaySheet_Click()
    frmVendSS.Show    'frmVendPaySheet.Show
End Sub

Private Sub mnuVendPymt_Click()
    frmVendAdvSalary.Show
End Sub

Private Sub mnuVendRcvItems_Click()
    frmVendRcvItems.Show
End Sub

Private Sub mnuVendRcvItemsFromSupervisor_Click()
    frmVendRcvItems1.Show
End Sub

Private Sub mnuVendRcvWst_Click()
    frmVendRcvWastage.Show
End Sub

Private Sub mnuVendShortLoanLedger_Click()
    frmVendShortAdvLedger.Show
End Sub

'-----------------------------------------------------------
'Code Of Pop Menu Ports.
Private Sub NewCity_Click()
    Call frmPorts.NewCity_Click
End Sub

Private Sub NewEntry_Click()
    Call frmPorts.NewEntry_Click
End Sub

Private Sub Rename_Click()
    Call frmPorts.Rename_Click
End Sub

Private Sub del_Click()
    Call frmPorts.del_Click
End Sub

'-----------------------------------------------------------
Private Sub Timer1_Timer()

    Timer1.Interval = 1000
    Dim SeverTime1 As Date
    SeverTime1 = con.Execute("Select GetDate()").Fields(0)
    ServerTime = Format(SeverTime1, "hh:mm:ss AMPM")
    MainForm.Caption = "Production Package." & Space$(5) & " SERVER TIME : " & Format(SeverTime1, "dddd dd-MMM-yyyy hh:mm:ss AMPM")
    
    If Format(ServerDate, "dd-MMM-yyyy") <> Format(SeverTime1, "dd-MMM-yyyy") Then ServerDate = Format(SeverTime1, "dd-MMM-yyyy")
    
End Sub

