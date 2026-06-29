Attribute VB_Name = "Mod_ExpBar"
Public Sub CreateExplorerBarForExport()
    
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite '&H8000000F
        .BackColorEnd = vbWhite '&H8000000F
        
'        .Bars.add , "Company", "Company"
'
'        .Bars.add , "Customers", "Customers"
        
        .Bars.add , "Orders", "Orders"
        
        
            .Bars.add , "ProformaInvoice", "Proforma Invoice"

            .Bars.add , "CustomDocuments", "Custom Documents"

            .Bars.add , "ComDocuments", "Comm. Documents"


            '.Bars.Add , "cGMP", "cGMP"

            If strCompany <> "Instrumed" Then
                .Bars.add , "BankDocuments", "Bank Documents"
            End If

            .Bars.add , "PaymentStatus", "Payment Status"

            .Bars.add , "Rebate", "Rebate Documents"

            .Bars.add , "Satistics", "Statistics"

            .Bars.add , "EEStatement", "EE Statement"
            
            If strCompany = "IAA" Or strCompany = "Kami" Then
                .Bars.add , "Others", "Others"
            End If
            '        .Bars.add , "UserManagement", "User Management"
        
        
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
            'Call MainForm.ExpBar_BarClick(.Bars(i))
        Next i
        
    '        .Bars.add(, "Quit", "Quit").CanExpand = False
    '        .Bars("Quit").Items.add , "Quit", "Click Here or Dbl Click On Main Form To Quit"
        
        .Visible = True
        
    End With
    
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub

Public Sub CreateExplorerBarForCompany()
    
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        
        .Bars.add , "Company", "Company"
        .Bars.add , "Customers", "Customers"
        
        
'        .Bars.add , "Others", "Others"
        
        
'        If strCompany = "Weldon Instruments" Or strCompany = "PAKSMITH GROUP" Or strCompany = "Sunlike" Or strCompany = "QEL" Or strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Instrumed" Or strCompany = "Towne" Then
'            .Bars.add , "Agents", "Agents"
'            '.Bars.add , "AccHeads", "Account Heads"
'            '.Bars.add , "AccMakers", "Makers"
'            '.Bars.add , "AccBanks", "Banks"
'            '.Bars.add , "PrlEmployees", "Employees"
'        ElseIf strCompany = "Weldon Industries" Or strCompany = "Tecno" Then
'            '.Bars.add , "AccMakers", "Makers / Venders"
'        End If
        
        .Bars.add , "UserManagement", "User Management"
        
'       .Bars.add , "UserManagement", "User Management"
                
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
            'Call MainForm.ExpBar_BarClick(.Bars(i))
        Next i
        
'        .Bars.add(, "Quit", "Quit").CanExpand = False
'        .Bars("Quit").Items.add , "Quit", "Click Here or Dbl Click On Main Form To Quit"
        
        .Visible = True
        
    End With
    
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub


Public Sub CreateExplorerBarForStar()
        
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        
        .Bars.add , "Company", "Company Information"
        
        .Bars.add , "BusPart", "Bussiness Partners"
        
        .Bars.add , "RM", "Raw Material"
        
        .Bars.add , "Production", "Production Management"
        
        .Bars.add , "Inventory", "Stock Status"
        
        .Bars.add , "Billing", "Billing Info."
        
        .Bars.add(, "UserManagement", "Admin Console").CanExpand = IIf(UserID = 1, True, False)
                
        '.Bars.Add , "RawMaterial", "Raw Material"
                
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
        Next i
        
        .Bars.add(, "Quit", "Quit").CanExpand = False
        .Bars("Quit").Items.add , "Quit", "Click Here or Dbl Click On Main Form To Quit"
        
        .Visible = True
        
    End With
    
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub


Public Sub CreateExplorerBarForAccounts()
        
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        
        .Bars.add , "AccHeads", "Account Heads"
        If strCompany = "Dr-Frgz" Then
            .Bars.add , "AccFinancialBudgeting", "Financial Budgeting"
        End If
        .Bars.add , "AccBanks", "Banks"
         
        If strCompany = "Instrumed" Then
            .Bars.add , "AccMakers", "Makers/Venders"
        End If
        
        .Bars.add , "AccVouchers", "Vouchers"
        
        'If Not (strCompany = "Banzai" Or strCompany = "Instrumed" Or strCompany = "Scard") Then
        .Bars.add , "AccEmployeeAdvances", "Employee Advances"
        .Bars.add , "AccMakerAdvances", "Advances"
        'End If
             
        
        .Bars.add , "AccReports", "Reports"
        
        If strCompany = "Scard" Then
            .Bars.add , "UserManagement", "User Management"
        End If
        '.Bars.add , "AccAuthorizations", "Authorizations"
'        .Bars.add , "UserManagement", "User Management"
    
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
'            Call MainForm.ExpBar_BarClick(.Bars(i))
        Next i
        
'        .Bars.add(, "Quit", "Quit").CanExpand = False
'        .Bars("Quit").Items.add , "Quit", "Click Here or Dbl Click On Main Form To Quit"
        
        .Visible = True
        
    End With
    
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub

Public Sub CreateExplorerBarForPayroll()
    
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        
        
'        .Bars.add , "PrlCVs", "CVs Detail"
        
        
        .Bars.add , "PrlEmployees", "Employees"
        
        .Bars.add , "PrlAttendance", "Attendance"

        .Bars.add , "PrlAdvances", "Advances"

        .Bars.add , "PrlEmployeePayments", "Employee Payments"
'
''        .Bars.add , "PrlContractors", "Contractors"
'
        .Bars.add , "PrlReports", "Reports"

        .Bars.add , "PrlSettings", "Settings"
                
'        .Bars.add , "PrlUserManagement", "User Management"
                
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
            'Call MainForm.ExpBar_BarClick(.Bars(i))
        Next i
        
'        .Bars.add(, "Quit", "Quit").CanExpand = False
'        .Bars("Quit").Items.add , "Quit", "Click Here or Dbl Click On Main Form To Quit"
        
        .Visible = True
        
    End With
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub

Public Sub CreateExplorerBarForStock()
    
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        
'        .Bars.add , "Company", "Company"
'        .Bars.add , "Customers", "Customers"
        
        .Bars.add , "StkRawMaterial", "Raw Material"
        '.Bars.add , "StkRawMaterialSampling", "Sampling"
        
'        If strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
            .Bars.add , "StkSemiFinishStock", "S/F Stock"
'        End If
'
'
        .Bars.add , "StkStock", "Finish Stock"
'
'
'        If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then
'            .Bars.add , "PrdStore", "Gatepass/Receiving"
'        End If
'
'        If strCompany = "Dr-Frgz" Then
'            .Bars.add , "StkGateEntries", "Gate Entries"
'        End If

'        .Bars.add , "UserManagement", "User Management"
'        .Bars.add , "StkRejectionStore", "Rejection Store"
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
            'Call MainForm.ExpBar_BarClick(.Bars(i))
        Next i
        
'        .Bars.add(, "Quit", "Quit").CanExpand = False
'        .Bars("Quit").Items.add , "Quit", "Click Here or Dbl Click On Main Form To Quit"
        
        .Visible = True
        
    End With
    
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub

Public Sub CreateExplorerBarForProduction()
    'CreateExplorerBarForQA
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        

        If iProductionType = 2 Then
            .Bars.add , "AccMakers", "Makers"
            .Bars.add , "PrdProduction", "Production"
            .Bars.add , "PrdManagementReports", "Management Reports"
        End If
        If iProductionType <> 2 Then
            .Bars.add , "AccMakers", "Makers"
            .Bars.add , "PrdPurchaseANDProcessing", "Purchasing && Processing"
            .Bars.add , "PrdProduction", "Production"
        End If

        .Bars.add , "PrdStatistics", "Statistics"

        .Bars.add , "PrdGeneralLists", "General Lists"
        

                
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
        Next i
        

        
        .Visible = True
         
    End With
    
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub

Public Sub CreateExplorerBarForQMS()
    'CreateExplorerBarForSampling
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        
        If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then
            .Bars.add , "QMSQA", "Q.A"
        Else
            .Bars.add , "QMSQMS", "QMS"
        End If
        .Bars.add , "QMSDocuments", "Documents"
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
        
        Next i
        
        .Visible = True
         
    End With
    
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub

Public Sub CreateExplorerBarForImport()
    'CreateExplorerBarForSampling
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        
        
        .Bars.add , "ImpImport", "Import"
       
      
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
        
        Next i
        
        .Visible = True
         
    End With
    
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub


Public Sub CreateExplorerBarForFixedAssets()
    '
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        
        
        .Bars.add , "FAFixedAssets", "Fixed Assets"
        
                
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
        
        Next i
        
        .Visible = True
         
    End With
    
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub

Public Sub CreateExplorerBarForDashBoard()
    '
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        
        
        .Bars.add , "DasDashBoard", "Dash Board"
        
'        If strCompany = "IAA" OR strCompany = "Kami" Then
'            .Bars.add , "DasStatistics", "Statistics"
'        End If
        
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
        
        Next i
        
        .Visible = True
         
    End With
    
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub

Public Sub CreateExplorerBarForHelp()
    '
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        
        
        .Bars.add , "HlpHelp", "Help"
        
                
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
        Next i
        
        .Visible = True
         
    End With
    
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub


Public Sub SetUserOptions(ModuleName As String)
    
End Sub

Public Sub AddAllOptions()
    
    If CurrentUserName <> "Administrator" Then
        Exit Sub
    End If
    
    If MsgBox("Are you sure to Reset all the MenuOptions ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    con.Execute "Delete From MenuOptions"
    Call CreateExplorerBarForAccounts
    Call AddModuleOption("Accounts", "MainForm")
     
    If strCompany <> "Scard" Then
    
        Call CreateExplorerBarForCompany
        Call AddModuleOption("Company", "MainForm")
        
        Call CreateExplorerBarForExport
        Call AddModuleOption("Export", "MainForm")
        
        Call CreateExplorerBarForSampling
        Call AddModuleOption("Sampling", "MainForm")
        
        Call CreateExplorerBarForPayroll
        Call AddModuleOption("Payroll", "MainForm")
        
        Call CreateExplorerBarForProduction
        Call AddModuleOption("Production", "MainForm")
        
    '    Call CreateExplorerBarForStar
    '    Call AddModuleOption("Star", "MainForm")
        
        Call CreateExplorerBarForStock
        Call AddModuleOption("Stock", "MainForm")
        
        Call CreateExplorerBarForDashBoard
        Call AddModuleOption("DashBoard", "MainForm")
        
        Call CreateExplorerBarForHelp
        Call AddModuleOption("Help", "MainForm")
        
        Call CreateExplorerBarForFixedAssets
        Call AddModuleOption("FixedAssets", "MainForm")
        
        Call CreateExplorerBarForQMS
        Call AddModuleOption("QMS", "MainForm")
    End If
    
    con.Execute "Delete From UserMenuOptions Where UserID=1"
    con.Execute "Insert into UserMenuOptions(UserID,OptionID) Select 1 as userID,OptionID from MenuOptions"
    
End Sub


Public Sub AddModuleOption(ModuleName As String, FormName As String)

    Dim i As Long, j As Long

    With MainForm.ExpBar
        For i = 1 To .Bars.count
            With .Bars(i)
            
            Call MainForm.ExpBar_BarClick(MainForm.ExpBar.Bars(i))
            
            For j = 1 To .Items.count
                On Error Resume Next        ''If an option is present multiple places, like S/F issuance being available in Stock & in Production.
                con.Execute "INSERT INTO MenuOptions(OptionID,OptionName,ModuleName,FormName) VALUES ('" & _
                .Items(j).key & "','" & .Items(j).Text & "','" & ModuleName & "','" & FormName & "')"
            Next
            
            End With
        Next
    End With
    
    
Debug.Print i, j
End Sub


Public Sub CreateExplorerBarForQMS_Bad()
    '
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        
        
        .Bars.add , "QMSQMS", "QMS"
        
                
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
        
        Next i
        
        .Visible = True
         
    End With
    
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub


Public Sub CreateExplorerBarForFavourites()
     
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBarFavourites
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
         
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        
        .Bars.add , "Favourites", "Favourites"
         
        For i = 1 To .Bars.count
            .Bars(i).State = eBarExpanded
            .Bars(i).CanExpand = False
        Next i
        .Visible = True
         
    End With
     
    Dim myBar As vbalExplorerBarLib6.cExplorerBar
   
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT * FROM VUser_FavouriteLinks WHERE UserID=" & UserID, con, adOpenForwardOnly, adLockReadOnly
        MainForm.ExpBarFavourites.Bars(1).Items.Clear
        Set myBar = MainForm.ExpBarFavourites.Bars(1)
        Do Until .EOF
            Call myBar.Items.add(, !OptionID & "", !OptionName & "")
            .MoveNEXT
        Loop
        .Close
    End With
    
End Sub

Public Sub CreateExplorerBarForSampling()
    'CreateExplorerBarForSampling
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eDefaultStyle
        
        .BackColorStart = vbWhite
        .BackColorEnd = vbWhite
        
        .Bars.add , "SAMSampling", "Sampling"
       
        
        For i = 1 To .Bars.count
            .Bars(i).State = eBarCollapsed
        
        Next i
        
        .Visible = True
         
    End With
    
    MainForm.Picture1.Visible = True
    MainForm.Picture1.Refresh
    
End Sub

