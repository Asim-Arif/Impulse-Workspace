Attribute VB_Name = "Mod_ExpBar_Vision"
Public Sub CreateExplorerBarForExport()
    
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eSearchStyle
        
        .BackColorStart = &H8000000F
        .BackColorEnd = &H8000000F
        
'        .Bars.add , "Company", "Company"
'
'        .Bars.add , "Customers", "Customers"
        
        .Bars.add , "Orders", "Orders"
        
        .Bars.add , "ProformaInvoice", "Proforma Invoice"
        
        .Bars.add , "CustomDocuments", "Custom Documents"
                
        .Bars.add , "ComDocuments", "Comm. Documents"
        
        
        '.Bars.Add , "cGMP", "cGMP"
        
        
        .Bars.add , "BankDocuments", "Bank Documents"
                
        .Bars.add , "PaymentStatus", "Payment Status"
        
        .Bars.add , "Rebate", "Rebate Documents"
        
        .Bars.add , "Satistics", "Statistics"
        
        .Bars.add , "EEStatement", "EE Statement"
        
                
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
        
        .Style = eSearchStyle
        
        .BackColorStart = &H8000000F
        .BackColorEnd = &H8000000F
        
        .Bars.add , "Company", "Company"
        .Bars.add , "Customers", "Customers"
         
        If strCompany = "Weldon Instruments" Or strCompany = "PAKSMITH GROUP" Or strCompany = "Sunlike" Or strCompany = "QEL" Then
            .Bars.add , "Agents", "Agents"
            .Bars.add , "AccHeads", "Account Heads"
            .Bars.add , "AccMakers", "Makers"
            .Bars.add , "AccBanks", "Banks"
            .Bars.add , "PrlEmployees", "Employees"
        ElseIf strCompany = "Weldon Industries" Then
            .Bars.add , "AccMakers", "Makers"
        End If
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
        
        .Style = eSearchStyle
        
        .BackColorStart = &H8000000F
        .BackColorEnd = &H8000000F
        
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
        
        .Style = eSearchStyle
        
        .BackColorStart = &H8000000F
        .BackColorEnd = &H8000000F
        
        
        .Bars.add , "AccVouchers", "Vouchers"
        .Bars.add , "AccEmployeeAdvances", "Employee Advances"
        .Bars.add , "AccReports", "Reports"
        '.Bars.add , "AccAuthorizations", "Authorizations"
'        .Bars.add , "UserManagement", "User Management"
    
        For i = 1 To .Bars.count
            '.Bars(i).State = eBarCollapsed
            Call MainForm.ExpBar_BarClick(.Bars(i))
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
        
        .Style = eSearchStyle
        
        .BackColorStart = &H8000000F
        .BackColorEnd = &H8000000F
        
        '.Bars.add , "PrlCVs", "CVs Detail"
        
        .Bars.add , "PrlAttendance", "Attendance"
        
        .Bars.add , "PrlAdvances", "Advances"
        
        .Bars.add , "PrlEmployeePayments", "Employee Payments"
        
'        .Bars.add , "PrlContractors", "Contractors"
        
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
        
        .Style = eSearchStyle
        
        .BackColorStart = &H8000000F
        .BackColorEnd = &H8000000F
        
'        .Bars.add , "Company", "Company"
'        .Bars.add , "Customers", "Customers"
        .Bars.add , "StkRawMaterial", "Raw Material"
        
        .Bars.add , "StkStock", "Finish Stock"
          
          
         
                
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

Public Sub CreateExplorerBarForProduction()
    
    
    Dim iBar As cExplorerBar
    Dim iItem As cExplorerBarItem
    
    With MainForm.ExpBar
        .Bars.Clear
        .Visible = False
        .UseExplorerStyle = False
        
        .Style = eSearchStyle
        
        .BackColorStart = &H8000000F
        .BackColorEnd = &H8000000F
        
'        .Bars.add , "Company", "Company"
'
'        .Bars.add , "Customers", "Customers"
        .Bars.add , "PrdProduction", "Production"
        .Bars.add , "PrdGeneralLists", "General Lists"
        
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
    
    Call CreateExplorerBarForCompany
    Call AddModuleOption("Company", "MainForm")
    
    Call CreateExplorerBarForExport
    Call AddModuleOption("Export", "MainForm")
    
    Call CreateExplorerBarForPayroll
    Call AddModuleOption("Payroll", "MainForm")
    
    Call CreateExplorerBarForProduction
    Call AddModuleOption("Production", "MainForm")
    
'    Call CreateExplorerBarForStar
'    Call AddModuleOption("Star", "MainForm")
    
    Call CreateExplorerBarForStock
    Call AddModuleOption("Stock", "MainForm")
    
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
                con.Execute "INSERT INTO MenuOptions(OptionID,OptionName,ModuleName,FormName) Values ('" & _
                .Items(j).key & "','" & .Items(j).Text & "','" & ModuleName & "','" & FormName & "')"
            Next
            
            End With
        Next
    End With
    
    
Debug.Print i, j
End Sub
