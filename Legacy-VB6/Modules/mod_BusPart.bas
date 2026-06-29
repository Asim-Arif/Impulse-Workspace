Attribute VB_Name = "mod_BusPart"

Public Sub AddNewBusPart(BusPartDesc As String)
    
    Dim f As Form
    Select Case BusPartDesc
        Case "L Customer"
        Case "Importer"
        Case "Employee"
        Case "R.M. Vender"
            
        Case "Steel Vender"
            Load frmNewSteelVend
            frmNewSteelVend.Show 1
        Case "Supplier"
        Case "provider"
        Case "Test Lab"
            Load frmNewLab
            frmNewLab.Show 1
        Case "Bank"
        Case "S. Agency"
        Case "C. Agency"
    End Select
    
    
End Sub
Public Sub ShowBusPartLedger(BusPartDesc As String)
    
    Dim f As Form
    Select Case BusPartDesc
        
        Case "Contractor"
            frmVendLedger.Show
            
        Case "R.M. Vender"
            
        Case "Steel Vender"
            Load frmNewSteelVend
            frmNewSteelVend.Show 1
        Case "Supplier"
        Case "provider"
        Case "Test Lab"
            Load frmNewLab
            frmNewLab.Show 1
        Case "Bank"
        Case "S. Agency"
        Case "C. Agency"
    End Select
    
    
End Sub
Public Sub ShowBusPartAssItems(BusPartDesc As String)
    
    Dim f As Form
    Select Case BusPartDesc
        
        Case "Contractor"
            frmVendAssItems.Show
            
        Case "R.M. Vender"
            frmRMVendAssItems.Show
        Case "Steel Vender"
            Load frmNewSteelVend
            frmNewSteelVend.Show 1
        Case "Supplier"
        Case "provider"
        Case "Test Lab"
            Load frmNewLab
            frmNewLab.Show 1
        Case "Bank"
        Case "S. Agency"
        Case "C. Agency"
    End Select
    
    
End Sub
Public Sub ShowBusPartList(BusPartDesc As String)

    Select Case BusPartDesc
        Case "L Customer"
        Case "Importer"
        Case "Employee"
        Case "Contractor"
            frmVenderList.Show
        Case "R.M. Vender"
            frmRMVenderList.Show
        Case "Steel Vender"
            frmSteelVends.Show
        Case "Supplier"
        Case "provider"
        Case "Test Lab"
            frmTestLabs.Show
        Case "Bank"
        Case "S. Agency"
        Case "C. Agency"
    End Select
    
End Sub
