-- =========================================================================
-- Standard Screen & Menu Options for Budgeting, Fixed Assets, and Gate Management
-- =========================================================================

-- 1. BUDGETING
IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'BudHub')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('BudHub', 'Budgeting Hub', 'Budgeting', 'BudgetingHub');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'BudDashboard')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('BudDashboard', 'Budgeting Analytics & Performance', 'Budgeting', 'BudgetingDashboard');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'BudExpenseEntry')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('BudExpenseEntry', 'Expense Budget Entry', 'Budgeting', 'BudgetExpenseEntry');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'BudIncomeEntry')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('BudIncomeEntry', 'Income Budget Entry', 'Budgeting', 'BudgetIncomeEntry');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'BudIncomeRevisions')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('BudIncomeRevisions', 'Income Budget Revisions', 'Budgeting', 'BudgetIncomeRevision');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'BudReports')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('BudReports', 'Budget Reports & Variance', 'Budgeting', 'BudgetReports');

-- 2. FIXED ASSETS
IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'FAHub')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('FAHub', 'Fixed Assets Hub', 'FixedAssets', 'FixedAssetsHub');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'FAList')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('FAList', 'Fixed Assets List', 'FixedAssets', 'FixedAssetsList');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'FANewAsset')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('FANewAsset', 'New Asset Profile / Entry', 'FixedAssets', 'FixedAssetProfile');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'FARegister')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('FARegister', 'Depreciation Register', 'FixedAssets', 'FixedAssetsRegister');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'FAAdjustment')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('FAAdjustment', 'Fixed Asset Adjustment', 'FixedAssets', 'FixedAssetAdjustment');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'FAMove')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('FAMove', 'Asset Relocation & Move', 'FixedAssets', 'FixedAssetMove');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'FADisposal')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('FADisposal', 'Asset Disposal', 'FixedAssets', 'FixedAssetDisposal');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'FADashboard')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('FADashboard', 'Fixed Assets Dashboard', 'FixedAssets', 'FixedAssetsDashboard');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'FAReports')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('FAReports', 'Fixed Assets Reports', 'FixedAssets', 'FixedAssetsReports');

-- 3. GATE MANAGEMENT
IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateHub')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateHub', 'Gate Management Hub', 'GateManagement', 'GateManagementHub');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateDashboard')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateDashboard', 'Gate Operations Dashboard', 'GateManagement', 'GateDashboard');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateInwardList')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateInwardList', 'Gate Inward Passes List', 'GateManagement', 'GateInList');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateInwardEntry')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateInwardEntry', 'Gate Inward Pass Entry', 'GateManagement', 'GateInEntry');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateInwardReceive')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateInwardReceive', 'Returnable Receiving', 'GateManagement', 'GateInReceive');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateMaterialInward')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateMaterialInward', 'Material Inward Register', 'GateManagement', 'MaterialInward');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateMaterialInwardEntry')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateMaterialInwardEntry', 'Material Inward Entry', 'GateManagement', 'MaterialInwardEntry');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateMaterialOutward')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateMaterialOutward', 'Material Outward Register', 'GateManagement', 'MaterialOutward');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateMaterialOutwardEntry')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateMaterialOutwardEntry', 'Material Outward Entry', 'GateManagement', 'MaterialOutwardEntry');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateVisitors')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateVisitors', 'Visitor Management Register', 'GateManagement', 'VisitorsList');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateVisitorEntry')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateVisitorEntry', 'Visitor Pass Entry', 'GateManagement', 'VisitorEntry');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateVisitorCheckout')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateVisitorCheckout', 'Visitor Gate Out Checkout', 'GateManagement', 'VisitorGateOut');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateVehicles')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateVehicles', 'Vehicle Movement Log', 'GateManagement', 'VehiclesList');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateVehicleEntry')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateVehicleEntry', 'Vehicle Entry / Exit', 'GateManagement', 'VehicleEntry');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GatePassTypes')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GatePassTypes', 'Gate Pass Types Setup', 'GateManagement', 'GatePassTypes');

IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'GateReports')
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName) VALUES ('GateReports', 'Gate Security Reports', 'GateManagement', 'GateReports');
GO
