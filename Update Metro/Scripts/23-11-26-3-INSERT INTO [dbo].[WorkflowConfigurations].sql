IF NOT EXISTS (SELECT 1 FROM [dbo].[WorkflowConfigurations] WHERE [WorkflowCode] = 'VoucherDeletion')
    INSERT INTO [dbo].[WorkflowConfigurations] ([WorkflowCode], [WorkflowName], [Module], [Description], [IsEnabled], [ApproverRole])
    VALUES ('VoucherDeletion', 'Voucher Deletion Approval', 'Accounts', 'Requires approval when non-exempt users request deletion of an accounting voucher.', 1, 'Director');

INSERT INTO [dbo].[WorkflowExemptRoles] ([WorkflowCode], [RoleName])
SELECT e.Code, e.RoleName
FROM (VALUES ('VoucherDeletion', 'Director'), ('VoucherDeletion', 'Admin'), ('VoucherDeletion', 'Administrator')) e(Code, RoleName)
WHERE NOT EXISTS (
    SELECT 1 FROM [dbo].[WorkflowExemptRoles] r WHERE r.WorkflowCode = e.Code AND r.RoleName = e.RoleName
);