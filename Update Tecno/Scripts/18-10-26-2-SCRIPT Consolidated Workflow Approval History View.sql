
IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'PrdApprovalHistory')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('PrdApprovalHistory', 'Approval History', 'Production', 'ApprovalHistory');
END
ELSE
BEGIN
    UPDATE MenuOptions SET OptionName = 'Approval History', ModuleName = 'Production', FormName = 'ApprovalHistory' WHERE OptionID = 'PrdApprovalHistory';
END
GO

-- ====================================================================================================
-- SCRIPT: Consolidated Workflow Approval History View
-- ====================================================================================================
CREATE OR ALTER VIEW [dbo].[vw_WorkflowApprovalHistory]
AS
SELECT 
    'ProductionDeletionRequests' AS SourceTable,
    p.[Id] AS RecordId,
    p.[TaskId],
    CASE 
        WHEN p.[RequestType] = 'LotReceiving' THEN 'DeleteLotAuthorization'
        WHEN p.[RequestType] IN ('LotIssuance', 'MasterPOIssuance') THEN 'DeleteProductionIssuance'
        WHEN p.[RequestType] = 'SkipProcess' THEN 'SkipProcess'
        ELSE p.[RequestType]
    END AS WorkflowCode,
    ISNULL(wc.[WorkflowName], p.[RequestType]) AS WorkflowName,
    ISNULL(wc.[Module], 'Production') AS Module,
    p.[LotNo],
    p.[OrderNo],
    p.[ItemCode],
    p.[ItemName],
    p.[ProcessName],
    CASE 
        WHEN p.[RequestType] = 'SkipProcess' THEN NULL 
        ELSE p.[MakerName] 
    END AS MakerName,
    CASE 
        WHEN p.[RequestType] = 'SkipProcess' THEN p.[MakerName] 
        ELSE NULL 
    END AS TargetProcessName,
    p.[Qty],
    p.[Reason] AS RequestReason,
    p.[RequestedBy],
    p.[RequestedDT],
    p.[MachineName],
    p.[Status],
    p.[ReviewedBy] AS ActionBy,
    p.[ReviewedDT] AS ActionDT,
    p.[DirectorRemarks] AS ActionRemarks
FROM [dbo].[ProductionDeletionRequests] p WITH (NOLOCK)
LEFT JOIN [dbo].[WorkflowConfigurations] wc WITH (NOLOCK) ON wc.[WorkflowCode] = (
    CASE 
        WHEN p.[RequestType] = 'LotReceiving' THEN 'DeleteLotAuthorization'
        WHEN p.[RequestType] IN ('LotIssuance', 'MasterPOIssuance') THEN 'DeleteProductionIssuance'
        WHEN p.[RequestType] = 'SkipProcess' THEN 'SkipProcess'
        ELSE p.[RequestType]
    END
)

UNION ALL

SELECT 
    'TaskItems' AS SourceTable,
    t.[Id] AS RecordId,
    t.[Id] AS TaskId,
    CASE 
        WHEN t.[SourceEntityType] IN ('AttendanceWorkflow', 'AttendanceApproval') THEN 'AttendanceApproval'
        WHEN t.[SourceEntityType] IN ('MakerItemRateWorkflow', 'MakerRateChange') THEN 'MakerRateChange'
        ELSE t.[SourceEntityType]
    END AS WorkflowCode,
    ISNULL(wc.[WorkflowName], t.[SourceEntityType]) AS WorkflowName,
    ISNULL(wc.[Module], 'General') AS Module,
    NULL AS LotNo,
    NULL AS OrderNo,
    NULL AS ItemCode,
    t.[Title] AS ItemName,
    NULL AS ProcessName,
    NULL AS MakerName,
    NULL AS TargetProcessName,
    NULL AS Qty,
    t.[Description] AS RequestReason,
    t.[AssignedBy] AS RequestedBy,
    t.[CreatedAt] AS RequestedDT,
    NULL AS MachineName,
    CASE t.[Status]
        WHEN 2 THEN 'Approved'
        WHEN 3 THEN 'Rejected'
        WHEN 0 THEN 'Pending'
        WHEN 1 THEN 'Pending'
        ELSE 'Pending'
    END AS Status,
    t.[CompletedBy] AS ActionBy,
    t.[CompletedAt] AS ActionDT,
    (SELECT TOP 1 tc.[Content] FROM [dbo].[TaskComments] tc WITH (NOLOCK) WHERE tc.[TaskId] = t.[Id] ORDER BY tc.[Id] DESC) AS ActionRemarks
FROM [dbo].[TaskItems] t WITH (NOLOCK)
INNER JOIN [dbo].[WorkflowConfigurations] wc WITH (NOLOCK) ON wc.[WorkflowCode] = (
    CASE 
        WHEN t.[SourceEntityType] IN ('AttendanceWorkflow', 'AttendanceApproval') THEN 'AttendanceApproval'
        WHEN t.[SourceEntityType] IN ('MakerItemRateWorkflow', 'MakerRateChange') THEN 'MakerRateChange'
        ELSE t.[SourceEntityType]
    END
)
WHERE t.[SourceEntityType] NOT IN ('DeleteLotAuthorization', 'DeleteProductionIssuance', 'SkipProcess');
GO

