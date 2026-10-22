CREATE VIEW [dbo].[VItems_Parts]
AS
SELECT dbo.Items_Parts.EntryID, dbo.Items_Parts.ItemID, dbo.Items_Parts.ItemID_Part, dbo.Items_Parts.Qty, dbo.Items_Parts.ProcessID, dbo.Items.ItemName, T1.ItemName AS ItemName_Part, dbo.Processes.Description
FROM  dbo.Items_Parts INNER JOIN
         dbo.Items ON dbo.Items.ItemID = dbo.Items_Parts.ItemID INNER JOIN
         dbo.Items AS T1 ON dbo.Items_Parts.ItemID_Part = T1.ItemID INNER JOIN
         dbo.Processes ON dbo.Items_Parts.ProcessID = dbo.Processes.ProcessID
GO