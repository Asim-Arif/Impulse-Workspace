BEGIN TRAN;

UPDATE d
SET d.ItemID = COALESCE(inv.ItemID, cat.CompItemID)
FROM dbo.CustomerComplaints_Detail d
INNER JOIN dbo.CustomerComplaints c ON c.EntryID = d.RefID
OUTER APPLY (
    SELECT TOP 1 v.ItemID
    FROM dbo.VCustomInvoiceItems v
    WHERE v.CustomInvoice = d.InvoiceNo
      AND v.CustItemID = d.ItemID
) inv
OUTER APPLY (
    SELECT TOP 1 fc.CompItemID
    FROM dbo.VFCustomerCatalog fc
    WHERE fc.CustCode = c.CustCode
      AND fc.ItemID = d.ItemID
) cat
WHERE d.RefID = 1058
  AND COALESCE(inv.ItemID, cat.CompItemID) IS NOT NULL;

-- Verify updated rows
SELECT EntryID, RefID, InvoiceNo, ItemID, Qty, ItemReason_Main, EmailReference
FROM dbo.CustomerComplaints_Detail
WHERE RefID = 1058
ORDER BY EntryID;

COMMIT TRAN;
-- ROLLBACK TRAN;