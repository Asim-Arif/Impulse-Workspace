-- Turn ON for a client DB ('True') or OFF ('False')
IF EXISTS (SELECT 1 FROM GeneralData WHERE DataName = 'AutoGenerateProformaOnOrderSave')
    UPDATE GeneralData SET DataValue = 'False' WHERE DataName = 'AutoGenerateProformaOnOrderSave';
ELSE
    INSERT INTO GeneralData (DataName, DataValue) VALUES ('AutoGenerateProformaOnOrderSave', 'False');