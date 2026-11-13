-- Turn ON for a client DB ('True') or OFF ('False')
IF EXISTS (SELECT 1 FROM GeneralData WHERE DataName = 'AutoGenerateProformaOnOrderSave')
    UPDATE GeneralData SET DataValue = 'True' WHERE DataName = 'AutoGenerateProformaOnOrderSave';
ELSE
    INSERT INTO GeneralData (DataName, DataValue) VALUES ('AutoGenerateProformaOnOrderSave', 'True');