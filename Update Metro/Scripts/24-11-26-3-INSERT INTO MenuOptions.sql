IF NOT EXISTS (SELECT 1 FROM MenuOptions WHERE OptionID = 'PrlAttendanceApproval')
BEGIN
    INSERT INTO MenuOptions (OptionID, OptionName, ModuleName, FormName)
    VALUES ('PrlAttendanceApproval', 'Attendance Approvals', 'Payroll', 'AttendanceApproval');
END
ELSE
BEGIN
    UPDATE MenuOptions
    SET OptionName = 'Attendance Approvals',
        ModuleName = 'Payroll',
        FormName = 'AttendanceApproval'
    WHERE OptionID = 'PrlAttendanceApproval';
END
GO