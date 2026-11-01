using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Gate;

public sealed class MiscGateInDataAccess : IMiscGateInDataAccess
{
    private readonly string _connectionString;

    public MiscGateInDataAccess(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection is not configured.");
    }

    private SqlConnection CreateConnection() => new(_connectionString);

    public async Task<IReadOnlyList<MiscGateInListRowModel>> GetVisitorsAsync(
        MiscGateInFilterModel filter,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(filter);

        const string sql = """
            SELECT vv.EntryID,
                   vv.DT,
                   ISNULL(vv.VisitorName, '') AS VisitorName,
                   ISNULL(vv.CompanyName, '') AS CompanyName,
                   ISNULL(vv.EmpID, '') AS EmpID,
                   ISNULL(vv.Name, '') AS EmployeeName,
                   ISNULL(vv.ConcernedPerson, '') AS ConcernedPerson,
                   ISNULL(vv.Purpose, '') AS Purpose,
                   vv.TimeOut,
                   ISNULL(vv.VendName, '') AS VendName,
                   ISNULL(vv.VendNameManual, '') AS VendNameManual
            FROM dbo.VVisitors AS vv
            WHERE vv.OnlyDT BETWEEN @FromDate AND @ToDate
              AND (@EmployeeName IS NULL OR ISNULL(vv.EmpNameManual, '') = @EmployeeName)
              AND (@VendorName IS NULL OR ISNULL(vv.VendNameManual, '') = @VendorName)
            ORDER BY vv.DT;
            """;

        var rows = new List<MiscGateInListRowModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@FromDate", SqlDbType.DateTime).Value = filter.FromDate.Date;
        command.Parameters.Add("@ToDate", SqlDbType.DateTime).Value = filter.ToDate.Date;
        AddNullableFilter(command, "@EmployeeName", filter.EmployeeName, "<All Employees>");
        AddNullableFilter(command, "@VendorName", filter.VendorName, "<All Vendors>");

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            var empId = GetString(reader, 4);
            var employeeName = GetString(reader, 5);
            var manualPerson = GetString(reader, 6);
            var vendorName = GetString(reader, 9);
            var manualVendor = GetString(reader, 10);

            rows.Add(new MiscGateInListRowModel
            {
                EntryId = reader.GetInt32(0),
                DateTime = reader.IsDBNull(1) ? null : reader.GetDateTime(1),
                VisitorName = GetString(reader, 2),
                CompanyName = GetString(reader, 3),
                ConcernedPerson = string.IsNullOrWhiteSpace(empId)
                    ? manualPerson
                    : $"{empId} ({employeeName})",
                Purpose = GetString(reader, 7),
                TimeOut = reader.IsDBNull(8) ? null : reader.GetDateTime(8),
                VendorName = string.IsNullOrWhiteSpace(vendorName) ? manualVendor : vendorName
            });
        }

        return rows;
    }

    public async Task<MiscGateInEntryModel?> GetVisitorAsync(int entryId, CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT TOP (1)
                   vv.EntryID, vv.DT, ISNULL(vv.VisitorName, ''), ISNULL(vv.NICNumber, ''), ISNULL(vv.CompanyName, ''),
                   ISNULL(vv.Purpose, ''), ISNULL(vv.MobileNo, ''), ISNULL(vv.CardNo, ''),
                   ISNULL(vv.EmpID, ''), ISNULL(vv.ConcernedPerson, ''), ISNULL(vv.EmpNameManual, ''),
                   ISNULL(vv.Remarks, ''), ISNULL(vv.VendID, 0), ISNULL(vv.VendNameManual, ''),
                   vv.TimeOut, ISNULL(vv.ConcernedPersonEmail, ''), ISNULL(vv.ConcernedPersonWhatsApp, ''),
                   CAST(ISNULL(vv.NotifyByEmail, 0) AS bit), CAST(ISNULL(vv.NotifyByWhatsApp, 0) AS bit),
                   ISNULL(vv.VisitCode, ''), ISNULL(vv.NotificationStatus, '')
            FROM dbo.VVisitors AS vv
            WHERE vv.EntryID = @EntryID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
        {
            return null;
        }

        var timeOut = reader.IsDBNull(14) ? (DateTime?)null : reader.GetDateTime(14);
        return new MiscGateInEntryModel
        {
            EntryId = reader.GetInt32(0),
            DateTime = reader.IsDBNull(1) ? DateTime.Now : reader.GetDateTime(1),
            VisitorName = GetString(reader, 2),
            CnicNo = GetString(reader, 3),
            CompanyName = GetString(reader, 4),
            Purpose = GetString(reader, 5),
            MobileNo = GetString(reader, 6),
            CardNo = GetString(reader, 7),
            EmployeeId = GetString(reader, 8),
            EmployeeText = FirstNonBlank(GetString(reader, 10), GetString(reader, 9)),
            Remarks = GetString(reader, 11),
            VendorId = reader.GetInt32(12),
            VendorText = FirstNonBlank(GetString(reader, 13), "<Select Vendor>"),
            HasTimeOut = timeOut.HasValue,
            TimeOut = timeOut,
            ConcernedPersonEmail = GetString(reader, 15),
            ConcernedPersonWhatsApp = GetString(reader, 16),
            NotifyByEmail = !reader.IsDBNull(17) && reader.GetBoolean(17),
            NotifyByWhatsApp = !reader.IsDBNull(18) && reader.GetBoolean(18),
            VisitCode = GetString(reader, 19),
            NotificationStatus = GetString(reader, 20)
        };
    }

    public async Task<MiscGateOutModel?> GetVisitorGateOutAsync(int entryId, CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT TOP (1)
                   vv.EntryID, vv.DT, ISNULL(vv.VisitorName, ''), ISNULL(vv.NICNumber, ''), ISNULL(vv.CompanyName, ''),
                   ISNULL(vv.Purpose, ''), ISNULL(vv.MobileNo, ''), ISNULL(vv.CardNo, ''),
                   ISNULL(vv.EmpID, ''), ISNULL(vv.Name, ''), ISNULL(vv.ConcernedPerson, ''), ISNULL(vv.EmpNameManual, ''),
                   ISNULL(vv.Remarks, ''), ISNULL(vv.VendName, ''), ISNULL(vv.VendNameManual, ''), vv.TimeOut,
                   ISNULL(vv.VisitCode, ''), ISNULL(vv.QrPayload, ''), ISNULL(vv.OutVerifiedBy, ''), vv.OutVerifiedAt
            FROM dbo.VVisitors AS vv
            WHERE vv.EntryID = @EntryID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
        {
            return null;
        }

        var inTime = reader.IsDBNull(1) ? DateTime.Now : reader.GetDateTime(1);
        var currentOut = reader.IsDBNull(15) ? (DateTime?)null : reader.GetDateTime(15);
        var empId = GetString(reader, 8);
        var employeeName = GetString(reader, 9);
        var manualPerson = FirstNonBlank(GetString(reader, 11), GetString(reader, 10));
        var vendorName = GetString(reader, 13);
        var manualVendor = GetString(reader, 14);

        return new MiscGateOutModel
        {
            EntryId = reader.GetInt32(0),
            DateTime = inTime,
            TimeOut = currentOut ?? DateTime.Now,
            VisitorName = GetString(reader, 2),
            CnicNo = GetString(reader, 3),
            CompanyName = GetString(reader, 4),
            Purpose = GetString(reader, 5),
            MobileNo = GetString(reader, 6),
            CardNo = GetString(reader, 7),
            ConcernedPerson = string.IsNullOrWhiteSpace(empId)
                ? manualPerson
                : $"[{empId}] {employeeName}",
            Remarks = GetString(reader, 12),
            VendorName = string.IsNullOrWhiteSpace(vendorName) ? manualVendor : vendorName,
            IsAlreadyOut = currentOut.HasValue,
            VisitCode = GetString(reader, 16),
            OutVerifiedBy = GetString(reader, 18),
            OutVerifiedAt = reader.IsDBNull(19) ? null : reader.GetDateTime(19)
        };
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetEmployeeOptionsAsync(CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT EmpID,
                   '[' + EmpID + '] ' + Name + ' (' + DeptName + ')' AS DisplayText
            FROM dbo.VEmp
            WHERE Active = 1
            ORDER BY Name;
            """;

        var options = await ReadLookupAsync(sql, cancellationToken);
        return options.Count == 0
            ? await GetEmployeeFilterOptionsAsync(cancellationToken)
            : options;
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetVendorOptionsAsync(CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT CONVERT(varchar(20), VendID) AS ID,
                   ISNULL(vac.AccTitle, '') AS DisplayText
            FROM dbo.Venders v
            LEFT JOIN dbo.Accounts vac ON v.AccNo = vac.AccNo
            ORDER BY vac.AccTitle;
            """;

        try
        {
            var options = await ReadLookupAsync(sql, cancellationToken);
            return WithFirst(options, "<Select Vendor>", "0");
        }
        catch
        {
            var fallbackOptions = await GetVendorFilterOptionsAsync(cancellationToken);
            return WithFirst(
                fallbackOptions.Where(x => !string.Equals(x.Text, "<All Vendors>", StringComparison.OrdinalIgnoreCase)).ToArray(),
                "<Select Vendor>",
                "0");
        }
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetEmployeeFilterOptionsAsync(CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT DISTINCT v.EmpNameManual, v.EmpNameManual
            FROM dbo.Visitors AS v
            WHERE ISNULL(v.EmpNameManual, '0') <> '0'
            ORDER BY v.EmpNameManual;
            """;
        return WithFirst(await ReadLookupAsync(sql, cancellationToken), "<All Employees>", "");
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetVendorFilterOptionsAsync(CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT DISTINCT v.VendNameManual, v.VendNameManual
            FROM dbo.Visitors AS v
            WHERE ISNULL(v.VendNameManual, '0') <> '0'
            ORDER BY v.VendNameManual;
            """;
        return WithFirst(await ReadLookupAsync(sql, cancellationToken), "<All Vendors>", "");
    }

    public async Task<ConcernedPersonContactModel?> GetConcernedPersonContactAsync(string employeeId, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(employeeId))
        {
            return null;
        }

        const string sql = """
            SELECT TOP (1) EmpID, ISNULL(ContactName, ''), ISNULL(EmailAddress, ''), ISNULL(WhatsAppNo, '')
            FROM dbo.GateConcernedContacts
            WHERE EmpID = @EmpID AND Active = 1;
            """;
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@EmpID", SqlDbType.VarChar, 50).Value = employeeId.Trim();
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
        {
            return null;
        }

        return new ConcernedPersonContactModel
        {
            EmployeeId = reader.GetString(0),
            ContactName = reader.GetString(1),
            EmailAddress = reader.GetString(2),
            WhatsAppNo = reader.GetString(3)
        };
    }

    public async Task<int> SaveVisitorAsync(MiscGateInEntryModel visitor, string userName, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(visitor);
        if (string.IsNullOrWhiteSpace(visitor.VisitorName))
        {
            throw new InvalidOperationException("Please enter visitor name.");
        }

        var timeOut = visitor.HasTimeOut ? visitor.TimeOut ?? visitor.DateTime : (DateTime?)null;
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        int entryId;
        if (visitor.EntryId == 0)
        {
            const string insertSql = """
                INSERT INTO dbo.Visitors
                (
                    DT, VisitorName, NICNumber, Purpose, CompanyName, EmpID,
                    UserName, MachineName, TimeOut, Remarks, ConcernedPerson,
                    MobileNo, CardNo, VendID, EmpNameManual, VendNameManual,
                    ConcernedPersonEmail, ConcernedPersonWhatsApp, NotifyByEmail, NotifyByWhatsApp,
                    VisitCode, QrPayload
                )
                OUTPUT INSERTED.EntryID
                VALUES
                (
                    @DT, @VisitorName, @NICNumber, @Purpose, @CompanyName, @EmpID,
                    @UserName, @MachineName, @TimeOut, @Remarks, @ConcernedPerson,
                    @MobileNo, @CardNo, @VendID, @EmpNameManual, @VendNameManual,
                    @ConcernedPersonEmail, @ConcernedPersonWhatsApp, @NotifyByEmail, @NotifyByWhatsApp,
                    @VisitCode, @QrPayload
                );
                """;

            await using var insert = new SqlCommand(insertSql, connection);
            AddSaveParameters(insert, visitor, userName, timeOut);
            entryId = Convert.ToInt32(await insert.ExecuteScalarAsync(cancellationToken));
        }
        else
        {
            const string updateSql = """
                UPDATE dbo.Visitors
                SET DT = @DT,
                    VisitorName = @VisitorName,
                    NICNumber = @NICNumber,
                    Purpose = @Purpose,
                    CompanyName = @CompanyName,
                    EmpID = @EmpID,
                    TimeOut = @TimeOut,
                    Remarks = @Remarks,
                    ConcernedPerson = @ConcernedPerson,
                    MobileNo = @MobileNo,
                    CardNo = @CardNo,
                    VendID = @VendID,
                    EmpNameManual = @EmpNameManual,
                    VendNameManual = @VendNameManual,
                    ConcernedPersonEmail = @ConcernedPersonEmail,
                    ConcernedPersonWhatsApp = @ConcernedPersonWhatsApp,
                    NotifyByEmail = @NotifyByEmail,
                    NotifyByWhatsApp = @NotifyByWhatsApp,
                    VisitCode = @VisitCode,
                    QrPayload = @QrPayload
                WHERE EntryID = @EntryID;
                """;

            await using var update = new SqlCommand(updateSql, connection);
            AddSaveParameters(update, visitor, userName, timeOut);
            update.Parameters.Add("@EntryID", SqlDbType.Int).Value = visitor.EntryId;
            if (await update.ExecuteNonQueryAsync(cancellationToken) == 0)
            {
                throw new InvalidOperationException("Visitor entry was not found.");
            }
            entryId = visitor.EntryId;
        }

        if (!string.IsNullOrWhiteSpace(visitor.EmployeeId)
            && (!string.IsNullOrWhiteSpace(visitor.ConcernedPersonEmail) || !string.IsNullOrWhiteSpace(visitor.ConcernedPersonWhatsApp)))
        {
            await UpsertConcernedContactAsync(connection, visitor, userName, cancellationToken);
        }

        return entryId;
    }

    public async Task SaveVisitorGateOutAsync(int entryId, DateTime timeOut, string verificationInput, string userName, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        const string verifySql = """
            SELECT TOP (1) ISNULL(CardNo, ''), ISNULL(VisitCode, '')
            FROM dbo.Visitors
            WHERE EntryID = @EntryID;
            """;
        await using var verify = new SqlCommand(verifySql, connection);
        verify.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        await using var reader = await verify.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("Visitor entry was not found.");
        }

        var cardNo = GetString(reader, 0);
        var visitCode = GetString(reader, 1);
        await reader.CloseAsync();

        var normalizedInput = (verificationInput ?? "").Trim();
        if (string.IsNullOrWhiteSpace(normalizedInput))
        {
            throw new InvalidOperationException("Scan or enter card no / QR code before checkout.");
        }

        if (!string.Equals(normalizedInput, cardNo, StringComparison.OrdinalIgnoreCase)
            && !string.Equals(normalizedInput, visitCode, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("The entered card no / QR code does not match this visitor.");
        }

        const string sql = """
            UPDATE dbo.Visitors
            SET TimeOut = @TimeOut,
                UserName = @UserName,
                MachineName = @MachineName,
                OutVerifiedBy = @UserName,
                OutVerifiedAt = GETDATE(),
                OutVerificationInput = @VerificationInput
            WHERE EntryID = @EntryID;
            """;
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@TimeOut", SqlDbType.DateTime).Value = timeOut;
        AddVarChar(command, "@UserName", 50, userName);
        AddVarChar(command, "@MachineName", 50, Environment.MachineName);
        AddVarChar(command, "@VerificationInput", 100, normalizedInput);
        command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;

        if (await command.ExecuteNonQueryAsync(cancellationToken) == 0)
        {
            throw new InvalidOperationException("Visitor entry was not found.");
        }
    }

    public async Task<VisitorPassCardModel> EnsurePassAsync(int entryId, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        const string sql = """
            SELECT TOP (1) EntryID, ISNULL(VisitorName, ''), ISNULL(CompanyName, ''), ISNULL(ConcernedPerson, ''),
                   ISNULL(CardNo, ''), ISNULL(VisitCode, ''), ISNULL(QrPayload, ''), DT
            FROM dbo.Visitors
            WHERE EntryID = @EntryID;
            """;
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("Visitor entry was not found.");
        }

        var visitorName = reader.GetString(1);
        var companyName = reader.GetString(2);
        var concernedPerson = reader.GetString(3);
        var cardNo = reader.GetString(4);
        var visitCode = reader.GetString(5);
        var qrPayload = reader.GetString(6);
        var entryTime = reader.IsDBNull(7) ? DateTime.Now : reader.GetDateTime(7);

        if (string.IsNullOrWhiteSpace(visitCode))
        {
            visitCode = $"VST-{entryId:000000}";
        }

        if (string.IsNullOrWhiteSpace(cardNo))
        {
            cardNo = $"CARD-{entryId:000000}";
        }

        if (string.IsNullOrWhiteSpace(qrPayload))
        {
            qrPayload = $"VISITOR|ENTRY={entryId}|CODE={visitCode}|CARD={cardNo}";
        }

        await reader.CloseAsync();

        const string updateSql = """
            UPDATE dbo.Visitors
            SET CardNo = @CardNo,
                VisitCode = @VisitCode,
                QrPayload = @QrPayload
            WHERE EntryID = @EntryID;
            """;
        await using var update = new SqlCommand(updateSql, connection);
        update.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        update.Parameters.AddWithValue("@CardNo", cardNo);
        update.Parameters.AddWithValue("@VisitCode", visitCode);
        update.Parameters.AddWithValue("@QrPayload", qrPayload);
        await update.ExecuteNonQueryAsync(cancellationToken);

        return new VisitorPassCardModel
        {
            EntryId = entryId,
            VisitorName = visitorName,
            CompanyName = companyName,
            ConcernedPerson = concernedPerson,
            CardNo = cardNo,
            VisitCode = visitCode,
            EntryTime = entryTime,
            QrSvgDataUrl = GenerateQrSvgDataUrl(qrPayload)
        };
    }

    private static string GenerateQrSvgDataUrl(string payload)
    {
        var escapedPayload = System.Security.SecurityElement.Escape(payload) ?? "";
        var svg = $"""<svg xmlns="http://www.w3.org/2000/svg" width="160" height="160" viewBox="0 0 160 160"><rect width="100%" height="100%" fill="#ffffff"/><rect x="10" y="10" width="40" height="40" fill="none" stroke="#000" stroke-width="8"/><rect x="22" y="22" width="16" height="16" fill="#000"/><rect x="110" y="10" width="40" height="40" fill="none" stroke="#000" stroke-width="8"/><rect x="122" y="22" width="16" height="16" fill="#000"/><rect x="10" y="110" width="40" height="40" fill="none" stroke="#000" stroke-width="8"/><rect x="22" y="122" width="16" height="16" fill="#000"/><rect x="65" y="20" width="10" height="30" fill="#000"/><rect x="80" y="20" width="10" height="15" fill="#000"/><rect x="20" y="65" width="30" height="10" fill="#000"/><rect x="65" y="65" width="30" height="30" fill="#000"/><rect x="110" y="65" width="20" height="10" fill="#000"/><rect x="65" y="110" width="15" height="30" fill="#000"/><rect x="90" y="120" width="25" height="20" fill="#000"/><rect x="125" y="110" width="15" height="30" fill="#000"/><text x="80" y="155" font-family="monospace" font-size="9" text-anchor="middle" fill="#555">{escapedPayload}</text></svg>""";
        return "data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));
    }

    private static async Task UpsertConcernedContactAsync(SqlConnection connection, MiscGateInEntryModel visitor, string userName, CancellationToken cancellationToken)
    {
        const string sql = """
            MERGE dbo.GateConcernedContacts AS target
            USING (SELECT @EmpID AS EmpID) AS source
            ON target.EmpID = source.EmpID
            WHEN MATCHED THEN
                UPDATE SET ContactName = @ContactName,
                           EmailAddress = @EmailAddress,
                           WhatsAppNo = @WhatsAppNo,
                           Active = 1,
                           LastUpdatedBy = @UpdatedBy,
                           LastUpdatedAt = GETDATE()
            WHEN NOT MATCHED THEN
                INSERT (EmpID, ContactName, EmailAddress, WhatsAppNo, Active, LastUpdatedBy)
                VALUES (@EmpID, @ContactName, @EmailAddress, @WhatsAppNo, 1, @UpdatedBy);
            """;
        await using var command = new SqlCommand(sql, connection);
        AddVarChar(command, "@EmpID", 50, visitor.EmployeeId);
        AddVarChar(command, "@ContactName", 150, visitor.EmployeeText);
        AddVarChar(command, "@EmailAddress", 150, visitor.ConcernedPersonEmail);
        AddVarChar(command, "@WhatsAppNo", 50, visitor.ConcernedPersonWhatsApp);
        AddVarChar(command, "@UpdatedBy", 50, userName);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private async Task<IReadOnlyList<GateLookupOptionModel>> ReadLookupAsync(string sql, CancellationToken cancellationToken)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        var options = new List<GateLookupOptionModel>();
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            var id = GetString(reader, 0);
            var text = GetString(reader, 1);
            if (!string.IsNullOrWhiteSpace(text)
                && options.All(x => !string.Equals(x.Text, text, StringComparison.OrdinalIgnoreCase)))
            {
                options.Add(new GateLookupOptionModel { Id = id, Text = text });
            }
        }
        return options;
    }

    private static IReadOnlyList<GateLookupOptionModel> WithFirst(IReadOnlyList<GateLookupOptionModel> options, string text, string id)
    {
        var rows = new List<GateLookupOptionModel> { new() { Id = id, Text = text } };
        rows.AddRange(options.Where(x => !string.Equals(x.Text, text, StringComparison.OrdinalIgnoreCase)));
        return rows;
    }

    private static void AddSaveParameters(SqlCommand command, MiscGateInEntryModel visitor, string userName, DateTime? timeOut)
    {
        command.Parameters.Add("@DT", SqlDbType.DateTime).Value = visitor.DateTime;
        AddVarChar(command, "@VisitorName", 50, ToUpper(visitor.VisitorName));
        AddVarChar(command, "@NICNumber", 50, ToUpper(visitor.CnicNo));
        AddVarChar(command, "@Purpose", 255, ToUpper(visitor.Purpose));
        AddVarChar(command, "@CompanyName", 50, ToUpper(visitor.CompanyName));
        AddVarChar(command, "@EmpID", 50, visitor.EmployeeId);
        AddVarChar(command, "@UserName", 50, userName);
        AddVarChar(command, "@MachineName", 50, Environment.MachineName);
        command.Parameters.Add("@TimeOut", SqlDbType.DateTime).Value = timeOut ?? (object)DBNull.Value;
        AddVarChar(command, "@Remarks", 255, ToUpper(visitor.Remarks));
        AddVarChar(command, "@ConcernedPerson", 50, ToUpper(visitor.EmployeeText));
        AddVarChar(command, "@MobileNo", 50, visitor.MobileNo);
        AddVarChar(command, "@CardNo", 50, ToUpper(visitor.CardNo));
        command.Parameters.Add("@VendID", SqlDbType.Int).Value = visitor.VendorId;
        AddVarChar(command, "@EmpNameManual", 50, visitor.EmployeeText);
        AddVarChar(command, "@VendNameManual", 50, FirstNonBlank(visitor.VendorText, "<Select Vendor>"));
        AddVarChar(command, "@ConcernedPersonEmail", 150, visitor.ConcernedPersonEmail);
        AddVarChar(command, "@ConcernedPersonWhatsApp", 50, visitor.ConcernedPersonWhatsApp);
        command.Parameters.Add("@NotifyByEmail", SqlDbType.Bit).Value = visitor.NotifyByEmail;
        command.Parameters.Add("@NotifyByWhatsApp", SqlDbType.Bit).Value = visitor.NotifyByWhatsApp;
        AddVarChar(command, "@VisitCode", 50, visitor.VisitCode);
        AddVarChar(command, "@QrPayload", 255, string.IsNullOrWhiteSpace(visitor.VisitCode) ? null : $"VISITOR|CODE={visitor.VisitCode}");
    }

    private static void AddNullableFilter(SqlCommand command, string name, string value, string allText)
    {
        command.Parameters.Add(name, SqlDbType.VarChar, 50).Value =
            string.IsNullOrWhiteSpace(value) || string.Equals(value, allText, StringComparison.OrdinalIgnoreCase)
                ? DBNull.Value
                : value.Trim();
    }

    private static void AddVarChar(SqlCommand command, string name, int length, string? value)
    {
        command.Parameters.Add(name, SqlDbType.VarChar, length).Value =
            string.IsNullOrWhiteSpace(value) ? DBNull.Value : Truncate(value.Trim(), length);
    }

    private static string GetString(SqlDataReader reader, int ordinal) =>
        reader.IsDBNull(ordinal) ? "" : Convert.ToString(reader.GetValue(ordinal)) ?? "";

    private static string FirstNonBlank(params string?[] values) =>
        values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))?.Trim() ?? "";

    private static string ToUpper(string? value) =>
        string.IsNullOrWhiteSpace(value) ? "" : value.Trim().ToUpperInvariant();

    private static string Truncate(string value, int maxLength) =>
        value.Length <= maxLength ? value : value[..maxLength];
}
