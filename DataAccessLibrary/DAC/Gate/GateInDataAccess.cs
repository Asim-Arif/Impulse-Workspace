using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Gate;

public sealed class GateInDataAccess : IGateInDataAccess
{
    private readonly string _connectionString;

    public GateInDataAccess(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection is not configured.");
    }

    private SqlConnection CreateConnection() => new(_connectionString);

    public async Task<IReadOnlyList<GateInListRowModel>> GetGateInEntriesAsync(
        GateInFilterModel filter,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(filter);

        const string sql = """
            SELECT v.EntryID,
                   v.DT,
                   ISNULL(v.GatePassNo, '') AS GatePassNo,
                   ISNULL(v.PersonName, '') AS PersonName,
                   ISNULL(v.VehicleNo, '') AS VehicleNo,
                   ISNULL(v.Description, '') AS Description,
                   ISNULL(v.VendNameManual, '') AS VendNameManual,
                   ISNULL(v.DeptID, '') AS DeptID,
                   ISNULL(v.GatePassType, 0) AS GatePassType,
                   ISNULL(m.GatePassTypeName, '') AS GatePassTypeName,
                   ISNULL(m.GatePassType2, 0) AS GatePassType2,
                   CAST(ISNULL(m.Approved_Finance, 0) AS bit) AS Approved_Finance,
                   CASE WHEN EXISTS
                        (
                            SELECT 1
                            FROM dbo.MiscEntries_GateInTemp_Details AS d
                            WHERE d.MEGO_RefID = m.EntryID
                              AND ISNULL(d.ReturnType, 'Non-Returnable') = 'Returnable'
                        ) THEN CAST(1 AS bit) ELSE CAST(0 AS bit) END AS HasReturnableItems,
                   CASE WHEN EXISTS
                        (
                            SELECT 1
                            FROM dbo.MiscEntries_GateInTemp_Details AS d
                            WHERE d.MEGO_RefID = m.EntryID
                              AND ISNULL(d.ReturnType, 'Non-Returnable') = 'Returnable'
                        )
                        AND NOT EXISTS
                        (
                            SELECT 1
                            FROM dbo.MiscEntries_GateInTemp_Details AS d
                            WHERE d.MEGO_RefID = m.EntryID
                              AND ISNULL(d.ReturnType, 'Non-Returnable') = 'Returnable'
                              AND ISNULL(d.QtyRcvd, 0) < ISNULL(d.Qty, 0)
                        ) THEN CAST(1 AS bit) ELSE CAST(0 AS bit) END AS IsReceived,
                   (
                        SELECT MAX(d.RcvdDT)
                        FROM dbo.MiscEntries_GateInTemp_Details AS d
                        WHERE d.MEGO_RefID = m.EntryID
                   ) AS ReceivedDate
            FROM dbo.VMiscEntries_GateInTemp AS v
            INNER JOIN dbo.MiscEntries_GateInTemp AS m ON m.EntryID = v.EntryID
            WHERE v.OnlyDT BETWEEN @FromDate AND @ToDate
              AND (@DeptID IS NULL OR v.DeptID = @DeptID)
              AND (@GatePassType IS NULL OR v.GatePassType = @GatePassType)
              AND (@GatePassTypeName IS NULL OR ISNULL(m.GatePassTypeName, '') = @GatePassTypeName)
              AND (@VendorName IS NULL OR ISNULL(v.VendNameManual, '') = @VendorName)
            ORDER BY v.DT;
            """;

        var rows = new List<GateInListRowModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@FromDate", SqlDbType.DateTime).Value = filter.FromDate.Date;
        command.Parameters.Add("@ToDate", SqlDbType.DateTime).Value = filter.ToDate.Date;
        AddNullableFilter(command, "@DeptID", filter.DepartmentId, "<All>");
        command.Parameters.Add("@GatePassType", SqlDbType.Int).Value = filter.GatePassType ?? (object)DBNull.Value;
        AddNullableFilter(command, "@GatePassTypeName", filter.GatePassTypeName, "<All Gatepass Types>");
        AddNullableFilter(command, "@VendorName", filter.VendorName, "<All Vendors>");

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            rows.Add(new GateInListRowModel
            {
                EntryId = reader.GetInt32(0),
                DateTime = reader.IsDBNull(1) ? null : reader.GetDateTime(1),
                GatePassNo = GetString(reader, 2),
                PersonName = GetString(reader, 3),
                VehicleNo = GetString(reader, 4),
                Description = GetString(reader, 5),
                VendorName = GetString(reader, 6),
                DepartmentId = GetString(reader, 7),
                GatePassType = reader.GetInt32(8),
                GatePassTypeName = GetString(reader, 9),
                GatePassKind = reader.GetInt32(10),
                IsFinanceApproved = reader.GetBoolean(11),
                HasReturnableItems = reader.GetBoolean(12),
                IsReceived = reader.GetBoolean(13),
                ReceivedDate = reader.IsDBNull(14) ? null : reader.GetDateTime(14)
            });
        }

        return rows;
    }

    public async Task<GateInEntryModel?> GetGateInAsync(
        int entryId,
        CancellationToken cancellationToken = default)
    {
        const string headerSql = """
            SELECT TOP (1)
                   EntryID,
                   DT,
                   ISNULL(GatePassNo, '') AS GatePassNo,
                   ISNULL(PersonName, '') AS PersonName,
                   ISNULL(VehicleNo, '') AS VehicleNo,
                   ISNULL(Description, '') AS Description,
                   ISNULL(EmpID, '') AS EmpID,
                   ISNULL(DeptID, '') AS DeptID,
                   ISNULL(GatePassType, 0) AS GatePassType,
                   ISNULL(GatePassType2, 0) AS GatePassType2,
                   ISNULL(VendID, 0) AS VendID,
                   ISNULL(EmpNameManual, '') AS EmpNameManual,
                   ISNULL(VendNameManual, '') AS VendNameManual,
                   CAST(ISNULL(Approved_Finance, 0) AS bit) AS Approved_Finance,
                   ISNULL(GatePassTypeName, '') AS GatePassTypeName
            FROM dbo.MiscEntries_GateInTemp
            WHERE EntryID = @EntryID;
            """;

        const string detailSql = """
            SELECT EntryID,
                   ISNULL(Detail, '') AS Detail,
                   ISNULL(Unit, '') AS Unit,
                   ISNULL(Qty, 0) AS Qty,
                   ISNULL(Rate, 0) AS Rate,
                   ISNULL(ReturnType, 'Non-Returnable') AS ReturnType,
                   ReturnDT,
                   ISNULL(QtyRcvd, 0) AS QtyRcvd,
                   RcvdDT,
                   ISNULL(RcvdBy, '') AS RcvdBy,
                   ISNULL(RcvdRemarks, '') AS RcvdRemarks
            FROM dbo.MiscEntries_GateInTemp_Details
            WHERE MEGO_RefID = @EntryID
            ORDER BY EntryID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        GateInEntryModel? entry;
        await using (var header = new SqlCommand(headerSql, connection))
        {
            header.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
            await using var reader = await header.ExecuteReaderAsync(cancellationToken);
            if (!await reader.ReadAsync(cancellationToken))
            {
                return null;
            }

            entry = new GateInEntryModel
            {
                EntryId = reader.GetInt32(0),
                DateTime = reader.IsDBNull(1) ? DateTime.Now : reader.GetDateTime(1),
                GatePassNo = GetString(reader, 2),
                PersonName = GetString(reader, 3),
                VehicleNo = GetString(reader, 4),
                Description = GetString(reader, 5),
                EmployeeId = GetString(reader, 6),
                DepartmentId = GetString(reader, 7),
                GatePassType = reader.GetInt32(8),
                GatePassKind = reader.GetInt32(9),
                VendorId = reader.GetInt32(10),
                EmployeeText = GetString(reader, 11),
                VendorText = FirstNonBlank(GetString(reader, 12), "<Select Vendor>"),
                IsFinanceApproved = reader.GetBoolean(13),
                GatePassTypeName = GetString(reader, 14)
            };
        }

        await using (var details = new SqlCommand(detailSql, connection))
        {
            details.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
            await using var reader = await details.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                entry.Details.Add(new GateInDetailModel
                {
                    EntryId = reader.GetInt32(0),
                    Detail = GetString(reader, 1),
                    Unit = GetString(reader, 2),
                    Qty = Convert.ToDecimal(reader.GetFloat(3)),
                    Rate = Convert.ToDecimal(reader.GetFloat(4)),
                    ReturnType = GetString(reader, 5),
                    ReturnDate = reader.IsDBNull(6) ? null : reader.GetDateTime(6),
                    ReceivedQty = Convert.ToDecimal(reader.GetFloat(7)),
                    ReceivedDate = reader.IsDBNull(8) ? null : reader.GetDateTime(8),
                    ReceivedBy = GetString(reader, 9),
                    ReceivedRemarks = GetString(reader, 10)
                });
            }
        }

        return entry;
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetDepartmentsAsync(
        CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT DeptID,
                   Name
            FROM dbo.Departments
            WHERE Active = 1
            ORDER BY Name;
            """;

        return await ReadLookupAsync(sql, cancellationToken);
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetEmployeesAsync(
        string departmentId,
        CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT EmpID,
                   '[' + EmpID + '] ' + ISNULL(Name, '') AS DisplayText
            FROM dbo.Employees
            WHERE Active = 1
            ORDER BY Name;
            """;

        return await ReadLookupAsync(sql, cancellationToken);
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetVendorOptionsAsync(
        CancellationToken cancellationToken = default)
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
            var fallback = await GetVendorFilterOptionsAsync(cancellationToken);
            return WithFirst(
                fallback.Where(x => !string.Equals(x.Text, "<All Vendors>", StringComparison.OrdinalIgnoreCase)).ToArray(),
                "<Select Vendor>",
                "0");
        }
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetVendorFilterOptionsAsync(
        CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT DISTINCT m.VendNameManual,
                   m.VendNameManual
            FROM dbo.MiscEntries_GateInTemp AS m
            WHERE ISNULL(m.VendNameManual, '0') <> '0'
            ORDER BY m.VendNameManual;
            """;

        var options = await ReadLookupAsync(sql, cancellationToken);
        return WithFirst(options, "<All Vendors>", "");
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetGatePassTypeOptionsAsync(
        CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT CONVERT(varchar(20), EntryID) AS ID,
                   GatePassType AS DisplayText
            FROM dbo.GatePassTypes
            WHERE ISNULL(Active, 1) = 1
            ORDER BY GatePassType;
            """;

        var options = await ReadLookupAsync(sql, cancellationToken);
        return WithFirst(options, "Select gate-pass type", "");
    }

    public async Task<string> GetNextGatePassNoAsync(
        DateTime date,
        CancellationToken cancellationToken = default)
    {
        var prefix = $"GP-{date:MMyy}-";
        const string sql = """
            SELECT ISNULL(MAX(CONVERT(int, RIGHT(GatePassNo, 5))), 0)
            FROM dbo.MiscEntries_GateInTemp
            WHERE LEFT(ISNULL(GatePassNo, ''), LEN(@Prefix)) = @Prefix
              AND LEN(ISNULL(GatePassNo, '')) >= LEN(@Prefix) + 5
              AND ISNUMERIC(RIGHT(GatePassNo, 5)) = 1
              AND MONTH(DT) = MONTH(@DT)
              AND YEAR(DT) = YEAR(@DT);
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        AddVarChar(command, "@Prefix", 20, prefix);
        command.Parameters.Add("@DT", SqlDbType.DateTime).Value = date;
        var next = Convert.ToInt32(await command.ExecuteScalarAsync(cancellationToken)) + 1;
        return prefix + next.ToString("D5");
    }

    public async Task<int> SaveGateInAsync(
        GateInEntryModel entry,
        string userName,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(entry);
        if (string.IsNullOrWhiteSpace(entry.DepartmentId))
        {
            throw new InvalidOperationException("Please select department.");
        }

        if (string.IsNullOrWhiteSpace(entry.GatePassNo))
        {
            throw new InvalidOperationException("Please enter Gate Pass No.");
        }

        if (entry.Details.Count == 0)
        {
            throw new InvalidOperationException("No items added.");
        }

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

        try
        {
            var entryId = entry.EntryId;
            if (entryId == 0)
            {
                const string insertSql = """
                    INSERT INTO dbo.MiscEntries_GateInTemp
                    (
                        DT, GatePassNo, PersonName, VehicleNo, Description,
                        UserName, MachineName, EmpID, GatePassType, DeptID,
                        Approved_HOD, Approved, Approved_Finance,
                        GatePassType2, VendID, EmpNameManual, VendNameManual, GatePassTypeName
                    )
                    OUTPUT INSERTED.EntryID
                    VALUES
                    (
                        @DT, @GatePassNo, @PersonName, @VehicleNo, @Description,
                        @UserName, @MachineName, @EmpID, @GatePassType, @DeptID,
                        0, 0, 0,
                        @GatePassType2, @VendID, @EmpNameManual, @VendNameManual, @GatePassTypeName
                    );
                    """;

                await using var insert = new SqlCommand(insertSql, connection, transaction);
                AddSaveParameters(insert, entry, userName);
                entryId = Convert.ToInt32(await insert.ExecuteScalarAsync(cancellationToken));
            }
            else
            {
                const string updateSql = """
                    UPDATE dbo.MiscEntries_GateInTemp
                    SET DT = @DT,
                        GatePassNo = @GatePassNo,
                        PersonName = @PersonName,
                        VehicleNo = @VehicleNo,
                        Description = @Description,
                        EmpID = @EmpID,
                        GatePassType = @GatePassType,
                        DeptID = @DeptID,
                        GatePassType2 = @GatePassType2,
                        VendID = @VendID,
                        EmpNameManual = @EmpNameManual,
                        VendNameManual = @VendNameManual,
                        GatePassTypeName = @GatePassTypeName
                    WHERE EntryID = @EntryID;
                    """;

                await using var update = new SqlCommand(updateSql, connection, transaction);
                AddSaveParameters(update, entry, userName);
                update.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
                if (await update.ExecuteNonQueryAsync(cancellationToken) == 0)
                {
                    throw new InvalidOperationException("Gate In entry was not found.");
                }

                await using var deleteDetails = new SqlCommand(
                    "DELETE FROM dbo.MiscEntries_GateInTemp_Details WHERE MEGO_RefID = @EntryID;",
                    connection,
                    transaction);
                deleteDetails.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
                await deleteDetails.ExecuteNonQueryAsync(cancellationToken);
            }

            const string detailSql = """
                INSERT INTO dbo.MiscEntries_GateInTemp_Details
                    (MEGO_RefID, Detail, Unit, Qty, Rate, ReturnType, ReturnDT, QtyRcvd, RcvdDT, RcvdBy, RcvdRemarks)
                VALUES
                    (@MEGO_RefID, @Detail, @Unit, @Qty, @Rate, @ReturnType, @ReturnDT, @QtyRcvd, @RcvdDT, @RcvdBy, @RcvdRemarks);
                """;

            foreach (var detail in entry.Details)
            {
                await using var insertDetail = new SqlCommand(detailSql, connection, transaction);
                insertDetail.Parameters.Add("@MEGO_RefID", SqlDbType.Int).Value = entryId;
                AddVarChar(insertDetail, "@Detail", 255, ToUpper(detail.Detail));
                AddVarChar(insertDetail, "@Unit", 50, ToUpper(detail.Unit));
                insertDetail.Parameters.Add("@Qty", SqlDbType.Real).Value = Convert.ToSingle(detail.Qty);
                insertDetail.Parameters.Add("@Rate", SqlDbType.Real).Value = Convert.ToSingle(detail.Rate);
                AddVarChar(insertDetail, "@ReturnType", 20, FirstNonBlank(detail.ReturnType, GateReturnTypes.NonReturnable));
                insertDetail.Parameters.Add("@ReturnDT", SqlDbType.DateTime).Value = detail.IsReturnable && detail.ReturnDate.HasValue
                    ? detail.ReturnDate.Value.Date
                    : DBNull.Value;
                insertDetail.Parameters.Add("@QtyRcvd", SqlDbType.Real).Value = Convert.ToSingle(detail.ReceivedQty);
                insertDetail.Parameters.Add("@RcvdDT", SqlDbType.DateTime).Value = detail.ReceivedDate.HasValue
                    ? detail.ReceivedDate.Value.Date
                    : DBNull.Value;
                AddVarChar(insertDetail, "@RcvdBy", 50, detail.ReceivedBy);
                AddVarChar(insertDetail, "@RcvdRemarks", 255, detail.ReceivedRemarks);
                await insertDetail.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
            return entryId;
        }
        catch
        {
            await transaction.RollbackAsync(cancellationToken);
            throw;
        }
    }

    public async Task<GateReceiveModel?> GetReceiveAsync(
        int entryId,
        CancellationToken cancellationToken = default)
    {
        const string headerSql = """
            SELECT TOP (1)
                   EntryID,
                   ISNULL(GatePassNo, '') AS GatePassNo,
                   ISNULL(DT, GETDATE()) AS DT,
                   ISNULL(GatePassTypeName, '') AS GatePassTypeName,
                   ISNULL(PersonName, '') AS PersonName,
                   ISNULL(VendNameManual, '') AS VendNameManual
            FROM dbo.MiscEntries_GateInTemp
            WHERE EntryID = @EntryID;
            """;

        const string lineSql = """
            SELECT EntryID,
                   ISNULL(Detail, '') AS Detail,
                   ISNULL(Unit, '') AS Unit,
                   ISNULL(Qty, 0) AS Qty,
                   ISNULL(QtyRcvd, 0) AS QtyRcvd
            FROM dbo.MiscEntries_GateInTemp_Details
            WHERE MEGO_RefID = @EntryID
              AND ISNULL(ReturnType, 'Non-Returnable') = 'Returnable'
              AND ISNULL(QtyRcvd, 0) < ISNULL(Qty, 0)
            ORDER BY EntryID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        GateReceiveModel? model;
        await using (var header = new SqlCommand(headerSql, connection))
        {
            header.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
            await using var reader = await header.ExecuteReaderAsync(cancellationToken);
            if (!await reader.ReadAsync(cancellationToken))
            {
                return null;
            }

            model = new GateReceiveModel
            {
                EntryId = reader.GetInt32(0),
                GatePassNo = GetString(reader, 1),
                DateTime = reader.GetDateTime(2),
                GatePassTypeName = GetString(reader, 3),
                PersonName = GetString(reader, 4),
                VendorName = GetString(reader, 5),
                ReceiveDate = DateTime.Today
            };
        }

        await using (var lines = new SqlCommand(lineSql, connection))
        {
            lines.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
            await using var reader = await lines.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                var qty = Convert.ToDecimal(reader.GetFloat(3));
                var already = Convert.ToDecimal(reader.GetFloat(4));
                model.Lines.Add(new GateReceiveLineModel
                {
                    DetailId = reader.GetInt32(0),
                    Detail = GetString(reader, 1),
                    Unit = GetString(reader, 2),
                    Qty = qty,
                    AlreadyReceivedQty = already,
                    BalanceQty = Math.Max(0, qty - already),
                    ReceiveQty = Math.Max(0, qty - already)
                });
            }
        }

        return model;
    }

    public async Task SaveReceiveAsync(
        GateReceiveModel model,
        string userName,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(model);
        var lines = model.Lines.Where(x => x.ReceiveQty > 0).ToArray();
        if (lines.Length == 0)
        {
            throw new InvalidOperationException("Please enter received quantity.");
        }

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

        try
        {
            foreach (var line in lines)
            {
                const string sql = """
                    UPDATE dbo.MiscEntries_GateInTemp_Details
                    SET QtyRcvd = CASE
                            WHEN ISNULL(QtyRcvd, 0) + @ReceiveQty > ISNULL(Qty, 0) THEN ISNULL(Qty, 0)
                            ELSE ISNULL(QtyRcvd, 0) + @ReceiveQty
                        END,
                        RcvdDT = @RcvdDT,
                        RcvdBy = @RcvdBy,
                        RcvdRemarks = @RcvdRemarks
                    WHERE EntryID = @DetailID
                      AND MEGO_RefID = @EntryID
                      AND ISNULL(ReturnType, 'Non-Returnable') = 'Returnable'
                      AND ISNULL(QtyRcvd, 0) < ISNULL(Qty, 0);
                    """;

                await using var command = new SqlCommand(sql, connection, transaction);
                command.Parameters.Add("@ReceiveQty", SqlDbType.Real).Value = Convert.ToSingle(line.ReceiveQty);
                command.Parameters.Add("@RcvdDT", SqlDbType.DateTime).Value = model.ReceiveDate.Date;
                AddVarChar(command, "@RcvdBy", 50, userName);
                AddVarChar(command, "@RcvdRemarks", 255, model.Remarks);
                command.Parameters.Add("@DetailID", SqlDbType.Int).Value = line.DetailId;
                command.Parameters.Add("@EntryID", SqlDbType.Int).Value = model.EntryId;
                await command.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
        }
        catch
        {
            await transaction.RollbackAsync(cancellationToken);
            throw;
        }
    }

    public async Task ApproveFinanceAsync(
        int entryId,
        string userName,
        CancellationToken cancellationToken = default)
    {
        const string sql = """
            UPDATE dbo.MiscEntries_GateInTemp
            SET Approved_Finance = 1
            WHERE EntryID = @EntryID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        if (await command.ExecuteNonQueryAsync(cancellationToken) == 0)
        {
            throw new InvalidOperationException("Gate In entry was not found.");
        }
    }

    private async Task<IReadOnlyList<GateLookupOptionModel>> ReadLookupAsync(
        string sql,
        CancellationToken cancellationToken)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        return await ReadLookupAsync(command, cancellationToken);
    }

    private static async Task<IReadOnlyList<GateLookupOptionModel>> ReadLookupAsync(
        SqlCommand command,
        CancellationToken cancellationToken)
    {
        var options = new List<GateLookupOptionModel>();
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            options.Add(new GateLookupOptionModel
            {
                Id = Convert.ToString(reader.GetValue(0))?.Trim() ?? "",
                Text = Convert.ToString(reader.GetValue(1))?.Trim() ?? ""
            });
        }

        return options;
    }

    private static IReadOnlyList<GateLookupOptionModel> WithFirst(
        IEnumerable<GateLookupOptionModel> items,
        string firstText,
        string firstId)
    {
        var list = new List<GateLookupOptionModel>
        {
            new() { Id = firstId, Text = firstText }
        };
        list.AddRange(items);
        return list;
    }

    private static void AddSaveParameters(SqlCommand command, GateInEntryModel entry, string userName)
    {
        command.Parameters.Add("@DT", SqlDbType.DateTime).Value = entry.DateTime;
        AddVarChar(command, "@GatePassNo", 50, entry.GatePassNo);
        AddVarChar(command, "@PersonName", 255, entry.PersonName);
        AddVarChar(command, "@VehicleNo", 50, entry.VehicleNo);
        AddVarChar(command, "@Description", 255, entry.Description);
        AddVarChar(command, "@UserName", 50, userName);
        AddVarChar(command, "@MachineName", 50, Environment.MachineName);
        AddVarChar(command, "@EmpID", 50, entry.EmployeeId);
        command.Parameters.Add("@GatePassType", SqlDbType.Int).Value = entry.GatePassType;
        AddVarChar(command, "@DeptID", 50, entry.DepartmentId);
        command.Parameters.Add("@GatePassType2", SqlDbType.Int).Value = entry.GatePassKind;
        command.Parameters.Add("@VendID", SqlDbType.Int).Value = entry.VendorId;
        AddVarChar(command, "@EmpNameManual", 50, entry.EmployeeText);
        AddVarChar(command, "@VendNameManual", 50, entry.VendorText);
        AddVarChar(command, "@GatePassTypeName", 50, entry.GatePassTypeName);
    }

    private static void AddNullableFilter(SqlCommand command, string name, string? value, string allPlaceholder)
    {
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value.Trim(), allPlaceholder, StringComparison.OrdinalIgnoreCase))
        {
            command.Parameters.Add(name, SqlDbType.VarChar, 100).Value = DBNull.Value;
        }
        else
        {
            command.Parameters.Add(name, SqlDbType.VarChar, 100).Value = value.Trim();
        }
    }

    private static void AddVarChar(SqlCommand command, string name, int size, string? value)
    {
        command.Parameters.Add(name, SqlDbType.VarChar, size).Value = string.IsNullOrWhiteSpace(value)
            ? (object)DBNull.Value
            : value.Trim();
    }

    private static string GetString(SqlDataReader reader, int ordinal) =>
        reader.IsDBNull(ordinal) ? "" : reader.GetString(ordinal).Trim();

    private static string ToUpper(string? value) => value?.Trim().ToUpperInvariant() ?? "";

    private static string FirstNonBlank(string? first, string second) =>
        string.IsNullOrWhiteSpace(first) ? second : first.Trim();
}

