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
            SELECT 
                v.EntryID,
                v.DT,
                ISNULL(v.GatePassNo, '') AS GatePassNo,
                ISNULL(v.PersonName, '') AS PersonName,
                ISNULL(v.VehicleNo, '') AS VehicleNo,
                ISNULL(v.Description, '') AS Description,
                ISNULL(v.DeptID, '') AS DeptID,
                ISNULL(v.DeptName, '') AS DeptName,
                ISNULL(v.EmpID, '') AS EmpID,
                ISNULL(v.EmpName, '') AS EmpName,
                ISNULL(v.GatePassType, 0) AS GatePassType,
                ISNULL(v.GatePassType2, 0) AS GatePassType2,
                v.VM_AccNo,
                v.Name,
                v.VendNameManual,
                CAST(ISNULL(v.Approved_HOD, 0) AS bit) AS Approved_HOD,
                CAST(ISNULL(v.Approved, 0) AS bit) AS Approved,
                CAST(ISNULL(v.Approved_Finance, 0) AS bit) AS Approved_Finance
            FROM dbo.VMiscEntries_GateInTemp AS v
            WHERE v.OnlyDT BETWEEN @FromDate AND @ToDate
              AND (@DeptID IS NULL OR @DeptID = '' OR v.DeptID = @DeptID)
              AND (@GatePassType IS NULL OR @GatePassType = 99 OR v.GatePassType = @GatePassType)
              AND (@VM_AccNo IS NULL OR @VM_AccNo = '' OR v.VM_AccNo = @VM_AccNo)
            ORDER BY v.DT DESC, v.EntryID DESC;
            """;

        var rows = new List<GateInListRowModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@FromDate", SqlDbType.DateTime).Value = filter.FromDate.Date;
        command.Parameters.Add("@ToDate", SqlDbType.DateTime).Value = filter.ToDate.Date.AddDays(1).AddTicks(-1);

        var deptVal = string.IsNullOrWhiteSpace(filter.DeptID) || filter.DeptID == "<All>" ? (object)DBNull.Value : filter.DeptID.Trim();
        command.Parameters.Add("@DeptID", SqlDbType.VarChar, 50).Value = deptVal;

        var typeVal = filter.GatePassType == GateInPassTypes.All ? (object)DBNull.Value : filter.GatePassType;
        command.Parameters.Add("@GatePassType", SqlDbType.Int).Value = typeVal;

        var accVal = string.IsNullOrWhiteSpace(filter.VM_AccNo) || filter.VM_AccNo == "0" ? (object)DBNull.Value : filter.VM_AccNo.Trim();
        command.Parameters.Add("@VM_AccNo", SqlDbType.VarChar, 50).Value = accVal;

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            var row = new GateInListRowModel
            {
                EntryId = reader.GetInt32(0),
                DT = reader.IsDBNull(1) ? null : reader.GetDateTime(1),
                GatePassNo = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                PersonName = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                VehicleNo = reader.IsDBNull(4) ? string.Empty : reader.GetString(4),
                Description = reader.IsDBNull(5) ? string.Empty : reader.GetString(5),
                DeptID = reader.IsDBNull(6) ? string.Empty : reader.GetString(6),
                DeptName = reader.IsDBNull(7) ? string.Empty : reader.GetString(7),
                EmpID = reader.IsDBNull(8) ? string.Empty : reader.GetString(8),
                EmpName = reader.IsDBNull(9) ? string.Empty : reader.GetString(9),
                GatePassType = reader.IsDBNull(10) ? 0 : reader.GetInt32(10),
                GatePassType2 = reader.IsDBNull(11) ? 0 : reader.GetInt32(11),
                VM_AccNo = reader.IsDBNull(12) ? null : reader.GetString(12),
                Name = reader.IsDBNull(13) ? null : reader.GetString(13),
                VendNameManual = reader.IsDBNull(14) ? null : reader.GetString(14),
                Approved_HOD = !reader.IsDBNull(15) && reader.GetBoolean(15),
                Approved = !reader.IsDBNull(16) && reader.GetBoolean(16),
                Approved_Finance = !reader.IsDBNull(17) && reader.GetBoolean(17)
            };

            if (!string.IsNullOrWhiteSpace(filter.SearchText))
            {
                var term = filter.SearchText.Trim();
                var matches = (row.GatePassNo?.Contains(term, StringComparison.OrdinalIgnoreCase) ?? false)
                    || (row.PersonName?.Contains(term, StringComparison.OrdinalIgnoreCase) ?? false)
                    || (row.VehicleNo?.Contains(term, StringComparison.OrdinalIgnoreCase) ?? false)
                    || (row.Description?.Contains(term, StringComparison.OrdinalIgnoreCase) ?? false)
                    || (row.DeptName?.Contains(term, StringComparison.OrdinalIgnoreCase) ?? false)
                    || (row.EmpName?.Contains(term, StringComparison.OrdinalIgnoreCase) ?? false)
                    || (row.DisplayVendorName?.Contains(term, StringComparison.OrdinalIgnoreCase) ?? false);

                if (!matches) continue;
            }

            rows.Add(row);
        }

        return rows;
    }

    public async Task<GateInEntryModel?> GetGateInAsync(int entryId, CancellationToken cancellationToken = default)
    {
        const string sqlMaster = """
            SELECT 
                m.EntryID,
                m.DT,
                ISNULL(m.GatePassNo, '') AS GatePassNo,
                ISNULL(m.PersonName, '') AS PersonName,
                ISNULL(m.VehicleNo, '') AS VehicleNo,
                ISNULL(m.Description, '') AS Description,
                ISNULL(m.DeptID, '') AS DeptID,
                ISNULL(d.name, '') AS DeptName,
                ISNULL(m.EmpID, '') AS EmpID,
                ISNULL(emp.Name, '') AS EmpName,
                ISNULL(m.GatePassType, 0) AS GatePassType,
                ISNULL(m.GatePassType2, 0) AS GatePassType2,
                ISNULL(m.VendID, 0) AS VendID,
                m.VM_AccNo,
                vm.AccTitle AS VendorName,
                m.VendNameManual,
                CAST(ISNULL(m.Approved_HOD, 0) AS bit) AS Approved_HOD,
                CAST(ISNULL(m.Approved, 0) AS bit) AS Approved,
                CAST(ISNULL(m.Approved_Finance, 0) AS bit) AS Approved_Finance
            FROM dbo.MiscEntries_GateInTemp AS m
            LEFT OUTER JOIN dbo.Departments AS d ON m.DeptID = d.deptid
            LEFT OUTER JOIN dbo.Employees AS emp ON m.EmpID = emp.EmpID
            LEFT OUTER JOIN dbo.VVendersAndMakers AS vm ON m.VM_AccNo = vm.AccNo
            WHERE m.EntryID = @EntryID;
            """;

        const string sqlDetails = """
            SELECT 
                d.EntryID,
                ISNULL(d.MEGO_RefID, 0) AS MEGO_RefID,
                ISNULL(d.Detail, '') AS Detail,
                ISNULL(d.Unit, '') AS Unit,
                CAST(ISNULL(d.Qty, 0) AS decimal(18,4)) AS Qty,
                CAST(ISNULL(d.Rate, 0) AS decimal(18,4)) AS Rate
            FROM dbo.MiscEntries_GateInTemp_Details AS d
            WHERE d.MEGO_RefID = @EntryID
            ORDER BY d.EntryID ASC;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        GateInEntryModel? entry = null;
        await using (var masterCmd = new SqlCommand(sqlMaster, connection))
        {
            masterCmd.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
            await using var reader = await masterCmd.ExecuteReaderAsync(cancellationToken);
            if (await reader.ReadAsync(cancellationToken))
            {
                entry = new GateInEntryModel
                {
                    EntryId = reader.GetInt32(0),
                    DT = reader.IsDBNull(1) ? DateTime.Now : reader.GetDateTime(1),
                    GatePassNo = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                    PersonName = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                    VehicleNo = reader.IsDBNull(4) ? string.Empty : reader.GetString(4),
                    Description = reader.IsDBNull(5) ? string.Empty : reader.GetString(5),
                    DeptID = reader.IsDBNull(6) ? string.Empty : reader.GetString(6),
                    DeptName = reader.IsDBNull(7) ? string.Empty : reader.GetString(7),
                    EmpID = reader.IsDBNull(8) ? string.Empty : reader.GetString(8),
                    EmpName = reader.IsDBNull(9) ? string.Empty : reader.GetString(9),
                    GatePassType = reader.IsDBNull(10) ? 0 : reader.GetInt32(10),
                    GatePassType2 = reader.IsDBNull(11) ? 0 : reader.GetInt32(11),
                    VendorId = reader.IsDBNull(12) ? 0 : reader.GetInt32(12),
                    VM_AccNo = reader.IsDBNull(13) ? null : reader.GetString(13),
                    VendorName = reader.IsDBNull(14) ? null : reader.GetString(14),
                    VendNameManual = reader.IsDBNull(15) ? null : reader.GetString(15),
                    Approved_HOD = !reader.IsDBNull(16) && reader.GetBoolean(16),
                    Approved = !reader.IsDBNull(17) && reader.GetBoolean(17),
                    Approved_Finance = !reader.IsDBNull(18) && reader.GetBoolean(18)
                };
            }
        }

        if (entry == null) return null;

        await using (var detailsCmd = new SqlCommand(sqlDetails, connection))
        {
            detailsCmd.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
            await using var reader = await detailsCmd.ExecuteReaderAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
            {
                entry.Details.Add(new GateInDetailModel
                {
                    EntryId = reader.GetInt32(0),
                    MEGO_RefID = reader.GetInt32(1),
                    Detail = reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                    Unit = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                    Qty = reader.GetDecimal(4),
                    Rate = reader.GetDecimal(5)
                });
            }
        }

        return entry;
    }

    public async Task<string> GetNextGatePassNoAsync(DateTime date, CancellationToken cancellationToken = default)
    {
        var prefix = $"GP-{date:MMyy}-";
        const string sql = """
            SELECT ISNULL(MAX(TRY_CAST(RIGHT(GatePassNo, 4) AS int)), 0) + 1
            FROM dbo.MiscEntries_GateInTemp
            WHERE LEFT(GatePassNo, 8) = @Prefix;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@Prefix", SqlDbType.VarChar, 8).Value = prefix;

        var result = await command.ExecuteScalarAsync(cancellationToken);
        var nextSeq = result != null && result != DBNull.Value ? Convert.ToInt32(result) : 1;
        return $"{prefix}{nextSeq:D4}";
    }

    public async Task<int> SaveGateInAsync(
        GateInEntryModel entry,
        string userName,
        string machineName,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(entry);

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var transaction = connection.BeginTransaction();

        try
        {
            int entryId = entry.EntryId;

            if (entryId == 0)
            {
                if (string.IsNullOrWhiteSpace(entry.GatePassNo))
                {
                    entry.GatePassNo = await GetNextGatePassNoAsync(entry.DT, cancellationToken);
                }

                const string insertMasterSql = """
                    INSERT INTO dbo.MiscEntries_GateInTemp
                    (
                        DT, GatePassNo, PersonName, VehicleNo, Description,
                        UserName, MachineName, DTEntry, EmpID, GatePassType,
                        DeptID, GatePassType2, VendID, VendNameManual, VM_AccNo,
                        Approved_HOD, Approved, Approved_Finance
                    )
                    VALUES
                    (
                        @DT, @GatePassNo, @PersonName, @VehicleNo, @Description,
                        @UserName, @MachineName, GETDATE(), @EmpID, @GatePassType,
                        @DeptID, @GatePassType2, 0, @VendNameManual, @VM_AccNo,
                        0, 0, 0
                    );
                    SELECT CAST(SCOPE_IDENTITY() AS int);
                    """;

                await using var cmd = new SqlCommand(insertMasterSql, connection, transaction);
                cmd.Parameters.Add("@DT", SqlDbType.DateTime).Value = entry.DT;
                cmd.Parameters.Add("@GatePassNo", SqlDbType.VarChar, 50).Value = entry.GatePassNo?.Trim() ?? string.Empty;
                cmd.Parameters.Add("@PersonName", SqlDbType.VarChar, 255).Value = (object?)entry.PersonName?.Trim() ?? DBNull.Value;
                cmd.Parameters.Add("@VehicleNo", SqlDbType.VarChar, 50).Value = (object?)entry.VehicleNo?.Trim() ?? DBNull.Value;
                cmd.Parameters.Add("@Description", SqlDbType.VarChar, 255).Value = (object?)entry.Description?.Trim() ?? DBNull.Value;
                cmd.Parameters.Add("@UserName", SqlDbType.VarChar, 50).Value = (object?)userName ?? DBNull.Value;
                cmd.Parameters.Add("@MachineName", SqlDbType.VarChar, 50).Value = (object?)machineName ?? DBNull.Value;
                cmd.Parameters.Add("@EmpID", SqlDbType.VarChar, 50).Value = (object?)entry.EmpID?.Trim() ?? DBNull.Value;
                cmd.Parameters.Add("@GatePassType", SqlDbType.Int).Value = entry.GatePassType;
                cmd.Parameters.Add("@DeptID", SqlDbType.VarChar, 50).Value = (object?)entry.DeptID?.Trim() ?? DBNull.Value;
                cmd.Parameters.Add("@GatePassType2", SqlDbType.Int).Value = entry.GatePassType2;
                cmd.Parameters.Add("@VendNameManual", SqlDbType.VarChar, 500).Value = (object?)entry.VendNameManual?.Trim() ?? DBNull.Value;
                cmd.Parameters.Add("@VM_AccNo", SqlDbType.VarChar, 50).Value = (object?)entry.VM_AccNo?.Trim() ?? DBNull.Value;

                var idObj = await cmd.ExecuteScalarAsync(cancellationToken);
                entryId = Convert.ToInt32(idObj);
                entry.EntryId = entryId;
            }
            else
            {
                const string updateMasterSql = """
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
                        VendNameManual = @VendNameManual,
                        VM_AccNo = @VM_AccNo
                    WHERE EntryID = @EntryID;

                    DELETE FROM dbo.MiscEntries_GateInTemp_Details WHERE MEGO_RefID = @EntryID;
                    """;

                await using var cmd = new SqlCommand(updateMasterSql, connection, transaction);
                cmd.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
                cmd.Parameters.Add("@DT", SqlDbType.DateTime).Value = entry.DT;
                cmd.Parameters.Add("@GatePassNo", SqlDbType.VarChar, 50).Value = entry.GatePassNo?.Trim() ?? string.Empty;
                cmd.Parameters.Add("@PersonName", SqlDbType.VarChar, 255).Value = (object?)entry.PersonName?.Trim() ?? DBNull.Value;
                cmd.Parameters.Add("@VehicleNo", SqlDbType.VarChar, 50).Value = (object?)entry.VehicleNo?.Trim() ?? DBNull.Value;
                cmd.Parameters.Add("@Description", SqlDbType.VarChar, 255).Value = (object?)entry.Description?.Trim() ?? DBNull.Value;
                cmd.Parameters.Add("@EmpID", SqlDbType.VarChar, 50).Value = (object?)entry.EmpID?.Trim() ?? DBNull.Value;
                cmd.Parameters.Add("@GatePassType", SqlDbType.Int).Value = entry.GatePassType;
                cmd.Parameters.Add("@DeptID", SqlDbType.VarChar, 50).Value = (object?)entry.DeptID?.Trim() ?? DBNull.Value;
                cmd.Parameters.Add("@GatePassType2", SqlDbType.Int).Value = entry.GatePassType2;
                cmd.Parameters.Add("@VendNameManual", SqlDbType.VarChar, 500).Value = (object?)entry.VendNameManual?.Trim() ?? DBNull.Value;
                cmd.Parameters.Add("@VM_AccNo", SqlDbType.VarChar, 50).Value = (object?)entry.VM_AccNo?.Trim() ?? DBNull.Value;

                await cmd.ExecuteNonQueryAsync(cancellationToken);
            }

            // Insert details
            if (entry.Details != null && entry.Details.Count > 0)
            {
                const string insertDetailSql = """
                    INSERT INTO dbo.MiscEntries_GateInTemp_Details
                    (MEGO_RefID, Detail, Unit, Qty, Rate)
                    VALUES
                    (@MEGO_RefID, @Detail, @Unit, @Qty, @Rate);
                    """;

                foreach (var item in entry.Details)
                {
                    if (string.IsNullOrWhiteSpace(item.Detail)) continue;

                    await using var dCmd = new SqlCommand(insertDetailSql, connection, transaction);
                    dCmd.Parameters.Add("@MEGO_RefID", SqlDbType.Int).Value = entryId;
                    dCmd.Parameters.Add("@Detail", SqlDbType.VarChar, 255).Value = item.Detail.Trim();
                    dCmd.Parameters.Add("@Unit", SqlDbType.VarChar, 50).Value = (object?)item.Unit?.Trim() ?? DBNull.Value;
                    dCmd.Parameters.Add("@Qty", SqlDbType.Real).Value = Convert.ToSingle(item.Qty);
                    dCmd.Parameters.Add("@Rate", SqlDbType.Real).Value = Convert.ToSingle(item.Rate);

                    await dCmd.ExecuteNonQueryAsync(cancellationToken);
                }
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

    public Task<int> SaveGateInAsync(GateInEntryModel entry, string userName, CancellationToken cancellationToken = default)
        => SaveGateInAsync(entry, userName, Environment.MachineName, cancellationToken);

    public async Task<bool> SetApprovalAsync(int entryId, string approvalColumn, bool approved, CancellationToken cancellationToken = default)
    {
        var validCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Approved_HOD", "Approved", "Approved_Finance" };
        if (!validCols.Contains(approvalColumn))
            throw new ArgumentException("Invalid approval column.", nameof(approvalColumn));

        var sql = $"UPDATE dbo.MiscEntries_GateInTemp SET [{approvalColumn}] = @Approved WHERE EntryID = @EntryID;";
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        command.Parameters.Add("@Approved", SqlDbType.Bit).Value = approved;

        var rows = await command.ExecuteNonQueryAsync(cancellationToken);
        return rows > 0;
    }

    public async Task ApproveFinanceAsync(int entryId, string userName, CancellationToken cancellationToken = default)
    {
        await SetApprovalAsync(entryId, "Approved_Finance", true, cancellationToken);
    }

    public async Task<bool> DeleteGateInAsync(int entryId, CancellationToken cancellationToken = default)
    {
        const string sql = """
            DELETE FROM dbo.MiscEntries_GateInTemp_Details WHERE MEGO_RefID = @EntryID;
            DELETE FROM dbo.MiscEntries_GateInTemp WHERE EntryID = @EntryID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;

        var rows = await command.ExecuteNonQueryAsync(cancellationToken);
        return rows > 0;
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetDepartmentsAsync(CancellationToken cancellationToken = default)
    {
        const string sql = "SELECT deptid, name FROM dbo.Departments WHERE ISNULL(name, '') <> '' ORDER BY name;";
        var list = new List<GateLookupOptionModel>();

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            list.Add(new GateLookupOptionModel
            {
                Id = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                Text = reader.IsDBNull(1) ? string.Empty : reader.GetString(1)
            });
        }

        return list;
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetEmployeesAsync(string departmentId, CancellationToken cancellationToken = default)
    {
        var items = await SearchEmployeesAsync(string.Empty, departmentId, cancellationToken);
        return items.Select(x => new GateLookupOptionModel { Id = x.EmpID, Text = x.DisplayText }).ToList();
    }

    public async Task<IReadOnlyList<GateInEmployeeLookupModel>> SearchEmployeesAsync(
        string searchText,
        string? deptId = null,
        CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT TOP 50 e.EmpID, e.Name, e.FName, e.Designation, e.DeptID, d.name AS DeptName
            FROM dbo.Employees e
            LEFT JOIN dbo.Departments d ON e.DeptID = d.deptid
            WHERE (ISNULL(e.Active, 1) = 1)
              AND (@DeptID IS NULL OR @DeptID = '' OR e.DeptID = @DeptID)
              AND (
                  @SearchText IS NULL OR @SearchText = ''
                  OR e.EmpID LIKE '%' + @SearchText + '%'
                  OR e.Name LIKE '%' + @SearchText + '%'
                  OR e.FName LIKE '%' + @SearchText + '%'
                  OR e.Designation LIKE '%' + @SearchText + '%'
              )
            ORDER BY e.Name;
            """;

        var list = new List<GateInEmployeeLookupModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);

        var deptVal = string.IsNullOrWhiteSpace(deptId) || deptId == "<All>" ? (object)DBNull.Value : deptId.Trim();
        command.Parameters.Add("@DeptID", SqlDbType.VarChar, 50).Value = deptVal;

        var searchVal = string.IsNullOrWhiteSpace(searchText) ? (object)DBNull.Value : searchText.Trim();
        command.Parameters.Add("@SearchText", SqlDbType.NVarChar, 100).Value = searchVal;

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            list.Add(new GateInEmployeeLookupModel
            {
                EmpID = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                Name = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                FName = reader.IsDBNull(2) ? null : reader.GetString(2),
                Designation = reader.IsDBNull(3) ? null : reader.GetString(3),
                DeptID = reader.IsDBNull(4) ? null : reader.GetString(4),
                DeptName = reader.IsDBNull(5) ? null : reader.GetString(5)
            });
        }

        return list;
    }

    public async Task<GateInEmployeeLookupModel?> GetEmployeeByIdAsync(string empId, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(empId)) return null;

        const string sql = """
            SELECT TOP 1 e.EmpID, e.Name, e.FName, e.Designation, e.DeptID, d.name AS DeptName
            FROM dbo.Employees e
            LEFT JOIN dbo.Departments d ON e.DeptID = d.deptid
            WHERE e.EmpID = @EmpID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@EmpID", SqlDbType.VarChar, 50).Value = empId.Trim();

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (await reader.ReadAsync(cancellationToken))
        {
            return new GateInEmployeeLookupModel
            {
                EmpID = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                Name = reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                FName = reader.IsDBNull(2) ? null : reader.GetString(2),
                Designation = reader.IsDBNull(3) ? null : reader.GetString(3),
                DeptID = reader.IsDBNull(4) ? null : reader.GetString(4),
                DeptName = reader.IsDBNull(5) ? null : reader.GetString(5)
            };
        }

        return null;
    }

    public async Task<IReadOnlyList<GateInVendorLookupModel>> SearchVendorsAsync(string searchText, CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT TOP 50 AccNo, AccTitle
            FROM dbo.VVendersAndMakers
            WHERE (
                @SearchText IS NULL OR @SearchText = ''
                OR AccNo LIKE '%' + @SearchText + '%'
                OR AccTitle LIKE '%' + @SearchText + '%'
            )
            ORDER BY AccTitle;
            """;

        var list = new List<GateInVendorLookupModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);

        var searchVal = string.IsNullOrWhiteSpace(searchText) ? (object)DBNull.Value : searchText.Trim();
        command.Parameters.Add("@SearchText", SqlDbType.NVarChar, 100).Value = searchVal;

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            list.Add(new GateInVendorLookupModel
            {
                AccNo = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                AccTitle = reader.IsDBNull(1) ? string.Empty : reader.GetString(1)
            });
        }

        return list;
    }

    public async Task<GateInVendorLookupModel?> GetVendorByAccNoAsync(string accNo, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(accNo)) return null;

        const string sql = """
            SELECT TOP 1 AccNo, AccTitle
            FROM dbo.VVendersAndMakers
            WHERE AccNo = @AccNo;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@AccNo", SqlDbType.VarChar, 50).Value = accNo.Trim();

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (await reader.ReadAsync(cancellationToken))
        {
            return new GateInVendorLookupModel
            {
                AccNo = reader.IsDBNull(0) ? string.Empty : reader.GetString(0),
                AccTitle = reader.IsDBNull(1) ? string.Empty : reader.GetString(1)
            };
        }

        return null;
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetVendorOptionsAsync(CancellationToken cancellationToken = default)
    {
        var vendors = await SearchVendorsAsync(string.Empty, cancellationToken);
        return vendors.Select(v => new GateLookupOptionModel { Id = v.AccNo, Text = v.DisplayText }).ToList();
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetVendorFilterOptionsAsync(CancellationToken cancellationToken = default)
    {
        var vendors = await GetVendorOptionsAsync(cancellationToken);
        var list = new List<GateLookupOptionModel>
        {
            new() { Id = "", Text = "<All Vendors>" }
        };
        list.AddRange(vendors);
        return list;
    }

    public Task<IReadOnlyList<GateLookupOptionModel>> GetGatePassTypeOptionsAsync(CancellationToken cancellationToken = default)
    {
        IReadOnlyList<GateLookupOptionModel> list = new List<GateLookupOptionModel>
        {
            new() { Id = GateInPassTypes.Foc.ToString(), Text = GateInPassTypes.Caption(GateInPassTypes.Foc) },
            new() { Id = GateInPassTypes.ExportSale.ToString(), Text = GateInPassTypes.Caption(GateInPassTypes.ExportSale) },
            new() { Id = GateInPassTypes.Sales.ToString(), Text = GateInPassTypes.Caption(GateInPassTypes.Sales) }
        };
        return Task.FromResult(list);
    }

    // Compatibility methods for GateInReceive if ever invoked
    public async Task<GateReceiveModel?> GetReceiveAsync(int entryId, CancellationToken cancellationToken = default)
    {
        var gateIn = await GetGateInAsync(entryId, cancellationToken);
        if (gateIn == null) return null;

        return new GateReceiveModel
        {
            EntryId = gateIn.EntryId,
            GatePassNo = gateIn.GatePassNo,
            DateTime = gateIn.DT,
            GatePassTypeName = gateIn.GatePassTypeName,
            PersonName = gateIn.PersonName,
            VendorName = gateIn.VendorName ?? string.Empty,
            ReceiveDate = DateTime.Today,
            Lines = gateIn.Details.Select(d => new GateReceiveLineModel
            {
                DetailId = d.EntryId,
                Detail = d.Detail,
                Unit = d.Unit,
                Qty = d.Qty,
                AlreadyReceivedQty = d.ReceivedQty,
                BalanceQty = d.BalanceQty,
                ReceiveQty = d.BalanceQty
            }).ToList()
        };
    }

    public Task SaveReceiveAsync(GateReceiveModel model, string userName, CancellationToken cancellationToken = default)
    {
        // No-op or log for compatibility
        return Task.CompletedTask;
    }
}
