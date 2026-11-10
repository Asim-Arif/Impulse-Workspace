using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Gate;

public sealed class GatePassPermanentDataAccess : IGatePassPermanentDataAccess
{
    private readonly string _connectionString;

    public GatePassPermanentDataAccess(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection") 
            ?? throw new InvalidOperationException("DefaultConnection connection string is not configured.");
    }

    private SqlConnection CreateConnection() => new(_connectionString);

    public async Task<IReadOnlyList<GatePassPermanentListRowModel>> GetListAsync(
        GatePassPermanentFilterModel filter, 
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(filter);

        const string sql = """
            SELECT EntryID,
                   GatePassNo,
                   DT,
                   ISNULL(PersonName, '') AS PersonName,
                   ISNULL(VehicleNo, '') AS VehicleNo,
                   ISNULL(Description, '') AS Description,
                   ISNULL(UserName, '') AS UserName,
                   ISNULL(MachineName, '') AS MachineName,
                   DTEntry,
                   OnlyDT,
                   ISNULL(EmpID, '') AS EmpID,
                   GatePassType,
                   ISNULL(DeptID, '') AS DeptID,
                   VendID,
                   ISNULL(Name, '') AS Name,
                   ISNULL(VendNameManual, '') AS VendNameManual,
                   ISNULL(VM_AccNo, '') AS VM_AccNo,
                   ISNULL(Approved_HOD, 0) AS Approved_HOD,
                   ISNULL(Approved, 0) AS Approved,
                   ISNULL(Approved_Finance, 0) AS Approved_Finance
            FROM dbo.VMiscEntries_GateOut
            WHERE DT >= @FromDate AND DT < @ToDate
              AND (@DeptID IS NULL OR DeptID = @DeptID)
              AND (@GatePassType IS NULL OR GatePassType = @GatePassType)
              AND (@VM_AccNo IS NULL OR VM_AccNo = @VM_AccNo)
              AND (
                  @SearchText IS NULL 
                  OR GatePassNo LIKE @SearchText 
                  OR PersonName LIKE @SearchText 
                  OR VehicleNo LIKE @SearchText 
                  OR Description LIKE @SearchText 
                  OR Name LIKE @SearchText
                  OR VendNameManual LIKE @SearchText
              )
            ORDER BY DT DESC, EntryID DESC;
            """;

        var list = new List<GatePassPermanentListRowModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);

        command.Parameters.Add("@FromDate", SqlDbType.DateTime).Value = filter.FromDate.Date;
        command.Parameters.Add("@ToDate", SqlDbType.DateTime).Value = filter.ToDate.Date.AddDays(1);
        command.Parameters.Add("@DeptID", SqlDbType.VarChar, 50).Value = !string.IsNullOrWhiteSpace(filter.DeptID) ? filter.DeptID.Trim() : DBNull.Value;
        command.Parameters.Add("@GatePassType", SqlDbType.Int).Value = filter.GatePassType.HasValue ? filter.GatePassType.Value : DBNull.Value;
        command.Parameters.Add("@VM_AccNo", SqlDbType.VarChar, 50).Value = !string.IsNullOrWhiteSpace(filter.VM_AccNo) ? filter.VM_AccNo.Trim() : DBNull.Value;

        var search = string.IsNullOrWhiteSpace(filter.SearchText) ? null : $"%{filter.SearchText.Trim()}%";
        command.Parameters.Add("@SearchText", SqlDbType.VarChar, 100).Value = (object?)search ?? DBNull.Value;

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            list.Add(new GatePassPermanentListRowModel
            {
                EntryID = reader.GetInt32(reader.GetOrdinal("EntryID")),
                GatePassNo = reader.IsDBNull(reader.GetOrdinal("GatePassNo")) ? null : reader.GetString(reader.GetOrdinal("GatePassNo")),
                DT = reader.IsDBNull(reader.GetOrdinal("DT")) ? null : reader.GetDateTime(reader.GetOrdinal("DT")),
                PersonName = reader.GetString(reader.GetOrdinal("PersonName")),
                VehicleNo = reader.GetString(reader.GetOrdinal("VehicleNo")),
                Description = reader.GetString(reader.GetOrdinal("Description")),
                UserName = reader.GetString(reader.GetOrdinal("UserName")),
                MachineName = reader.GetString(reader.GetOrdinal("MachineName")),
                DTEntry = reader.IsDBNull(reader.GetOrdinal("DTEntry")) ? null : reader.GetDateTime(reader.GetOrdinal("DTEntry")),
                OnlyDT = reader.IsDBNull(reader.GetOrdinal("OnlyDT")) ? null : reader.GetDateTime(reader.GetOrdinal("OnlyDT")),
                EmpID = reader.GetString(reader.GetOrdinal("EmpID")),
                GatePassType = reader.IsDBNull(reader.GetOrdinal("GatePassType")) ? null : reader.GetInt32(reader.GetOrdinal("GatePassType")),
                DeptID = reader.GetString(reader.GetOrdinal("DeptID")),
                VendID = reader.IsDBNull(reader.GetOrdinal("VendID")) ? null : reader.GetInt32(reader.GetOrdinal("VendID")),
                Name = reader.GetString(reader.GetOrdinal("Name")),
                VendNameManual = reader.GetString(reader.GetOrdinal("VendNameManual")),
                VM_AccNo = reader.GetString(reader.GetOrdinal("VM_AccNo")),
                Approved_HOD = reader.GetBoolean(reader.GetOrdinal("Approved_HOD")),
                Approved = reader.GetBoolean(reader.GetOrdinal("Approved")),
                Approved_Finance = reader.GetBoolean(reader.GetOrdinal("Approved_Finance"))
            });
        }

        return list;
    }

    public async Task<GatePassPermanentEntryModel?> GetByIdAsync(int entryId, CancellationToken cancellationToken = default)
    {
        const string masterSql = """
            SELECT EntryID,
                   GatePassNo,
                   DT,
                   ISNULL(DeptID, '') AS DeptID,
                   ISNULL(GatePassType, 0) AS GatePassType,
                   ISNULL(PersonName, '') AS PersonName,
                   ISNULL(VehicleNo, '') AS VehicleNo,
                   ISNULL(EmpID, '') AS EmpID,
                   ISNULL(VM_AccNo, '') AS VM_AccNo,
                   ISNULL(VendNameManual, '') AS VendNameManual,
                   ISNULL(Description, '') AS Description
            FROM dbo.MiscEntries_GateOut
            WHERE EntryID = @EntryID;
            """;

        const string detailsSql = """
            SELECT EntryID,
                   MEGO_RefID,
                   ISNULL(Detail, '') AS Detail,
                   ISNULL(Unit, 'Pcs') AS Unit,
                   ISNULL(Qty, 1) AS Qty,
                   ISNULL(Rate, 0) AS Rate
            FROM dbo.MiscEntries_GateOut_Details
            WHERE MEGO_RefID = @EntryID
            ORDER BY EntryID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        GatePassPermanentEntryModel? model = null;

        await using (var command = new SqlCommand(masterSql, connection))
        {
            command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
            await using var reader = await command.ExecuteReaderAsync(cancellationToken);
            if (await reader.ReadAsync(cancellationToken))
            {
                model = new GatePassPermanentEntryModel
                {
                    EntryID = reader.GetInt32(reader.GetOrdinal("EntryID")),
                    GatePassNo = reader.IsDBNull(reader.GetOrdinal("GatePassNo")) ? "" : reader.GetString(reader.GetOrdinal("GatePassNo")),
                    DT = reader.IsDBNull(reader.GetOrdinal("DT")) ? DateTime.Now : reader.GetDateTime(reader.GetOrdinal("DT")),
                    DeptID = reader.GetString(reader.GetOrdinal("DeptID")),
                    GatePassType = reader.GetInt32(reader.GetOrdinal("GatePassType")),
                    PersonName = reader.GetString(reader.GetOrdinal("PersonName")),
                    VehicleNo = reader.GetString(reader.GetOrdinal("VehicleNo")),
                    EmpID = reader.GetString(reader.GetOrdinal("EmpID")),
                    VM_AccNo = reader.GetString(reader.GetOrdinal("VM_AccNo")),
                    VendNameManual = reader.GetString(reader.GetOrdinal("VendNameManual")),
                    Description = reader.GetString(reader.GetOrdinal("Description"))
                };
            }
        }

        if (model == null) return null;

        await using (var detailCommand = new SqlCommand(detailsSql, connection))
        {
            detailCommand.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
            await using var detailReader = await detailCommand.ExecuteReaderAsync(cancellationToken);
            while (await detailReader.ReadAsync(cancellationToken))
            {
                model.Items.Add(new GatePassPermanentDetailItemModel
                {
                    EntryID = detailReader.GetInt32(detailReader.GetOrdinal("EntryID")),
                    MEGO_RefID = detailReader.GetInt32(detailReader.GetOrdinal("MEGO_RefID")),
                    Detail = detailReader.GetString(detailReader.GetOrdinal("Detail")),
                    Unit = detailReader.GetString(detailReader.GetOrdinal("Unit")),
                    Qty = detailReader.GetFloat(detailReader.GetOrdinal("Qty")),
                    Rate = detailReader.GetFloat(detailReader.GetOrdinal("Rate"))
                });
            }
        }

        return model;
    }

    public async Task<int> SaveAsync(
        GatePassPermanentEntryModel model, 
        string userName, 
        string machineName, 
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(model);

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

        try
        {
            int entryId = model.EntryID;

            if (entryId == 0)
            {
                // Generate GatePassNo if not provided
                if (string.IsNullOrWhiteSpace(model.GatePassNo))
                {
                    model.GatePassNo = await GenerateGatePassNoInternalAsync(connection, transaction, model.DT, cancellationToken);
                }

                const string insertSql = """
                    INSERT INTO dbo.MiscEntries_GateOut (
                        DT, GatePassNo, PersonName, VehicleNo, Description,
                        UserName, MachineName, DTEntry, EmpID, GatePassType,
                        DeptID, VM_AccNo, VendNameManual,
                        Approved_HOD, Approved, Approved_Finance
                    ) VALUES (
                        @DT, @GatePassNo, UPPER(@PersonName), UPPER(@VehicleNo), UPPER(@Description),
                        @UserName, @MachineName, GETDATE(), @EmpID, @GatePassType,
                        @DeptID, @VM_AccNo, UPPER(@VendNameManual),
                        0, 0, 0
                    );
                    SELECT CAST(SCOPE_IDENTITY() AS int);
                    """;

                await using var insertCmd = new SqlCommand(insertSql, connection, transaction);
                BindMasterParameters(insertCmd, model, userName, machineName);
                var scalar = await insertCmd.ExecuteScalarAsync(cancellationToken);
                entryId = Convert.ToInt32(scalar);
            }
            else
            {
                const string updateSql = """
                    UPDATE dbo.MiscEntries_GateOut SET
                        DT = @DT,
                        GatePassNo = @GatePassNo,
                        PersonName = UPPER(@PersonName),
                        VehicleNo = UPPER(@VehicleNo),
                        Description = UPPER(@Description),
                        EmpID = @EmpID,
                        GatePassType = @GatePassType,
                        DeptID = @DeptID,
                        VM_AccNo = @VM_AccNo,
                        VendNameManual = UPPER(@VendNameManual)
                    WHERE EntryID = @EntryID;
                    """;

                await using var updateCmd = new SqlCommand(updateSql, connection, transaction);
                updateCmd.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
                BindMasterParameters(updateCmd, model, userName, machineName);
                await updateCmd.ExecuteNonQueryAsync(cancellationToken);

                const string deleteDetailsSql = "DELETE FROM dbo.MiscEntries_GateOut_Details WHERE MEGO_RefID = @RefID;";
                await using var deleteCmd = new SqlCommand(deleteDetailsSql, connection, transaction);
                deleteCmd.Parameters.Add("@RefID", SqlDbType.Int).Value = entryId;
                await deleteCmd.ExecuteNonQueryAsync(cancellationToken);
            }

            // Insert Details
            if (model.Items.Count > 0)
            {
                const string insertItemSql = """
                    INSERT INTO dbo.MiscEntries_GateOut_Details (
                        MEGO_RefID, Detail, Unit, Qty, Rate
                    ) VALUES (
                        @MEGO_RefID, UPPER(@Detail), UPPER(@Unit), @Qty, @Rate
                    );
                    """;

                foreach (var item in model.Items)
                {
                    if (string.IsNullOrWhiteSpace(item.Detail)) continue;

                    await using var itemCmd = new SqlCommand(insertItemSql, connection, transaction);
                    itemCmd.Parameters.Add("@MEGO_RefID", SqlDbType.Int).Value = entryId;
                    itemCmd.Parameters.Add("@Detail", SqlDbType.VarChar, 255).Value = item.Detail.Trim();
                    itemCmd.Parameters.Add("@Unit", SqlDbType.VarChar, 50).Value = (object?)item.Unit?.Trim() ?? "Pcs";
                    itemCmd.Parameters.Add("@Qty", SqlDbType.Real).Value = item.Qty;
                    itemCmd.Parameters.Add("@Rate", SqlDbType.Real).Value = item.Rate;
                    await itemCmd.ExecuteNonQueryAsync(cancellationToken);
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

    private static void BindMasterParameters(SqlCommand command, GatePassPermanentEntryModel model, string userName, string machineName)
    {
        command.Parameters.Add("@DT", SqlDbType.DateTime).Value = model.DT;
        command.Parameters.Add("@GatePassNo", SqlDbType.VarChar, 50).Value = model.GatePassNo.Trim();
        command.Parameters.Add("@PersonName", SqlDbType.VarChar, 255).Value = (object?)model.PersonName?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@VehicleNo", SqlDbType.VarChar, 50).Value = (object?)model.VehicleNo?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@Description", SqlDbType.VarChar, 255).Value = (object?)model.Description?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@UserName", SqlDbType.VarChar, 50).Value = (object?)userName ?? DBNull.Value;
        command.Parameters.Add("@MachineName", SqlDbType.VarChar, 50).Value = (object?)machineName ?? DBNull.Value;
        command.Parameters.Add("@EmpID", SqlDbType.VarChar, 50).Value = (object?)model.EmpID?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@GatePassType", SqlDbType.Int).Value = model.GatePassType;
        command.Parameters.Add("@DeptID", SqlDbType.VarChar, 50).Value = model.DeptID.Trim();
        command.Parameters.Add("@VM_AccNo", SqlDbType.VarChar, 50).Value = (object?)model.VM_AccNo?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@VendNameManual", SqlDbType.VarChar, 50).Value = (object?)model.VendNameManual?.Trim() ?? DBNull.Value;
    }

    public async Task DeleteAsync(int entryId, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

        try
        {
            const string sqlDetails = "DELETE FROM dbo.MiscEntries_GateOut_Details WHERE MEGO_RefID = @EntryID;";
            await using var cmdDetails = new SqlCommand(sqlDetails, connection, transaction);
            cmdDetails.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
            await cmdDetails.ExecuteNonQueryAsync(cancellationToken);

            const string sqlMaster = "DELETE FROM dbo.MiscEntries_GateOut WHERE EntryID = @EntryID;";
            await using var cmdMaster = new SqlCommand(sqlMaster, connection, transaction);
            cmdMaster.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
            await cmdMaster.ExecuteNonQueryAsync(cancellationToken);

            await transaction.CommitAsync(cancellationToken);
        }
        catch
        {
            await transaction.RollbackAsync(cancellationToken);
            throw;
        }
    }

    public async Task<string> GenerateGatePassNoAsync(DateTime dt, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        return await GenerateGatePassNoInternalAsync(connection, null, dt, cancellationToken);
    }

    private static async Task<string> GenerateGatePassNoInternalAsync(
        SqlConnection connection, 
        SqlTransaction? transaction, 
        DateTime dt, 
        CancellationToken cancellationToken)
    {
        var prefix = $"GP-{dt:MMyy}-";

        const string sql = """
            SELECT ISNULL(MAX(CAST(RIGHT(GatePassNo, 4) AS int)), 0)
            FROM dbo.MiscEntries_GateOut
            WHERE GatePassNo LIKE @Prefix + '%'
              AND MONTH(DT) = @Month AND YEAR(DT) = @Year
              AND ISNUMERIC(RIGHT(GatePassNo, 4)) = 1;
            """;

        await using var command = new SqlCommand(sql, connection, transaction);
        command.Parameters.Add("@Prefix", SqlDbType.VarChar, 10).Value = prefix;
        command.Parameters.Add("@Month", SqlDbType.Int).Value = dt.Month;
        command.Parameters.Add("@Year", SqlDbType.Int).Value = dt.Year;

        var result = await command.ExecuteScalarAsync(cancellationToken);
        var maxVal = result != null && result != DBNull.Value ? Convert.ToInt32(result) : 0;
        return $"{prefix}{(maxVal + 1):D4}";
    }

    public async Task SetApprovalAsync(int entryId, string approvalColumn, bool value, CancellationToken cancellationToken = default)
    {
        // Whitelist valid column names to prevent SQL injection
        var safeColumn = approvalColumn switch
        {
            "Approved_HOD" => "Approved_HOD",
            "Approved" => "Approved",
            "Approved_Finance" => "Approved_Finance",
            _ => throw new ArgumentException($"Invalid approval column: {approvalColumn}", nameof(approvalColumn))
        };

        var sql = $"UPDATE dbo.MiscEntries_GateOut SET {safeColumn} = @Value WHERE EntryID = @EntryID;";

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@Value", SqlDbType.Bit).Value = value;
        command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetDepartmentsAsync(CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT DeptID, ISNULL(Name, '') AS Name
            FROM dbo.Departments
            ORDER BY Name;
            """;

        var list = new List<GateLookupOptionModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            list.Add(new GateLookupOptionModel
            {
                Id = reader.GetString(0),
                Text = reader.GetString(1)
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<CourierEmployeeLookupModel>> GetEmployeeOptionsAsync(CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT ISNULL(EmpID, '') AS EmpID, 
                   ISNULL(Name, '') AS Name, 
                   ISNULL(FName, '') AS FName, 
                   ISNULL(Designation, '') AS Designation
            FROM dbo.Employees
            WHERE ISNULL(Active, 1) = 1
            ORDER BY Name;
            """;

        var list = new List<CourierEmployeeLookupModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            list.Add(new CourierEmployeeLookupModel
            {
                EmpID = reader.GetString(0),
                Name = reader.GetString(1),
                FName = reader.GetString(2),
                Designation = reader.GetString(3)
            });
        }
        return list;
    }

    public async Task<IReadOnlyList<CourierVendorLookupModel>> GetVendorOptionsAsync(CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT ISNULL(AccNo, '') AS AccNo, 
                   ISNULL(AccTitle, '') AS AccTitle, 
                   ISNULL(Address, '') AS Address
            FROM dbo.VVendersAndMakers
            WHERE Active = 1
            ORDER BY AccTitle;
            """;

        var list = new List<CourierVendorLookupModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            list.Add(new CourierVendorLookupModel
            {
                AccNo = reader.GetString(0),
                AccTitle = reader.GetString(1),
                Address = reader.GetString(2)
            });
        }
        return list;
    }
}

