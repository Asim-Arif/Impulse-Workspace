using System.Data;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Gate;

public class GatePassReturnableDataAccess : IGatePassReturnableDataAccess
{
    private readonly string _connectionString;

    public GatePassReturnableDataAccess(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection connection string is not configured.");
    }

    private SqlConnection CreateConnection() => new SqlConnection(_connectionString);

    public async Task<IReadOnlyList<GatePassReturnableListRowModel>> GetListAsync(
        GatePassReturnableFilterModel filter, 
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(filter);

        var startBoundary = filter.DtFrom.Date;
        var endBoundary = filter.DtTo.Date.AddDays(1).AddTicks(-1);

        var list = new List<GatePassReturnableListRowModel>();

        var sql = """
            SELECT 
                EntryID, GatePassNo, DT, PersonName, VehicleNo, Description,
                UserName, MachineName, DTEntry, SenderName, Returnable,
                IssdTo, GatePassType, ReturnDT, DeptID_From, DeptID_To,
                DeptName_From, DeptName_To, VendID, VM_AccNo, Name,
                ApprovedBy_Name, ApprovedBy_Designation,
                ISNULL(TotalQty, 0) AS TotalQty,
                ISNULL(TotalQtyRcvd, 0) AS TotalQtyRcvd
            FROM dbo.VGatePass_Return
            WHERE DT >= @DtFrom AND DT <= @DtTo
            """;

        if (!string.IsNullOrWhiteSpace(filter.GatePassType))
        {
            sql += " AND GatePassType = @GatePassType";
        }

        if (string.Equals(filter.ReturnableFilter, "Yes", StringComparison.OrdinalIgnoreCase))
        {
            sql += " AND Returnable = 1";
        }
        else if (string.Equals(filter.ReturnableFilter, "No", StringComparison.OrdinalIgnoreCase))
        {
            sql += " AND Returnable = 0";
        }

        if (!string.IsNullOrWhiteSpace(filter.VM_AccNo))
        {
            sql += " AND VM_AccNo = @VM_AccNo";
        }

        if (!string.IsNullOrWhiteSpace(filter.SearchText))
        {
            sql += """
                 AND (
                    GatePassNo LIKE @Search OR
                    IssdTo LIKE @Search OR
                    PersonName LIKE @Search OR
                    SenderName LIKE @Search OR
                    Name LIKE @Search OR
                    DeptName_From LIKE @Search OR
                    DeptName_To LIKE @Search OR
                    ApprovedBy_Name LIKE @Search
                )
                """;
        }

        sql += " ORDER BY DT DESC, EntryID DESC;";

        await using var connection = CreateConnection();
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@DtFrom", SqlDbType.DateTime).Value = startBoundary;
        command.Parameters.Add("@DtTo", SqlDbType.DateTime).Value = endBoundary;

        if (!string.IsNullOrWhiteSpace(filter.GatePassType))
        {
            command.Parameters.Add("@GatePassType", SqlDbType.VarChar, 50).Value = filter.GatePassType.Trim();
        }

        if (!string.IsNullOrWhiteSpace(filter.VM_AccNo))
        {
            command.Parameters.Add("@VM_AccNo", SqlDbType.VarChar, 50).Value = filter.VM_AccNo.Trim();
        }

        if (!string.IsNullOrWhiteSpace(filter.SearchText))
        {
            command.Parameters.Add("@Search", SqlDbType.VarChar, 255).Value = $"%{filter.SearchText.Trim()}%";
        }

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            list.Add(new GatePassReturnableListRowModel
            {
                EntryID = reader.GetInt32(reader.GetOrdinal("EntryID")),
                GatePassNo = reader.IsDBNull(reader.GetOrdinal("GatePassNo")) ? "" : reader.GetString(reader.GetOrdinal("GatePassNo")),
                DT = reader.IsDBNull(reader.GetOrdinal("DT")) ? null : reader.GetDateTime(reader.GetOrdinal("DT")),
                PersonName = reader.IsDBNull(reader.GetOrdinal("PersonName")) ? null : reader.GetString(reader.GetOrdinal("PersonName")),
                VehicleNo = reader.IsDBNull(reader.GetOrdinal("VehicleNo")) ? null : reader.GetString(reader.GetOrdinal("VehicleNo")),
                Description = reader.IsDBNull(reader.GetOrdinal("Description")) ? null : reader.GetString(reader.GetOrdinal("Description")),
                UserName = reader.IsDBNull(reader.GetOrdinal("UserName")) ? null : reader.GetString(reader.GetOrdinal("UserName")),
                MachineName = reader.IsDBNull(reader.GetOrdinal("MachineName")) ? null : reader.GetString(reader.GetOrdinal("MachineName")),
                DTEntry = reader.IsDBNull(reader.GetOrdinal("DTEntry")) ? null : reader.GetDateTime(reader.GetOrdinal("DTEntry")),
                SenderName = reader.IsDBNull(reader.GetOrdinal("SenderName")) ? null : reader.GetString(reader.GetOrdinal("SenderName")),
                Returnable = !reader.IsDBNull(reader.GetOrdinal("Returnable")) && reader.GetBoolean(reader.GetOrdinal("Returnable")),
                IssdTo = reader.IsDBNull(reader.GetOrdinal("IssdTo")) ? null : reader.GetString(reader.GetOrdinal("IssdTo")),
                GatePassType = reader.IsDBNull(reader.GetOrdinal("GatePassType")) ? null : reader.GetString(reader.GetOrdinal("GatePassType")),
                ReturnDT = reader.IsDBNull(reader.GetOrdinal("ReturnDT")) ? null : reader.GetDateTime(reader.GetOrdinal("ReturnDT")),
                DeptID_From = reader.IsDBNull(reader.GetOrdinal("DeptID_From")) ? null : reader.GetString(reader.GetOrdinal("DeptID_From")),
                DeptID_To = reader.IsDBNull(reader.GetOrdinal("DeptID_To")) ? null : reader.GetString(reader.GetOrdinal("DeptID_To")),
                DeptName_From = reader.IsDBNull(reader.GetOrdinal("DeptName_From")) ? null : reader.GetString(reader.GetOrdinal("DeptName_From")),
                DeptName_To = reader.IsDBNull(reader.GetOrdinal("DeptName_To")) ? null : reader.GetString(reader.GetOrdinal("DeptName_To")),
                VendID = reader.IsDBNull(reader.GetOrdinal("VendID")) ? null : reader.GetInt32(reader.GetOrdinal("VendID")),
                VM_AccNo = reader.IsDBNull(reader.GetOrdinal("VM_AccNo")) ? null : reader.GetString(reader.GetOrdinal("VM_AccNo")),
                Name = reader.IsDBNull(reader.GetOrdinal("Name")) ? null : reader.GetString(reader.GetOrdinal("Name")),
                ApprovedBy_Name = reader.IsDBNull(reader.GetOrdinal("ApprovedBy_Name")) ? null : reader.GetString(reader.GetOrdinal("ApprovedBy_Name")),
                ApprovedBy_Designation = reader.IsDBNull(reader.GetOrdinal("ApprovedBy_Designation")) ? null : reader.GetString(reader.GetOrdinal("ApprovedBy_Designation")),
                TotalQty = Convert.ToDouble(reader.GetValue(reader.GetOrdinal("TotalQty"))),
                TotalQtyRcvd = Convert.ToDouble(reader.GetValue(reader.GetOrdinal("TotalQtyRcvd")))
            });
        }

        return list;
    }

    public async Task<GatePassReturnableEntryModel?> GetByIdAsync(int entryId, CancellationToken cancellationToken = default)
    {
        const string masterSql = """
            SELECT 
                EntryID, GatePassNo, DT, PersonName, SenderName, Returnable,
                IssdTo, GatePassType, ReturnDT, DeptID_From, DeptID_To,
                VM_AccNo, ApprovedBy_Name, ApprovedBy_Designation
            FROM dbo.GatePass_Return
            WHERE EntryID = @EntryID;
            """;

        const string detailsSql = """
            SELECT EntryID, MEGO_RefID, Detail, Unit, Packing, Qty
            FROM dbo.GatePass_Return_Details
            WHERE MEGO_RefID = @EntryID
            ORDER BY EntryID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        GatePassReturnableEntryModel? model = null;

        await using (var masterCommand = new SqlCommand(masterSql, connection))
        {
            masterCommand.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
            await using var reader = await masterCommand.ExecuteReaderAsync(cancellationToken);
            if (await reader.ReadAsync(cancellationToken))
            {
                model = new GatePassReturnableEntryModel
                {
                    EntryID = reader.GetInt32(reader.GetOrdinal("EntryID")),
                    GatePassNo = reader.IsDBNull(reader.GetOrdinal("GatePassNo")) ? "" : reader.GetString(reader.GetOrdinal("GatePassNo")),
                    DT = reader.IsDBNull(reader.GetOrdinal("DT")) ? DateTime.Now : reader.GetDateTime(reader.GetOrdinal("DT")),
                    PersonName = reader.IsDBNull(reader.GetOrdinal("PersonName")) ? null : reader.GetString(reader.GetOrdinal("PersonName")),
                    SenderName = reader.IsDBNull(reader.GetOrdinal("SenderName")) ? null : reader.GetString(reader.GetOrdinal("SenderName")),
                    Returnable = !reader.IsDBNull(reader.GetOrdinal("Returnable")) && reader.GetBoolean(reader.GetOrdinal("Returnable")),
                    IssdTo = reader.IsDBNull(reader.GetOrdinal("IssdTo")) ? null : reader.GetString(reader.GetOrdinal("IssdTo")),
                    GatePassType = reader.IsDBNull(reader.GetOrdinal("GatePassType")) ? "" : reader.GetString(reader.GetOrdinal("GatePassType")),
                    ReturnDT = reader.IsDBNull(reader.GetOrdinal("ReturnDT")) ? null : reader.GetDateTime(reader.GetOrdinal("ReturnDT")),
                    DeptID_From = reader.IsDBNull(reader.GetOrdinal("DeptID_From")) ? "" : reader.GetString(reader.GetOrdinal("DeptID_From")),
                    DeptID_To = reader.IsDBNull(reader.GetOrdinal("DeptID_To")) ? null : reader.GetString(reader.GetOrdinal("DeptID_To")),
                    VM_AccNo = reader.IsDBNull(reader.GetOrdinal("VM_AccNo")) ? null : reader.GetString(reader.GetOrdinal("VM_AccNo")),
                    ApprovedBy_Name = reader.IsDBNull(reader.GetOrdinal("ApprovedBy_Name")) ? null : reader.GetString(reader.GetOrdinal("ApprovedBy_Name")),
                    ApprovedBy_Designation = reader.IsDBNull(reader.GetOrdinal("ApprovedBy_Designation")) ? null : reader.GetString(reader.GetOrdinal("ApprovedBy_Designation"))
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
                model.Items.Add(new GatePassReturnableDetailItemModel
                {
                    EntryID = detailReader.GetInt32(detailReader.GetOrdinal("EntryID")),
                    MEGO_RefID = detailReader.GetInt32(detailReader.GetOrdinal("MEGO_RefID")),
                    Detail = detailReader.GetString(detailReader.GetOrdinal("Detail")),
                    Unit = detailReader.GetString(detailReader.GetOrdinal("Unit")),
                    Packing = detailReader.IsDBNull(detailReader.GetOrdinal("Packing")) ? null : detailReader.GetString(detailReader.GetOrdinal("Packing")),
                    Qty = Convert.ToDouble(detailReader.GetValue(detailReader.GetOrdinal("Qty")))
                });
            }
        }

        return model;
    }

    public async Task<GatePassReceiveModel?> GetReceiveModelByIdAsync(int entryId, CancellationToken cancellationToken = default)
    {
        const string masterSql = """
            SELECT EntryID, GatePassNo, DT, IssdTo, Name
            FROM dbo.VGatePass_Return
            WHERE EntryID = @EntryID;
            """;

        const string detailsSql = """
            SELECT EntryID, Detail, Unit, Packing, Qty, TotalQtyRcvd
            FROM dbo.VGatePass_Return_Details
            WHERE MEGO_RefID = @EntryID
            ORDER BY EntryID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        GatePassReceiveModel? model = null;

        await using (var masterCommand = new SqlCommand(masterSql, connection))
        {
            masterCommand.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
            await using var reader = await masterCommand.ExecuteReaderAsync(cancellationToken);
            if (await reader.ReadAsync(cancellationToken))
            {
                model = new GatePassReceiveModel
                {
                    RefEntryID = reader.GetInt32(reader.GetOrdinal("EntryID")),
                    GatePassNo = reader.IsDBNull(reader.GetOrdinal("GatePassNo")) ? "" : reader.GetString(reader.GetOrdinal("GatePassNo")),
                    DT = reader.IsDBNull(reader.GetOrdinal("DT")) ? DateTime.Now : reader.GetDateTime(reader.GetOrdinal("DT")),
                    IssdTo = reader.IsDBNull(reader.GetOrdinal("IssdTo")) ? null : reader.GetString(reader.GetOrdinal("IssdTo")),
                    VendorName = reader.IsDBNull(reader.GetOrdinal("Name")) ? null : reader.GetString(reader.GetOrdinal("Name")),
                    RcvDT = DateTime.Today
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
                var issuedQty = Convert.ToDouble(detailReader.GetValue(detailReader.GetOrdinal("Qty")));
                var rcvdQty = Convert.ToDouble(detailReader.GetValue(detailReader.GetOrdinal("TotalQtyRcvd")));
                var balance = Math.Max(0, issuedQty - rcvdQty);

                model.Items.Add(new GatePassReceiveDetailItemModel
                {
                    DetailRefID = detailReader.GetInt32(detailReader.GetOrdinal("EntryID")),
                    Detail = detailReader.GetString(detailReader.GetOrdinal("Detail")),
                    Unit = detailReader.GetString(detailReader.GetOrdinal("Unit")),
                    Packing = detailReader.IsDBNull(detailReader.GetOrdinal("Packing")) ? null : detailReader.GetString(detailReader.GetOrdinal("Packing")),
                    IssuedQty = issuedQty,
                    PreviouslyReceivedQty = rcvdQty,
                    QtyReceivedNow = balance // Default to balance quantity for convenient 1-click full receiving
                });
            }
        }

        return model;
    }

    public async Task<int> SaveAsync(
        GatePassReturnableEntryModel model, 
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
                if (string.IsNullOrWhiteSpace(model.GatePassNo))
                {
                    model.GatePassNo = await GenerateGatePassNoInternalAsync(connection, transaction, model.DT, cancellationToken);
                }

                const string insertSql = """
                    INSERT INTO dbo.GatePass_Return (
                        DT, GatePassNo, IssdTo, PersonName, SenderName,
                        UserName, MachineName, DTEntry, Returnable, GatePassType,
                        ReturnDT, DeptID_From, DeptID_To, ApprovedBy_Name,
                        ApprovedBy_Designation, VendID, VM_AccNo
                    ) VALUES (
                        @DT, @GatePassNo, UPPER(@IssdTo), UPPER(@PersonName), UPPER(@SenderName),
                        @UserName, @MachineName, GETDATE(), @Returnable, UPPER(@GatePassType),
                        @ReturnDT, @DeptID_From, @DeptID_To, @ApprovedBy_Name,
                        UPPER(@ApprovedBy_Designation), 0, @VM_AccNo
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
                    UPDATE dbo.GatePass_Return SET
                        DT = @DT,
                        GatePassNo = @GatePassNo,
                        IssdTo = UPPER(@IssdTo),
                        PersonName = UPPER(@PersonName),
                        SenderName = UPPER(@SenderName),
                        Returnable = @Returnable,
                        GatePassType = UPPER(@GatePassType),
                        ReturnDT = @ReturnDT,
                        DeptID_From = @DeptID_From,
                        DeptID_To = @DeptID_To,
                        ApprovedBy_Name = @ApprovedBy_Name,
                        ApprovedBy_Designation = UPPER(@ApprovedBy_Designation),
                        VM_AccNo = @VM_AccNo
                    WHERE EntryID = @EntryID;
                    """;

                await using var updateCmd = new SqlCommand(updateSql, connection, transaction);
                updateCmd.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
                BindMasterParameters(updateCmd, model, userName, machineName);
                await updateCmd.ExecuteNonQueryAsync(cancellationToken);

                const string deleteDetailsSql = "DELETE FROM dbo.GatePass_Return_Details WHERE MEGO_RefID = @RefID;";
                await using var deleteCmd = new SqlCommand(deleteDetailsSql, connection, transaction);
                deleteCmd.Parameters.Add("@RefID", SqlDbType.Int).Value = entryId;
                await deleteCmd.ExecuteNonQueryAsync(cancellationToken);
            }

            // Insert Details
            if (model.Items.Count > 0)
            {
                const string insertItemSql = """
                    INSERT INTO dbo.GatePass_Return_Details (
                        MEGO_RefID, Detail, Unit, Packing, Qty
                    ) VALUES (
                        @MEGO_RefID, UPPER(@Detail), UPPER(@Unit), UPPER(@Packing), @Qty
                    );
                    """;

                foreach (var item in model.Items)
                {
                    if (string.IsNullOrWhiteSpace(item.Detail)) continue;

                    await using var itemCmd = new SqlCommand(insertItemSql, connection, transaction);
                    itemCmd.Parameters.Add("@MEGO_RefID", SqlDbType.Int).Value = entryId;
                    itemCmd.Parameters.Add("@Detail", SqlDbType.VarChar, 255).Value = item.Detail.Trim();
                    itemCmd.Parameters.Add("@Unit", SqlDbType.VarChar, 50).Value = (object?)item.Unit?.Trim() ?? "Pcs";
                    itemCmd.Parameters.Add("@Packing", SqlDbType.VarChar, 50).Value = (object?)item.Packing?.Trim() ?? DBNull.Value;
                    itemCmd.Parameters.Add("@Qty", SqlDbType.Int).Value = Convert.ToInt32(Math.Round(item.Qty));
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

    public async Task<int> SaveReceivingAsync(
        GatePassReceiveModel model, 
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
            const string insertRcvSql = """
                INSERT INTO dbo.GatePass_Return_Receiving (
                    ME_GO_RefID, UserName, MachineName, DTEntry, RcvDT
                ) VALUES (
                    @RefID, @UserName, @MachineName, GETDATE(), @RcvDT
                );
                SELECT CAST(SCOPE_IDENTITY() AS int);
                """;

            await using var rcvCmd = new SqlCommand(insertRcvSql, connection, transaction);
            rcvCmd.Parameters.Add("@RefID", SqlDbType.Int).Value = model.RefEntryID;
            rcvCmd.Parameters.Add("@UserName", SqlDbType.VarChar, 50).Value = userName;
            rcvCmd.Parameters.Add("@MachineName", SqlDbType.VarChar, 50).Value = machineName;
            rcvCmd.Parameters.Add("@RcvDT", SqlDbType.DateTime).Value = model.RcvDT.Date;

            var rcvScalar = await rcvCmd.ExecuteScalarAsync(cancellationToken);
            var rcvId = Convert.ToInt32(rcvScalar);

            const string insertDetailSql = """
                INSERT INTO dbo.GatePass_Return_Receiving_Detail (
                    ME_GO_Rcv_RefID, ME_GO_D_RefID, QtyRcvd
                ) VALUES (
                    @RcvRefID, @DRefID, @QtyRcvd
                );
                """;

            foreach (var item in model.Items)
            {
                if (item.QtyReceivedNow <= 0) continue;

                await using var itemCmd = new SqlCommand(insertDetailSql, connection, transaction);
                itemCmd.Parameters.Add("@RcvRefID", SqlDbType.Int).Value = rcvId;
                itemCmd.Parameters.Add("@DRefID", SqlDbType.Int).Value = item.DetailRefID;
                itemCmd.Parameters.Add("@QtyRcvd", SqlDbType.Int).Value = Convert.ToInt32(Math.Round(item.QtyReceivedNow));
                await itemCmd.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
            return rcvId;
        }
        catch
        {
            await transaction.RollbackAsync(cancellationToken);
            throw;
        }
    }

    public async Task<string> GenerateGatePassNoAsync(DateTime date, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        return await GenerateGatePassNoInternalAsync(connection, null, date, cancellationToken);
    }

    private static async Task<string> GenerateGatePassNoInternalAsync(
        SqlConnection connection, 
        SqlTransaction? transaction, 
        DateTime date, 
        CancellationToken cancellationToken)
    {
        // Daily sequence: GO-ddMMyy-###
        const string sql = """
            SELECT ISNULL(MAX(CAST(RIGHT(GatePassNo, 3) AS INT)), 0)
            FROM dbo.GatePass_Return
            WHERE YEAR(DT) = @Year AND MONTH(DT) = @Month AND DAY(DT) = @Day;
            """;

        await using var command = new SqlCommand(sql, connection, transaction);
        command.Parameters.Add("@Year", SqlDbType.Int).Value = date.Year;
        command.Parameters.Add("@Month", SqlDbType.Int).Value = date.Month;
        command.Parameters.Add("@Day", SqlDbType.Int).Value = date.Day;

        var scalar = await command.ExecuteScalarAsync(cancellationToken);
        var currentSeq = scalar != null && scalar != DBNull.Value ? Convert.ToInt32(scalar) : 0;
        var nextSeq = currentSeq + 1;

        return $"GO-{date:ddMMyy}-{nextSeq:D3}";
    }

    private static void BindMasterParameters(
        SqlCommand cmd, 
        GatePassReturnableEntryModel model, 
        string userName, 
        string machineName)
    {
        cmd.Parameters.Add("@DT", SqlDbType.DateTime).Value = model.DT;
        cmd.Parameters.Add("@GatePassNo", SqlDbType.VarChar, 50).Value = model.GatePassNo.Trim();
        cmd.Parameters.Add("@IssdTo", SqlDbType.VarChar, 255).Value = (object?)model.IssdTo?.Trim() ?? DBNull.Value;
        cmd.Parameters.Add("@PersonName", SqlDbType.VarChar, 255).Value = (object?)model.PersonName?.Trim() ?? DBNull.Value;
        cmd.Parameters.Add("@SenderName", SqlDbType.VarChar, 255).Value = (object?)model.SenderName?.Trim() ?? DBNull.Value;
        cmd.Parameters.Add("@Returnable", SqlDbType.Bit).Value = model.Returnable;
        cmd.Parameters.Add("@GatePassType", SqlDbType.VarChar, 50).Value = (object?)model.GatePassType?.Trim() ?? DBNull.Value;
        cmd.Parameters.Add("@ReturnDT", SqlDbType.DateTime).Value = model.Returnable && model.ReturnDT.HasValue ? model.ReturnDT.Value : DBNull.Value;
        cmd.Parameters.Add("@DeptID_From", SqlDbType.VarChar, 50).Value = (object?)model.DeptID_From?.Trim() ?? DBNull.Value;
        cmd.Parameters.Add("@DeptID_To", SqlDbType.VarChar, 50).Value = (object?)model.DeptID_To?.Trim() ?? DBNull.Value;
        cmd.Parameters.Add("@ApprovedBy_Name", SqlDbType.VarChar, 50).Value = (object?)model.ApprovedBy_Name?.Trim() ?? DBNull.Value;
        cmd.Parameters.Add("@ApprovedBy_Designation", SqlDbType.VarChar, 50).Value = (object?)model.ApprovedBy_Designation?.Trim() ?? DBNull.Value;
        cmd.Parameters.Add("@VM_AccNo", SqlDbType.VarChar, 50).Value = (object?)model.VM_AccNo?.Trim() ?? DBNull.Value;
        cmd.Parameters.Add("@UserName", SqlDbType.VarChar, 50).Value = userName;
        cmd.Parameters.Add("@MachineName", SqlDbType.VarChar, 50).Value = machineName;
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetDepartmentsAsync(CancellationToken cancellationToken = default)
    {
        const string sql = "SELECT DeptID, Name FROM dbo.Departments ORDER BY Name;";
        var results = new List<GateLookupOptionModel>();

        await using var connection = CreateConnection();
        await using var command = new SqlCommand(sql, connection);
        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            results.Add(new GateLookupOptionModel
            {
                Id = reader.GetString(0),
                Text = reader.GetString(1)
            });
        }
        return results;
    }

    public async Task<IReadOnlyList<GateLookupOptionModel>> GetGatePassTypesAsync(CancellationToken cancellationToken = default)
    {
        const string sql = "SELECT DISTINCT GatepassType FROM dbo.GatepassTypes WHERE ISNULL(Active, 1) = 1 ORDER BY GatepassType;";
        var results = new List<GateLookupOptionModel>();

        await using var connection = CreateConnection();
        await using var command = new SqlCommand(sql, connection);
        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            var val = reader.GetString(0);
            results.Add(new GateLookupOptionModel
            {
                Id = val,
                Text = val
            });
        }
        return results;
    }

    public async Task<IReadOnlyList<CourierEmployeeLookupModel>> GetEmployeesAsync(CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT EmpID, Name, ISNULL(FName, '') as FName, ISNULL(Designation, '') as Designation
            FROM dbo.Employees
            WHERE ISNULL(Active, 1) = 1
            ORDER BY Name;
            """;
        var results = new List<CourierEmployeeLookupModel>();

        await using var connection = CreateConnection();
        await using var command = new SqlCommand(sql, connection);
        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            results.Add(new CourierEmployeeLookupModel
            {
                EmpID = reader.GetString(0),
                Name = reader.GetString(1),
                FName = reader.GetString(2),
                Designation = reader.GetString(3)
            });
        }
        return results;
    }

    public async Task<IReadOnlyList<CourierVendorLookupModel>> GetVendorsAsync(CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT AccNo, AccTitle, ISNULL(Address, '') as Address
            FROM dbo.VVendersAndMakers
            ORDER BY AccTitle;
            """;
        var results = new List<CourierVendorLookupModel>();

        await using var connection = CreateConnection();
        await using var command = new SqlCommand(sql, connection);
        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            results.Add(new CourierVendorLookupModel
            {
                AccNo = reader.GetString(0),
                AccTitle = reader.GetString(1),
                Address = reader.GetString(2)
            });
        }
        return results;
    }
}

