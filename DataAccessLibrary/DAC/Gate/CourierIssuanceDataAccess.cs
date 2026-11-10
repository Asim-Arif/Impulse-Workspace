using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Gate;
using DataAccessLibrary.Models.ViewModels.Gate;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataAccessLibrary.DAC.Gate;

public sealed class CourierIssuanceDataAccess : ICourierIssuanceDataAccess
{
    private readonly string _connectionString;

    public CourierIssuanceDataAccess(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection is not configured.");
    }

    private SqlConnection CreateConnection() => new(_connectionString);

    public async Task<IReadOnlyList<CourierIssuanceListRowModel>> GetCourierIssuanceListAsync(
        CourierIssuanceFilterModel filter, 
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(filter);

        const string sql = """
            SELECT EntryID,
                   DT,
                   ISNULL(TrackingNo, '') AS TrackingNo,
                   ISNULL(Description, '') AS Description,
                   ISNULL(VendorName, '') AS VendorName,
                   ISNULL(VendorAddress, '') AS VendorAddress,
                   ISNULL(IssuedBy, '') AS IssuedBy,
                   ISNULL(ConcernedPerson, '') AS ConcernedPerson,
                   ISNULL(ReceivedBy, '') AS ReceivedBy,
                   ISNULL(Self, 1) AS Self,
                   ISNULL(Paid, 0) AS Paid,
                   VendID,
                   ISNULL(Weight, '') AS Weight,
                   ISNULL(Dimension, '') AS Dimension,
                   ISNULL(CourierStatus, 0) AS CourierStatus,
                   ISNULL(MachineName, '') AS MachineName,
                   ISNULL(UserName, '') AS UserName,
                   Parcel_Qty,
                   ISNULL(Vehicle_No, '') AS Vehicle_No
            FROM dbo.Courier_Issuance
            WHERE DT >= @FromDate AND DT < @ToDate
              AND (@VendID IS NULL OR VendID = @VendID)
              AND (
                  @SearchText IS NULL 
                  OR TrackingNo LIKE @SearchText 
                  OR Description LIKE @SearchText 
                  OR VendorName LIKE @SearchText 
                  OR Vehicle_No LIKE @SearchText
                  OR ConcernedPerson LIKE @SearchText
              )
            ORDER BY DT DESC, EntryID DESC;
            """;

        var list = new List<CourierIssuanceListRowModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);

        command.Parameters.Add("@FromDate", SqlDbType.DateTime).Value = filter.FromDate.Date;
        command.Parameters.Add("@ToDate", SqlDbType.DateTime).Value = filter.ToDate.Date.AddDays(1);
        command.Parameters.Add("@VendID", SqlDbType.VarChar, 50).Value = !string.IsNullOrWhiteSpace(filter.VendID) ? filter.VendID.Trim() : DBNull.Value;

        var search = string.IsNullOrWhiteSpace(filter.SearchText) ? null : $"%{filter.SearchText.Trim()}%";
        command.Parameters.Add("@SearchText", SqlDbType.VarChar, 100).Value = (object?)search ?? DBNull.Value;

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            list.Add(new CourierIssuanceListRowModel
            {
                EntryID = reader.GetInt32(reader.GetOrdinal("EntryID")),
                DT = reader.IsDBNull(reader.GetOrdinal("DT")) ? null : reader.GetDateTime(reader.GetOrdinal("DT")),
                TrackingNo = reader.GetString(reader.GetOrdinal("TrackingNo")),
                Description = reader.GetString(reader.GetOrdinal("Description")),
                VendorName = reader.GetString(reader.GetOrdinal("VendorName")),
                VendorAddress = reader.GetString(reader.GetOrdinal("VendorAddress")),
                IssuedBy = reader.GetString(reader.GetOrdinal("IssuedBy")),
                ConcernedPerson = reader.GetString(reader.GetOrdinal("ConcernedPerson")),
                ReceivedBy = reader.GetString(reader.GetOrdinal("ReceivedBy")),
                Self = reader.GetBoolean(reader.GetOrdinal("Self")),
                Paid = reader.GetBoolean(reader.GetOrdinal("Paid")),
                VendID = reader.IsDBNull(reader.GetOrdinal("VendID")) ? null : reader.GetString(reader.GetOrdinal("VendID")),
                Weight = reader.GetString(reader.GetOrdinal("Weight")),
                Dimension = reader.GetString(reader.GetOrdinal("Dimension")),
                CourierStatus = reader.GetBoolean(reader.GetOrdinal("CourierStatus")),
                MachineName = reader.GetString(reader.GetOrdinal("MachineName")),
                UserName = reader.GetString(reader.GetOrdinal("UserName")),
                Parcel_Qty = reader.IsDBNull(reader.GetOrdinal("Parcel_Qty")) ? null : reader.GetFloat(reader.GetOrdinal("Parcel_Qty")),
                Vehicle_No = reader.GetString(reader.GetOrdinal("Vehicle_No"))
            });
        }

        return list;
    }

    public async Task<CourierIssuanceEntryModel?> GetCourierIssuanceAsync(int entryId, CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT EntryID,
                   DT,
                   ISNULL(TrackingNo, '') AS TrackingNo,
                   ISNULL(Description, '') AS Description,
                   ISNULL(VendorName, '') AS VendorName,
                   ISNULL(VendorAddress, '') AS VendorAddress,
                   ISNULL(IssuedBy, '') AS IssuedBy,
                   ISNULL(ConcernedPerson, '') AS ConcernedPerson,
                   ISNULL(ReceivedBy, '') AS ReceivedBy,
                   ISNULL(Self, 1) AS Self,
                   ISNULL(Paid, 0) AS Paid,
                   VendID,
                   ISNULL(Weight, '') AS Weight,
                   ISNULL(Dimension, '') AS Dimension,
                   Parcel_Qty,
                   ISNULL(Vehicle_No, '') AS Vehicle_No
            FROM dbo.Courier_Issuance
            WHERE EntryID = @EntryID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (await reader.ReadAsync(cancellationToken))
        {
            var isSelf = reader.GetBoolean(reader.GetOrdinal("Self"));
            var isPaid = reader.GetBoolean(reader.GetOrdinal("Paid"));

            return new CourierIssuanceEntryModel
            {
                EntryID = reader.GetInt32(reader.GetOrdinal("EntryID")),
                DT = reader.IsDBNull(reader.GetOrdinal("DT")) ? DateTime.Now : reader.GetDateTime(reader.GetOrdinal("DT")),
                TrackingNo = reader.GetString(reader.GetOrdinal("TrackingNo")),
                Description = reader.GetString(reader.GetOrdinal("Description")),
                VendorName = reader.GetString(reader.GetOrdinal("VendorName")),
                VendorAddress = reader.GetString(reader.GetOrdinal("VendorAddress")),
                IssuedBy = reader.GetString(reader.GetOrdinal("IssuedBy")),
                ConcernedPerson = reader.GetString(reader.GetOrdinal("ConcernedPerson")),
                ReceivedBy = reader.GetString(reader.GetOrdinal("ReceivedBy")),
                PaymentMode = isPaid && !isSelf ? "Paid" : "Self",
                VendID = reader.IsDBNull(reader.GetOrdinal("VendID")) ? null : reader.GetString(reader.GetOrdinal("VendID")),
                Weight = reader.GetString(reader.GetOrdinal("Weight")),
                Dimension = reader.GetString(reader.GetOrdinal("Dimension")),
                Parcel_Qty = reader.IsDBNull(reader.GetOrdinal("Parcel_Qty")) ? null : reader.GetFloat(reader.GetOrdinal("Parcel_Qty")),
                Vehicle_No = reader.GetString(reader.GetOrdinal("Vehicle_No"))
            };
        }

        return null;
    }

    public async Task<int> SaveCourierIssuanceAsync(
        CourierIssuanceEntryModel model, 
        string userName, 
        string machineName, 
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(model);

        var courierStatus = !string.IsNullOrWhiteSpace(model.TrackingNo);
        var isSelf = model.Self;
        var isPaid = model.Paid;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        if (model.EntryID == 0)
        {
            const string insertSql = """
                INSERT INTO dbo.Courier_Issuance (
                    DT, TrackingNo, Description, VendorName, VendorAddress,
                    IssuedBy, ConcernedPerson, ReceivedBy, Self, Paid,
                    VendID, Weight, Dimension, CourierStatus, MachineName,
                    UserName, Parcel_Qty, Vehicle_No
                ) VALUES (
                    @DT, UPPER(@TrackingNo), UPPER(@Description), UPPER(@VendorName), UPPER(@VendorAddress),
                    @IssuedBy, UPPER(@ConcernedPerson), UPPER(@ReceivedBy), @Self, @Paid,
                    @VendID, UPPER(@Weight), UPPER(@Dimension), @CourierStatus, @MachineName,
                    @UserName, @Parcel_Qty, UPPER(@Vehicle_No)
                );
                SELECT CAST(SCOPE_IDENTITY() AS int);
                """;

            await using var command = new SqlCommand(insertSql, connection);
            BindParameters(command, model, userName, machineName, courierStatus, isSelf, isPaid);

            var result = await command.ExecuteScalarAsync(cancellationToken);
            return Convert.ToInt32(result);
        }
        else
        {
            const string updateSql = """
                UPDATE dbo.Courier_Issuance SET
                    DT = @DT,
                    TrackingNo = UPPER(@TrackingNo),
                    Description = UPPER(@Description),
                    VendorName = UPPER(@VendorName),
                    VendorAddress = UPPER(@VendorAddress),
                    IssuedBy = @IssuedBy,
                    ConcernedPerson = UPPER(@ConcernedPerson),
                    ReceivedBy = UPPER(@ReceivedBy),
                    Self = @Self,
                    Paid = @Paid,
                    VendID = @VendID,
                    Weight = UPPER(@Weight),
                    Dimension = UPPER(@Dimension),
                    CourierStatus = @CourierStatus,
                    MachineName = @MachineName,
                    UserName = @UserName,
                    Parcel_Qty = @Parcel_Qty,
                    Vehicle_No = UPPER(@Vehicle_No)
                WHERE EntryID = @EntryID;
                """;

            await using var command = new SqlCommand(updateSql, connection);
            command.Parameters.Add("@EntryID", SqlDbType.Int).Value = model.EntryID;
            BindParameters(command, model, userName, machineName, courierStatus, isSelf, isPaid);

            await command.ExecuteNonQueryAsync(cancellationToken);
            return model.EntryID;
        }
    }

    private static void BindParameters(
        SqlCommand command, 
        CourierIssuanceEntryModel model, 
        string userName, 
        string machineName, 
        bool courierStatus, 
        bool isSelf, 
        bool isPaid)
    {
        command.Parameters.Add("@DT", SqlDbType.DateTime).Value = model.DT;
        command.Parameters.Add("@TrackingNo", SqlDbType.VarChar, 50).Value = (object?)model.TrackingNo?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@Description", SqlDbType.VarChar, 255).Value = (object?)model.Description?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@VendorName", SqlDbType.VarChar, 50).Value = (object?)model.VendorName?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@VendorAddress", SqlDbType.VarChar, 255).Value = (object?)model.VendorAddress?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@IssuedBy", SqlDbType.VarChar, 50).Value = (object?)model.IssuedBy?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@ConcernedPerson", SqlDbType.VarChar, 50).Value = (object?)model.ConcernedPerson?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@ReceivedBy", SqlDbType.VarChar, 50).Value = (object?)model.ReceivedBy?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@Self", SqlDbType.Bit).Value = isSelf;
        command.Parameters.Add("@Paid", SqlDbType.Bit).Value = isPaid;
        command.Parameters.Add("@VendID", SqlDbType.VarChar, 50).Value = (object?)model.VendID?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@Weight", SqlDbType.VarChar, 50).Value = (object?)model.Weight?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@Dimension", SqlDbType.VarChar, 50).Value = (object?)model.Dimension?.Trim() ?? DBNull.Value;
        command.Parameters.Add("@CourierStatus", SqlDbType.Bit).Value = courierStatus;
        command.Parameters.Add("@MachineName", SqlDbType.VarChar, 50).Value = (object?)machineName ?? DBNull.Value;
        command.Parameters.Add("@UserName", SqlDbType.VarChar, 50).Value = (object?)userName ?? DBNull.Value;
        command.Parameters.Add("@Parcel_Qty", SqlDbType.Real).Value = model.Parcel_Qty.HasValue ? model.Parcel_Qty.Value : DBNull.Value;
        command.Parameters.Add("@Vehicle_No", SqlDbType.VarChar, 50).Value = (object?)model.Vehicle_No?.Trim() ?? DBNull.Value;
    }

    public async Task DeleteCourierIssuanceAsync(int entryId, CancellationToken cancellationToken = default)
    {
        const string sql = "DELETE FROM dbo.Courier_Issuance WHERE EntryID = @EntryID;";
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        await command.ExecuteNonQueryAsync(cancellationToken);
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

    public async Task<(string? Name, string? Address)> GetVendorDetailsAsync(string accNo, CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT ISNULL(AccTitle, '') AS AccTitle, ISNULL(Address, '') AS Address
            FROM dbo.VVendersAndMakers
            WHERE AccNo = @AccNo;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@AccNo", SqlDbType.VarChar, 50).Value = accNo;

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (await reader.ReadAsync(cancellationToken))
        {
            return (reader.GetString(0), reader.GetString(1));
        }

        return (null, null);
    }
}

