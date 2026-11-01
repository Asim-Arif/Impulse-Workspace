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

public sealed class VehicleManagementDataAccess : IVehicleManagementDataAccess
{
    private readonly string _connectionString;

    public VehicleManagementDataAccess(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection is not configured.");
    }

    private SqlConnection CreateConnection() => new(_connectionString);

    public async Task<IReadOnlyList<VehicleCategoryModel>> GetCategoriesAsync(CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT EntryID, CategoryName, SortOrder, Active
            FROM dbo.VehicleCategories
            WHERE Active = 1
            ORDER BY SortOrder, CategoryName;
            """;

        var rows = new List<VehicleCategoryModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            rows.Add(new VehicleCategoryModel
            {
                EntryID = reader.GetInt32(0),
                CategoryName = reader.GetString(1),
                SortOrder = reader.GetInt32(2),
                Active = reader.GetBoolean(3)
            });
        }

        return rows;
    }

    public async Task<IReadOnlyList<VehicleEntryListRowModel>> GetVehiclesAsync(VehicleFilterModel filter, CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT v.EntryID, v.VehicleNumber, ISNULL(v.VehicleType, ''), ISNULL(v.DriverName, ''),
                   ISNULL(v.SupplierCompany, ''), ISNULL(v.Transporter, ''), ISNULL(c.CategoryName, ''),
                   v.EntryTime, v.ExitTime, ISNULL(v.Status, '')
            FROM dbo.VehicleEntries v
            LEFT JOIN dbo.VehicleCategories c ON c.EntryID = v.CategoryId
            WHERE (@CategoryId = 0 OR v.CategoryId = @CategoryId)
              AND (@ShowClosed = 1 OR ISNULL(v.Status, 'Inside') <> 'Closed')
              AND (@SearchText = '' OR v.VehicleNumber LIKE '%' + @SearchText + '%'
                   OR ISNULL(v.DriverName, '') LIKE '%' + @SearchText + '%'
                   OR ISNULL(v.SupplierCompany, '') LIKE '%' + @SearchText + '%'
                   OR ISNULL(v.Transporter, '') LIKE '%' + @SearchText + '%')
            ORDER BY v.EntryTime DESC, v.EntryID DESC;
            """;

        var rows = new List<VehicleEntryListRowModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@CategoryId", SqlDbType.Int).Value = filter.CategoryId;
        command.Parameters.Add("@ShowClosed", SqlDbType.Bit).Value = filter.ShowClosed;
        command.Parameters.Add("@SearchText", SqlDbType.VarChar, 100).Value = filter.SearchText?.Trim() ?? "";
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            rows.Add(new VehicleEntryListRowModel
            {
                EntryID = reader.GetInt32(0),
                VehicleNumber = reader.GetString(1),
                VehicleType = reader.GetString(2),
                DriverName = reader.GetString(3),
                SupplierCompany = reader.GetString(4),
                Transporter = reader.GetString(5),
                CategoryName = reader.GetString(6),
                EntryTime = reader.GetDateTime(7),
                ExitTime = reader.IsDBNull(8) ? null : reader.GetDateTime(8),
                Status = reader.GetString(9)
            });
        }

        return rows;
    }

    public async Task<VehicleEntryModel?> GetVehicleAsync(int entryId, CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT TOP (1) EntryID, VehicleNumber, ISNULL(VehicleType, ''), ISNULL(DriverName, ''),
                   ISNULL(DriverCNIC, ''), ISNULL(DriverMobile, ''), ISNULL(SupplierCompany, ''),
                   ISNULL(Transporter, ''), ISNULL(Purpose, ''), CategoryId, EntryTime, ExitTime,
                   ISNULL(VehiclePhotoBase64, ''), ISNULL(Status, 'Inside'), ISNULL(Remarks, '')
            FROM dbo.VehicleEntries
            WHERE EntryID = @EntryID;
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

        return new VehicleEntryModel
        {
            EntryID = reader.GetInt32(0),
            VehicleNumber = reader.GetString(1),
            VehicleType = reader.GetString(2),
            DriverName = reader.GetString(3),
            DriverCNIC = reader.GetString(4),
            DriverMobile = reader.GetString(5),
            SupplierCompany = reader.GetString(6),
            Transporter = reader.GetString(7),
            Purpose = reader.GetString(8),
            CategoryId = reader.IsDBNull(9) ? null : reader.GetInt32(9),
            EntryTime = reader.GetDateTime(10),
            ExitTime = reader.IsDBNull(11) ? null : reader.GetDateTime(11),
            VehiclePhotoDataUrl = reader.IsDBNull(12) || string.IsNullOrWhiteSpace(reader.GetString(12)) ? null : reader.GetString(12),
            Status = reader.GetString(13),
            Remarks = reader.GetString(14)
        };
    }

    public async Task<int> SaveVehicleAsync(VehicleEntryModel model, string userName, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(model);

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        if (model.EntryID == 0)
        {
            const string insertSql = """
                INSERT INTO dbo.VehicleEntries
                (
                    VehicleNumber, VehicleType, DriverName, DriverCNIC, DriverMobile,
                    SupplierCompany, Transporter, Purpose, CategoryId, EntryTime, ExitTime,
                    VehiclePhotoBase64, Status, CreatedBy, Remarks
                )
                OUTPUT INSERTED.EntryID
                VALUES
                (
                    @VehicleNumber, @VehicleType, @DriverName, @DriverCNIC, @DriverMobile,
                    @SupplierCompany, @Transporter, @Purpose, @CategoryId, @EntryTime, @ExitTime,
                    @VehiclePhotoBase64, @Status, @UserName, @Remarks
                );
                """;
            await using var insert = new SqlCommand(insertSql, connection);
            AddParameters(insert, model, userName);
            return Convert.ToInt32(await insert.ExecuteScalarAsync(cancellationToken));
        }

        const string updateSql = """
            UPDATE dbo.VehicleEntries
            SET VehicleNumber=@VehicleNumber, VehicleType=@VehicleType, DriverName=@DriverName,
                DriverCNIC=@DriverCNIC, DriverMobile=@DriverMobile, SupplierCompany=@SupplierCompany,
                Transporter=@Transporter, Purpose=@Purpose, CategoryId=@CategoryId, EntryTime=@EntryTime,
                ExitTime=@ExitTime, VehiclePhotoBase64=@VehiclePhotoBase64, Status=@Status,
                UpdatedBy=@UserName, UpdatedAt=GETDATE(), Remarks=@Remarks
            WHERE EntryID=@EntryID;
            """;
        await using var update = new SqlCommand(updateSql, connection);
        AddParameters(update, model, userName);
        update.Parameters.Add("@EntryID", SqlDbType.Int).Value = model.EntryID;
        await update.ExecuteNonQueryAsync(cancellationToken);
        return model.EntryID;
    }

    public async Task DeleteVehicleAsync(int entryId, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand("DELETE FROM dbo.VehicleEntries WHERE EntryID=@EntryID;", connection);
        command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static void AddParameters(SqlCommand command, VehicleEntryModel model, string userName)
    {
        command.Parameters.AddWithValue("@VehicleNumber", Db(model.VehicleNumber));
        command.Parameters.AddWithValue("@VehicleType", Db(model.VehicleType));
        command.Parameters.AddWithValue("@DriverName", Db(model.DriverName));
        command.Parameters.AddWithValue("@DriverCNIC", Db(model.DriverCNIC));
        command.Parameters.AddWithValue("@DriverMobile", Db(model.DriverMobile));
        command.Parameters.AddWithValue("@SupplierCompany", Db(model.SupplierCompany));
        command.Parameters.AddWithValue("@Transporter", Db(model.Transporter));
        command.Parameters.AddWithValue("@Purpose", Db(model.Purpose));
        command.Parameters.Add("@CategoryId", SqlDbType.Int).Value = model.CategoryId ?? (object)DBNull.Value;
        command.Parameters.Add("@EntryTime", SqlDbType.DateTime).Value = model.EntryTime;
        command.Parameters.Add("@ExitTime", SqlDbType.DateTime).Value = model.ExitTime ?? (object)DBNull.Value;
        command.Parameters.AddWithValue("@VehiclePhotoBase64", Db(model.VehiclePhotoDataUrl));
        command.Parameters.AddWithValue("@Status", Db(model.Status));
        command.Parameters.AddWithValue("@UserName", Db(userName));
        command.Parameters.AddWithValue("@Remarks", Db(model.Remarks));
    }

    private static object Db(string? value) => string.IsNullOrWhiteSpace(value) ? DBNull.Value : value.Trim();
}

