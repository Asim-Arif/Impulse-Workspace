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

public sealed class MaterialOutwardDataAccess : IMaterialOutwardDataAccess
{
    private readonly string _connectionString;

    public MaterialOutwardDataAccess(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection is not configured.");
    }

    private SqlConnection CreateConnection() => new(_connectionString);

    public async Task<IReadOnlyList<MaterialOutwardListRowModel>> GetEntriesAsync(MaterialOutwardFilterModel filter, CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT h.EntryID, ISNULL(h.SalesOrderNo, ''), ISNULL(h.DeliveryOrderNo, ''), ISNULL(h.CustomerName, ''),
                   ISNULL(h.VehicleNumber, ''), ISNULL(h.DriverName, ''), ISNULL(h.Weighment, 0), h.EntryTime,
                   ISNULL(h.Status, ''), (SELECT COUNT(*) FROM dbo.MaterialOutwardLines l WHERE l.MaterialOutwardId = h.EntryID)
            FROM dbo.MaterialOutward h
            WHERE (@ShowClosed = 1 OR ISNULL(h.Status, 'Dispatched') <> 'Closed')
              AND (@SearchText = '' OR ISNULL(h.SalesOrderNo, '') LIKE '%' + @SearchText + '%'
                   OR ISNULL(h.DeliveryOrderNo, '') LIKE '%' + @SearchText + '%'
                   OR ISNULL(h.CustomerName, '') LIKE '%' + @SearchText + '%'
                   OR ISNULL(h.VehicleNumber, '') LIKE '%' + @SearchText + '%')
            ORDER BY h.EntryTime DESC, h.EntryID DESC;
            """;
        var rows = new List<MaterialOutwardListRowModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@ShowClosed", SqlDbType.Bit).Value = filter.ShowClosed;
        command.Parameters.Add("@SearchText", SqlDbType.VarChar, 100).Value = filter.SearchText?.Trim() ?? "";
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            rows.Add(new MaterialOutwardListRowModel
            {
                EntryID = reader.GetInt32(0),
                SalesOrderNo = reader.GetString(1),
                DeliveryOrderNo = reader.GetString(2),
                CustomerName = reader.GetString(3),
                VehicleNumber = reader.GetString(4),
                DriverName = reader.GetString(5),
                Weighment = reader.IsDBNull(6) ? 0m : Convert.ToDecimal(reader.GetValue(6)),
                EntryTime = reader.GetDateTime(7),
                Status = reader.GetString(8),
                LineCount = reader.GetInt32(9)
            });
        }
        return rows;
    }

    public async Task<MaterialOutwardEntryModel?> GetEntryAsync(int entryId, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        const string headerSql = """
            SELECT TOP (1) EntryID, ISNULL(SalesOrderNo, ''), ISNULL(DeliveryOrderNo, ''), ISNULL(CustomerName, ''),
                   VehicleEntryId, ISNULL(VehicleNumber, ''), ISNULL(DriverName, ''), Weighment,
                   EntryTime, ISNULL(Status, 'Dispatched'), ISNULL(Remarks, '')
            FROM dbo.MaterialOutward
            WHERE EntryID = @EntryID;
            """;
        await using var header = new SqlCommand(headerSql, connection);
        header.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        await using var reader = await header.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
        {
            return null;
        }

        var model = new MaterialOutwardEntryModel
        {
            EntryID = reader.GetInt32(0),
            SalesOrderNo = reader.GetString(1),
            DeliveryOrderNo = reader.GetString(2),
            CustomerName = reader.GetString(3),
            VehicleEntryId = reader.IsDBNull(4) ? null : reader.GetInt32(4),
            VehicleNumber = reader.GetString(5),
            DriverName = reader.GetString(6),
            Weighment = reader.IsDBNull(7) ? null : Convert.ToDecimal(reader.GetValue(7)),
            EntryTime = reader.GetDateTime(8),
            Status = reader.GetString(9),
            Remarks = reader.GetString(10)
        };
        await reader.CloseAsync();

        const string lineSql = """
            SELECT EntryID, ISNULL(MaterialName, ''), ISNULL(Quantity, 0), ISNULL(Uom, '')
            FROM dbo.MaterialOutwardLines
            WHERE MaterialOutwardId = @EntryID
            ORDER BY EntryID;
            """;
        await using var lineCommand = new SqlCommand(lineSql, connection);
        lineCommand.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        await using var lineReader = await lineCommand.ExecuteReaderAsync(cancellationToken);
        while (await lineReader.ReadAsync(cancellationToken))
        {
            model.Lines.Add(new MaterialOutwardLineModel
            {
                EntryID = lineReader.GetInt32(0),
                MaterialName = lineReader.GetString(1),
                Quantity = lineReader.IsDBNull(2) ? 0m : Convert.ToDecimal(lineReader.GetValue(2)),
                Uom = lineReader.GetString(3)
            });
        }
        return model;
    }

    public async Task<int> SaveAsync(MaterialOutwardEntryModel model, string userName, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);
        try
        {
            int entryId;
            if (model.EntryID == 0)
            {
                const string insertSql = """
                    INSERT INTO dbo.MaterialOutward
                    (
                        SalesOrderNo, DeliveryOrderNo, CustomerName, VehicleEntryId, VehicleNumber,
                        DriverName, Weighment, EntryTime, Status, CreatedBy, Remarks
                    )
                    OUTPUT INSERTED.EntryID
                    VALUES
                    (
                        @SalesOrderNo, @DeliveryOrderNo, @CustomerName, @VehicleEntryId, @VehicleNumber,
                        @DriverName, @Weighment, @EntryTime, @Status, @UserName, @Remarks
                    );
                    """;
                await using var insert = new SqlCommand(insertSql, connection, transaction);
                AddParameters(insert, model, userName);
                entryId = Convert.ToInt32(await insert.ExecuteScalarAsync(cancellationToken));
            }
            else
            {
                const string updateSql = """
                    UPDATE dbo.MaterialOutward
                    SET SalesOrderNo=@SalesOrderNo, DeliveryOrderNo=@DeliveryOrderNo, CustomerName=@CustomerName,
                        VehicleEntryId=@VehicleEntryId, VehicleNumber=@VehicleNumber, DriverName=@DriverName,
                        Weighment=@Weighment, EntryTime=@EntryTime, Status=@Status, UpdatedBy=@UserName,
                        UpdatedAt=GETDATE(), Remarks=@Remarks
                    WHERE EntryID=@EntryID;
                    DELETE FROM dbo.MaterialOutwardLines WHERE MaterialOutwardId=@EntryID;
                    """;
                await using var update = new SqlCommand(updateSql, connection, transaction);
                AddParameters(update, model, userName);
                update.Parameters.Add("@EntryID", SqlDbType.Int).Value = model.EntryID;
                await update.ExecuteNonQueryAsync(cancellationToken);
                entryId = model.EntryID;
            }

            foreach (var line in model.Lines.Where(x => !string.IsNullOrWhiteSpace(x.MaterialName) && x.Quantity > 0))
            {
                const string lineInsert = """
                    INSERT INTO dbo.MaterialOutwardLines(MaterialOutwardId, MaterialName, Quantity, Uom)
                    VALUES(@HeaderId, @MaterialName, @Quantity, @Uom);
                    """;
                await using var detail = new SqlCommand(lineInsert, connection, transaction);
                detail.Parameters.Add("@HeaderId", SqlDbType.Int).Value = entryId;
                detail.Parameters.AddWithValue("@MaterialName", Db(line.MaterialName));
                detail.Parameters.Add("@Quantity", SqlDbType.Decimal).Value = line.Quantity;
                detail.Parameters["@Quantity"].Precision = 18;
                detail.Parameters["@Quantity"].Scale = 3;
                detail.Parameters.AddWithValue("@Uom", Db(line.Uom));
                await detail.ExecuteNonQueryAsync(cancellationToken);
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

    public async Task<IReadOnlyList<VehicleEntryListRowModel>> GetVehicleOptionsAsync(CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT TOP (200) EntryID, VehicleNumber, ISNULL(VehicleType, ''), ISNULL(DriverName, ''),
                   ISNULL(SupplierCompany, ''), ISNULL(Transporter, ''), '', EntryTime, ExitTime, ISNULL(Status, '')
            FROM dbo.VehicleEntries
            ORDER BY EntryTime DESC, EntryID DESC;
            """;
        var rows = new List<VehicleEntryListRowModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
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

    public async Task DeleteAsync(int entryId, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);
        try
        {
            await using (var lines = new SqlCommand("DELETE FROM dbo.MaterialOutwardLines WHERE MaterialOutwardId=@EntryID;", connection, transaction))
            {
                lines.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
                await lines.ExecuteNonQueryAsync(cancellationToken);
            }

            await using (var header = new SqlCommand("DELETE FROM dbo.MaterialOutward WHERE EntryID=@EntryID;", connection, transaction))
            {
                header.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
                await header.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
        }
        catch
        {
            await transaction.RollbackAsync(cancellationToken);
            throw;
        }
    }

    private static void AddParameters(SqlCommand command, MaterialOutwardEntryModel model, string userName)
    {
        command.Parameters.AddWithValue("@SalesOrderNo", Db(model.SalesOrderNo));
        command.Parameters.AddWithValue("@DeliveryOrderNo", Db(model.DeliveryOrderNo));
        command.Parameters.AddWithValue("@CustomerName", Db(model.CustomerName));
        command.Parameters.Add("@VehicleEntryId", SqlDbType.Int).Value = model.VehicleEntryId ?? (object)DBNull.Value;
        command.Parameters.AddWithValue("@VehicleNumber", Db(model.VehicleNumber));
        command.Parameters.AddWithValue("@DriverName", Db(model.DriverName));
        command.Parameters.Add("@Weighment", SqlDbType.Decimal).Value = model.Weighment ?? (object)DBNull.Value;
        command.Parameters["@Weighment"].Precision = 18;
        command.Parameters["@Weighment"].Scale = 3;
        command.Parameters.Add("@EntryTime", SqlDbType.DateTime).Value = model.EntryTime;
        command.Parameters.AddWithValue("@Status", Db(model.Status));
        command.Parameters.AddWithValue("@UserName", Db(userName));
        command.Parameters.AddWithValue("@Remarks", Db(model.Remarks));
    }

    private static object Db(string? value) => string.IsNullOrWhiteSpace(value) ? DBNull.Value : value.Trim();
}

