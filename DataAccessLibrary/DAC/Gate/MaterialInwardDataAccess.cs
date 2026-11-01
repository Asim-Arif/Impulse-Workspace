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

public sealed class MaterialInwardDataAccess : IMaterialInwardDataAccess
{
    private readonly string _connectionString;

    public MaterialInwardDataAccess(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection is not configured.");
    }

    private SqlConnection CreateConnection() => new(_connectionString);

    public async Task<IReadOnlyList<MaterialInwardListRowModel>> GetEntriesAsync(MaterialInwardFilterModel filter, CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT h.EntryID, ISNULL(h.PoNumber, ''), ISNULL(h.SupplierName, ''), ISNULL(h.VehicleNumber, ''),
                   ISNULL(h.DriverName, ''), ISNULL(h.Weighment, 0), h.EntryTime, ISNULL(h.Status, ''),
                   (SELECT COUNT(*) FROM dbo.MaterialInwardLines l WHERE l.MaterialInwardId = h.EntryID)
            FROM dbo.MaterialInward h
            WHERE (@ShowClosed = 1 OR ISNULL(h.Status, 'Open') <> 'Closed')
              AND (@SearchText = '' OR ISNULL(h.PoNumber, '') LIKE '%' + @SearchText + '%'
                   OR ISNULL(h.SupplierName, '') LIKE '%' + @SearchText + '%'
                   OR ISNULL(h.VehicleNumber, '') LIKE '%' + @SearchText + '%'
                   OR ISNULL(h.DriverName, '') LIKE '%' + @SearchText + '%')
            ORDER BY h.EntryTime DESC, h.EntryID DESC;
            """;
        var rows = new List<MaterialInwardListRowModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@ShowClosed", SqlDbType.Bit).Value = filter.ShowClosed;
        command.Parameters.Add("@SearchText", SqlDbType.VarChar, 100).Value = filter.SearchText?.Trim() ?? "";
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            rows.Add(new MaterialInwardListRowModel
            {
                EntryID = reader.GetInt32(0),
                PoNumber = reader.GetString(1),
                SupplierName = reader.GetString(2),
                VehicleNumber = reader.GetString(3),
                DriverName = reader.GetString(4),
                Weighment = reader.IsDBNull(5) ? 0m : Convert.ToDecimal(reader.GetValue(5)),
                EntryTime = reader.GetDateTime(6),
                Status = reader.GetString(7),
                LineCount = reader.GetInt32(8)
            });
        }
        return rows;
    }

    public async Task<MaterialInwardEntryModel?> GetEntryAsync(int entryId, CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        const string headerSql = """
            SELECT TOP (1) EntryID, ISNULL(PoNumber, ''), ISNULL(SupplierName, ''), VehicleEntryId,
                   ISNULL(VehicleNumber, ''), ISNULL(DriverName, ''), ISNULL(DeliveryChallanNo, ''),
                   ISNULL(InvoiceNo, ''), Weighment, EntryTime, ISNULL(Status, 'Open'), ISNULL(Remarks, '')
            FROM dbo.MaterialInward
            WHERE EntryID = @EntryID;
            """;
        await using var header = new SqlCommand(headerSql, connection);
        header.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        await using var reader = await header.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
        {
            return null;
        }
        var model = new MaterialInwardEntryModel
        {
            EntryID = reader.GetInt32(0),
            PoNumber = reader.GetString(1),
            SupplierName = reader.GetString(2),
            VehicleEntryId = reader.IsDBNull(3) ? null : reader.GetInt32(3),
            VehicleNumber = reader.GetString(4),
            DriverName = reader.GetString(5),
            DeliveryChallanNo = reader.GetString(6),
            InvoiceNo = reader.GetString(7),
            Weighment = reader.IsDBNull(8) ? null : Convert.ToDecimal(reader.GetValue(8)),
            EntryTime = reader.GetDateTime(9),
            Status = reader.GetString(10),
            Remarks = reader.GetString(11)
        };
        await reader.CloseAsync();

        const string lineSql = """
            SELECT EntryID, ISNULL(MaterialName, ''), ISNULL(Quantity, 0), ISNULL(Uom, '')
            FROM dbo.MaterialInwardLines
            WHERE MaterialInwardId = @EntryID
            ORDER BY EntryID;
            """;
        await using var lineCommand = new SqlCommand(lineSql, connection);
        lineCommand.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        await using var lineReader = await lineCommand.ExecuteReaderAsync(cancellationToken);
        while (await lineReader.ReadAsync(cancellationToken))
        {
            model.Lines.Add(new MaterialInwardLineModel
            {
                EntryID = lineReader.GetInt32(0),
                MaterialName = lineReader.GetString(1),
                Quantity = lineReader.IsDBNull(2) ? 0m : Convert.ToDecimal(lineReader.GetValue(2)),
                Uom = lineReader.GetString(3)
            });
        }
        return model;
    }

    public async Task<int> SaveAsync(MaterialInwardEntryModel model, string userName, CancellationToken cancellationToken = default)
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
                    INSERT INTO dbo.MaterialInward
                    (
                        PoNumber, SupplierName, VehicleEntryId, VehicleNumber, DriverName,
                        DeliveryChallanNo, InvoiceNo, Weighment, EntryTime, Status, CreatedBy, Remarks
                    )
                    OUTPUT INSERTED.EntryID
                    VALUES
                    (
                        @PoNumber, @SupplierName, @VehicleEntryId, @VehicleNumber, @DriverName,
                        @DeliveryChallanNo, @InvoiceNo, @Weighment, @EntryTime, @Status, @UserName, @Remarks
                    );
                    """;
                await using var insert = new SqlCommand(insertSql, connection, transaction);
                AddParameters(insert, model, userName);
                entryId = Convert.ToInt32(await insert.ExecuteScalarAsync(cancellationToken));
            }
            else
            {
                const string updateSql = """
                    UPDATE dbo.MaterialInward
                    SET PoNumber=@PoNumber, SupplierName=@SupplierName, VehicleEntryId=@VehicleEntryId,
                        VehicleNumber=@VehicleNumber, DriverName=@DriverName, DeliveryChallanNo=@DeliveryChallanNo,
                        InvoiceNo=@InvoiceNo, Weighment=@Weighment, EntryTime=@EntryTime, Status=@Status,
                        UpdatedBy=@UserName, UpdatedAt=GETDATE(), Remarks=@Remarks
                    WHERE EntryID=@EntryID;
                    DELETE FROM dbo.MaterialInwardLines WHERE MaterialInwardId=@EntryID;
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
                    INSERT INTO dbo.MaterialInwardLines(MaterialInwardId, MaterialName, Quantity, Uom)
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
            await using (var lines = new SqlCommand("DELETE FROM dbo.MaterialInwardLines WHERE MaterialInwardId=@EntryID;", connection, transaction))
            {
                lines.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
                await lines.ExecuteNonQueryAsync(cancellationToken);
            }

            await using (var header = new SqlCommand("DELETE FROM dbo.MaterialInward WHERE EntryID=@EntryID;", connection, transaction))
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

    private static void AddParameters(SqlCommand command, MaterialInwardEntryModel model, string userName)
    {
        command.Parameters.AddWithValue("@PoNumber", Db(model.PoNumber));
        command.Parameters.AddWithValue("@SupplierName", Db(model.SupplierName));
        command.Parameters.Add("@VehicleEntryId", SqlDbType.Int).Value = model.VehicleEntryId ?? (object)DBNull.Value;
        command.Parameters.AddWithValue("@VehicleNumber", Db(model.VehicleNumber));
        command.Parameters.AddWithValue("@DriverName", Db(model.DriverName));
        command.Parameters.AddWithValue("@DeliveryChallanNo", Db(model.DeliveryChallanNo));
        command.Parameters.AddWithValue("@InvoiceNo", Db(model.InvoiceNo));
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

