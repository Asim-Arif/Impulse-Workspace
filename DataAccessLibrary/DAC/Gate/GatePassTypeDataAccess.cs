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

public sealed class GatePassTypeDataAccess : IGatePassTypeDataAccess
{
    private readonly string _connectionString;

    public GatePassTypeDataAccess(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection is not configured.");
    }

    private SqlConnection CreateConnection() => new(_connectionString);

    public async Task<IReadOnlyList<GatePassTypeModel>> GetAsync(CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT EntryID,
                   GatePassType,
                   CAST(ISNULL(Active, 1) AS bit) AS Active
            FROM dbo.GatePassTypes
            ORDER BY Active DESC, GatePassType;
            """;

        var rows = new List<GatePassTypeModel>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            rows.Add(new GatePassTypeModel
            {
                EntryId = reader.GetInt32(0),
                GatePassType = Convert.ToString(reader.GetValue(1))?.Trim() ?? "",
                Active = reader.GetBoolean(2)
            });
        }

        return rows;
    }

    public async Task<int> SaveAsync(GatePassTypeModel model, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(model);
        if (string.IsNullOrWhiteSpace(model.GatePassType))
        {
            throw new InvalidOperationException("Please enter gate-pass type.");
        }

        var name = model.GatePassType.Trim();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        if (model.EntryId == 0)
        {
            const string insertSql = """
                IF EXISTS (SELECT 1 FROM dbo.GatePassTypes WHERE UPPER(LTRIM(RTRIM(GatePassType))) = UPPER(@GatePassType))
                    THROW 51000, 'Gate-pass type already exists.', 1;

                INSERT INTO dbo.GatePassTypes (GatePassType, Active)
                OUTPUT INSERTED.EntryID
                VALUES (@GatePassType, @Active);
                """;

            await using var command = new SqlCommand(insertSql, connection);
            command.Parameters.Add("@GatePassType", SqlDbType.VarChar, 50).Value = name;
            command.Parameters.Add("@Active", SqlDbType.Bit).Value = model.Active;
            return Convert.ToInt32(await command.ExecuteScalarAsync(cancellationToken));
        }

        const string updateSql = """
            IF EXISTS
            (
                SELECT 1
                FROM dbo.GatePassTypes
                WHERE UPPER(LTRIM(RTRIM(GatePassType))) = UPPER(@GatePassType)
                  AND EntryID <> @EntryID
            )
                THROW 51000, 'Another gate-pass type with this name already exists.', 1;

            UPDATE dbo.GatePassTypes
            SET GatePassType = @GatePassType,
                Active = @Active
            WHERE EntryID = @EntryID;
            """;

        await using var updateCommand = new SqlCommand(updateSql, connection);
        updateCommand.Parameters.Add("@GatePassType", SqlDbType.VarChar, 50).Value = name;
        updateCommand.Parameters.Add("@Active", SqlDbType.Bit).Value = model.Active;
        updateCommand.Parameters.Add("@EntryID", SqlDbType.Int).Value = model.EntryId;
        if (await updateCommand.ExecuteNonQueryAsync(cancellationToken) == 0)
        {
            throw new InvalidOperationException("Gate-pass type was not found.");
        }

        return model.EntryId;
    }

    public async Task SetActiveAsync(int entryId, bool active, CancellationToken cancellationToken = default)
    {
        const string sql = """
            UPDATE dbo.GatePassTypes
            SET Active = @Active
            WHERE EntryID = @EntryID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@Active", SqlDbType.Bit).Value = active;
        command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        if (await command.ExecuteNonQueryAsync(cancellationToken) == 0)
        {
            throw new InvalidOperationException("Gate-pass type was not found.");
        }
    }

    public async Task DeleteAsync(int entryId, CancellationToken cancellationToken = default)
    {
        const string sql = """
            IF EXISTS (SELECT 1 FROM dbo.MiscEntries_GateInTemp WHERE GatePassTypeName IN (SELECT GatePassType FROM dbo.GatePassTypes WHERE EntryID = @EntryID))
                THROW 51000, 'Cannot delete gate-pass type because it is in use by Gate In records.', 1;

            DELETE FROM dbo.GatePassTypes WHERE EntryID = @EntryID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.Add("@EntryID", SqlDbType.Int).Value = entryId;
        await command.ExecuteNonQueryAsync(cancellationToken);
    }
}

