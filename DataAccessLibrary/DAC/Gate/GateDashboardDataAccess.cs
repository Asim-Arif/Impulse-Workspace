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

public sealed class GateDashboardDataAccess : IGateDashboardDataAccess
{
    private readonly string _connectionString;

    public GateDashboardDataAccess(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection is not configured.");
    }

    private SqlConnection CreateConnection() => new(_connectionString);

    public async Task<GateDashboardModel> GetDashboardAsync(CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        return new GateDashboardModel
        {
            Today = await GetSectionAsync(connection, DateTime.Today, DateTime.Today.AddDays(1), cancellationToken),
            LastSevenDays = await GetSectionAsync(connection, DateTime.Today.AddDays(-6), DateTime.Today.AddDays(1), cancellationToken),
            LastSevenDayTrend = await GetTrendAsync(connection, DateTime.Today.AddDays(-6), DateTime.Today.AddDays(1), cancellationToken)
        };
    }

    private static async Task<IReadOnlyList<GateDashboardDayModel>> GetTrendAsync(
        SqlConnection connection,
        DateTime from,
        DateTime to,
        CancellationToken cancellationToken)
    {
        var hasGateOut = await TableExistsAsync(connection, "dbo.MiscEntries_GateOut", cancellationToken);
        var gateOutCte = BuildGateOutTrendCte(hasGateOut);

        var sql = $$"""
            WITH Days AS
            (
                SELECT CAST(@From AS date) AS WorkDate
                UNION ALL
                SELECT DATEADD(day, 1, WorkDate)
                FROM Days
                WHERE WorkDate < DATEADD(day, -1, CAST(@To AS date))
            ),
            VisitorsData AS
            (
                SELECT CAST(DT AS date) AS WorkDate, COUNT(*) AS Cnt
                FROM dbo.Visitors
                WHERE OBJECT_ID(N'dbo.Visitors', N'U') IS NOT NULL
                  AND DT >= @From AND DT < @To
                GROUP BY CAST(DT AS date)
            ),
            GateInData AS
            (
                SELECT CAST(DT AS date) AS WorkDate, COUNT(*) AS Cnt
                FROM dbo.MiscEntries_GateInTemp
                WHERE OBJECT_ID(N'dbo.MiscEntries_GateInTemp', N'U') IS NOT NULL
                  AND DT >= @From AND DT < @To
                GROUP BY CAST(DT AS date)
            ),
            GateOutData AS
            (
                {{gateOutCte}}
            )
            SELECT d.WorkDate,
                   ISNULL(v.Cnt, 0) AS Visitors,
                   ISNULL(gi.Cnt, 0) AS GateIn,
                   ISNULL(goa.Cnt, 0) AS GateOut
            FROM Days AS d
            LEFT JOIN VisitorsData AS v ON v.WorkDate = d.WorkDate
            LEFT JOIN GateInData AS gi ON gi.WorkDate = d.WorkDate
            LEFT JOIN
            (
                SELECT WorkDate, SUM(Cnt) AS Cnt
                FROM GateOutData
                GROUP BY WorkDate
            ) AS goa ON goa.WorkDate = d.WorkDate
            ORDER BY d.WorkDate
            OPTION (MAXRECURSION 10);
            """;

        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.Add("@From", SqlDbType.DateTime).Value = from;
        command.Parameters.Add("@To", SqlDbType.DateTime).Value = to;

        var result = new List<GateDashboardDayModel>();
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new GateDashboardDayModel
            {
                Date = Convert.ToDateTime(reader["WorkDate"]),
                Visitors = Convert.ToInt32(reader["Visitors"]),
                GateIn = Convert.ToInt32(reader["GateIn"]),
                GateOut = Convert.ToInt32(reader["GateOut"])
            });
        }

        return result;
    }

    private static async Task<GateDashboardSection> GetSectionAsync(
        SqlConnection connection,
        DateTime from,
        DateTime to,
        CancellationToken cancellationToken)
    {
        var hasGateOut = await TableExistsAsync(connection, "dbo.MiscEntries_GateOut", cancellationToken);
        var gateOutSql = BuildGateOutSectionSql(hasGateOut);

        var sql = $$"""
            SELECT
                Visitors =
                    CASE WHEN OBJECT_ID(N'dbo.Visitors', N'U') IS NULL THEN 0 ELSE
                    (SELECT COUNT(*) FROM dbo.Visitors WHERE DT >= @From AND DT < @To) END,
                GateIn =
                    CASE WHEN OBJECT_ID(N'dbo.MiscEntries_GateInTemp', N'U') IS NULL THEN 0 ELSE
                    (SELECT COUNT(*) FROM dbo.MiscEntries_GateInTemp WHERE DT >= @From AND DT < @To) END,
                GateOut = {{gateOutSql}};
            """;

        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.Add("@From", SqlDbType.DateTime).Value = from;
        command.Parameters.Add("@To", SqlDbType.DateTime).Value = to;

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
        {
            return new GateDashboardSection();
        }

        return new GateDashboardSection
        {
            Visitors = Convert.ToInt32(reader["Visitors"]),
            GateIn = Convert.ToInt32(reader["GateIn"]),
            GateOut = Convert.ToInt32(reader["GateOut"])
        };
    }

    private static async Task<bool> TableExistsAsync(
        SqlConnection connection,
        string fullyQualifiedName,
        CancellationToken cancellationToken)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT CASE WHEN OBJECT_ID(@ObjectName, 'U') IS NULL THEN 0 ELSE 1 END;";
        command.Parameters.Add("@ObjectName", SqlDbType.NVarChar, 256).Value = fullyQualifiedName;
        return Convert.ToInt32(await command.ExecuteScalarAsync(cancellationToken)) == 1;
    }

    private static string BuildGateOutTrendCte(bool hasGateOut)
    {
        var parts = new List<string>();

        if (hasGateOut)
        {
            parts.Add("""
                SELECT CAST(DT AS date) AS WorkDate, COUNT(*) AS Cnt
                FROM dbo.MiscEntries_GateOut
                WHERE DT >= @From AND DT < @To
                GROUP BY CAST(DT AS date)
                """);
        }

        return parts.Count == 0
            ? "SELECT CAST(@From AS date) AS WorkDate, CAST(0 AS int) AS Cnt WHERE 1 = 0"
            : string.Join("\r\nUNION ALL\r\n", parts);
    }

    private static string BuildGateOutSectionSql(bool hasGateOut)
    {
        var parts = new List<string>();

        if (hasGateOut)
        {
            parts.Add("(SELECT COUNT(*) FROM dbo.MiscEntries_GateOut WHERE DT >= @From AND DT < @To)");
        }

        return parts.Count == 0
            ? "0"
            : string.Join(" + ", parts);
    }
}

