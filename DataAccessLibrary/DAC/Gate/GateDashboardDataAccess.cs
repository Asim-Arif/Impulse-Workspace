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
        const string sql = """
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
            GateOutPermData AS
            (
                SELECT CAST(DT AS date) AS WorkDate, COUNT(*) AS Cnt
                FROM dbo.MiscEntries_GateOut
                WHERE OBJECT_ID(N'dbo.MiscEntries_GateOut', N'U') IS NOT NULL
                  AND DT >= @From AND DT < @To
                GROUP BY CAST(DT AS date)
            ),
            GateOutRetData AS
            (
                SELECT CAST(DT AS date) AS WorkDate, COUNT(*) AS Cnt
                FROM dbo.GatePass_Return
                WHERE OBJECT_ID(N'dbo.GatePass_Return', N'U') IS NOT NULL
                  AND DT >= @From AND DT < @To
                GROUP BY CAST(DT AS date)
            ),
            CourierData AS
            (
                SELECT CAST(DT AS date) AS WorkDate, COUNT(*) AS Cnt
                FROM dbo.Courier_Issuance
                WHERE OBJECT_ID(N'dbo.Courier_Issuance', N'U') IS NOT NULL
                  AND DT >= @From AND DT < @To
                GROUP BY CAST(DT AS date)
            ),
            VehicleData AS
            (
                SELECT CAST(ISNULL(EntryTime, CreatedAt) AS date) AS WorkDate, COUNT(*) AS Cnt
                FROM dbo.VehicleEntries
                WHERE OBJECT_ID(N'dbo.VehicleEntries', N'U') IS NOT NULL
                  AND ISNULL(EntryTime, CreatedAt) >= @From AND ISNULL(EntryTime, CreatedAt) < @To
                GROUP BY CAST(ISNULL(EntryTime, CreatedAt) AS date)
            )
            SELECT d.WorkDate,
                   ISNULL(v.Cnt, 0) AS Visitors,
                   ISNULL(gi.Cnt, 0) AS GateIn,
                   ISNULL(gop.Cnt, 0) AS GateOutPermanent,
                   ISNULL(gor.Cnt, 0) AS GateOutReturnable,
                   ISNULL(c.Cnt, 0) AS Courier,
                   ISNULL(vh.Cnt, 0) AS Vehicles
            FROM Days AS d
            LEFT JOIN VisitorsData AS v ON v.WorkDate = d.WorkDate
            LEFT JOIN GateInData AS gi ON gi.WorkDate = d.WorkDate
            LEFT JOIN GateOutPermData AS gop ON gop.WorkDate = d.WorkDate
            LEFT JOIN GateOutRetData AS gor ON gor.WorkDate = d.WorkDate
            LEFT JOIN CourierData AS c ON c.WorkDate = d.WorkDate
            LEFT JOIN VehicleData AS vh ON vh.WorkDate = d.WorkDate
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
                GateOutPermanent = Convert.ToInt32(reader["GateOutPermanent"]),
                GateOutReturnable = Convert.ToInt32(reader["GateOutReturnable"]),
                Courier = Convert.ToInt32(reader["Courier"]),
                Vehicles = Convert.ToInt32(reader["Vehicles"])
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
        const string sql = """
            SELECT
                Visitors =
                    CASE WHEN OBJECT_ID(N'dbo.Visitors', N'U') IS NULL THEN 0 ELSE
                    (SELECT COUNT(*) FROM dbo.Visitors WHERE DT >= @From AND DT < @To) END,
                GateIn =
                    CASE WHEN OBJECT_ID(N'dbo.MiscEntries_GateInTemp', N'U') IS NULL THEN 0 ELSE
                    (SELECT COUNT(*) FROM dbo.MiscEntries_GateInTemp WHERE DT >= @From AND DT < @To) END,
                GateOutPermanent =
                    CASE WHEN OBJECT_ID(N'dbo.MiscEntries_GateOut', N'U') IS NULL THEN 0 ELSE
                    (SELECT COUNT(*) FROM dbo.MiscEntries_GateOut WHERE DT >= @From AND DT < @To) END,
                GateOutReturnable =
                    CASE WHEN OBJECT_ID(N'dbo.GatePass_Return', N'U') IS NULL THEN 0 ELSE
                    (SELECT COUNT(*) FROM dbo.GatePass_Return WHERE DT >= @From AND DT < @To) END,
                Courier =
                    CASE WHEN OBJECT_ID(N'dbo.Courier_Issuance', N'U') IS NULL THEN 0 ELSE
                    (SELECT COUNT(*) FROM dbo.Courier_Issuance WHERE DT >= @From AND DT < @To) END,
                Vehicles =
                    CASE WHEN OBJECT_ID(N'dbo.VehicleEntries', N'U') IS NULL THEN 0 ELSE
                    (SELECT COUNT(*) FROM dbo.VehicleEntries WHERE (EntryTime >= @From AND EntryTime < @To) OR (CreatedAt >= @From AND CreatedAt < @To)) END;
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
            GateOutPermanent = Convert.ToInt32(reader["GateOutPermanent"]),
            GateOutReturnable = Convert.ToInt32(reader["GateOutReturnable"]),
            Courier = Convert.ToInt32(reader["Courier"]),
            Vehicles = Convert.ToInt32(reader["Vehicles"])
        };
    }
}
