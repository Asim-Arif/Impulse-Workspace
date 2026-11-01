using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.FixedAssets;
using DataAccessLibrary.Models.ViewModels.FixedAssets;
using Microsoft.Extensions.Configuration;
using Microsoft.Data.SqlClient;

namespace DataAccessLibrary.DAC.FixedAssets;

public sealed class FixedAssetProfileDataAccess : IFixedAssetProfileDataAccess
{
    private readonly string _connectionString;

    public FixedAssetProfileDataAccess(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection")
            ?? throw new InvalidOperationException("DefaultConnection is not configured.");
    }

    public async Task<FixedAssetLookupOptions> GetLookupOptionsAsync(CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        return new FixedAssetLookupOptions
        {
            Masters = await LoadLookupAsync(connection, "SELECT EntryID, ISNULL(Description,''), ISNULL(CAST(PercentRate AS decimal(18,2)),0) FROM dbo.FixedAssets_Master ORDER BY Description", cancellationToken),
            Departments = await LoadLookupAsync(connection, "SELECT DeptID, ISNULL(Name,''), 0 FROM dbo.Departments WHERE ISNULL(Active,1)=1 ORDER BY Name", cancellationToken),
            Employees = await LoadLookupAsync(connection, "SELECT EmpID, ISNULL(Name,''), 0 FROM dbo.Employees WHERE ISNULL(Active,1)=1 ORDER BY Name", cancellationToken),
            Accounts = await LoadLookupAsync(connection, "SELECT AccNo, ISNULL(AccTitle,''), 0 FROM dbo.Accounts WHERE ISNULL(Active,1)=1 AND ISNULL(Parent,0)=0 ORDER BY AccTitle", cancellationToken),
            Conventions = await LoadStringsAsync(connection, "SELECT DISTINCT ISNULL(Convention,'') FROM dbo.FixedAssets WHERE ISNULL(Convention,'')<>'' ORDER BY ISNULL(Convention,'')", cancellationToken)
        };
    }

    public async Task<FixedAssetProfileModel> NewAsync(CancellationToken cancellationToken = default)
    {
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        var nextAssetId = await NextAssetIdAsync(connection, cancellationToken);
        return new FixedAssetProfileModel
        {
            AssetId = nextAssetId,
            AcquiredDate = DateTime.Today,
            PlacedInServiceDate = DateTime.Today,
            Status = 0
        };
    }

    public async Task<IReadOnlyList<FixedAssetListRow>> GetListAsync(int? masterId, string? searchText, bool includeInactive, CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT TOP 1000 f.EntryID, ISNULL(f.AssetID,''), ISNULL(f.Description,''), ISNULL(f.OldCode,''), ISNULL(f.Location,''),
                   f.AcquiredDT, ISNULL(CAST(f.AcquisitionCost AS decimal(18,2)),0),
                   ISNULL(CAST(f.AcumulatedDepriciation AS decimal(18,2)),0), CAST(0 AS bit),
                   ISNULL(f.VchrNo,''), 'Available', '', ISNULL(f.InActive,0),
                   ISNULL(f.Disposed,0), ISNULL(f.Disposed_Type,0),
                   ISNULL(m.Description,'')
            FROM dbo.FixedAssets f
            LEFT JOIN dbo.FixedAssets_Master m ON f.MasterFA_RefID=m.EntryID
            WHERE (@IncludeInactive=1 OR ISNULL(f.InActive,0)=0)
              AND (@MasterId IS NULL OR f.MasterFA_RefID=@MasterId)
              AND (@Search='' OR ISNULL(f.AssetID,'') LIKE @Like OR ISNULL(f.Description,'') LIKE @Like OR ISNULL(f.OldCode,'') LIKE @Like)
            ORDER BY ISNULL(m.Description,''),
                     CASE WHEN ISNUMERIC(f.AssetID)=1 THEN CAST(f.AssetID AS int) ELSE 2147483647 END,
                     f.AssetID;
            """;

        var rows = new List<FixedAssetListRow>();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        var search = (searchText ?? "").Trim();
        command.Parameters.AddWithValue("@IncludeInactive", includeInactive);
        command.Parameters.AddWithValue("@MasterId", masterId.HasValue ? masterId.Value : DBNull.Value);
        command.Parameters.AddWithValue("@Search", search);
        command.Parameters.AddWithValue("@Like", $"%{search}%");

        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            rows.Add(new FixedAssetListRow
            {
                EntryId = reader.GetInt32(0),
                AssetId = ReadString(reader, 1),
                Description = ReadString(reader, 2),
                OldCode = ReadString(reader, 3),
                Location = ReadString(reader, 4),
                AcquiredDate = ReadDate(reader, 5),
                Cost = ReadDecimal(reader, 6),
                AccumulatedDepreciation = ReadDecimal(reader, 7),
                AutoEntry = ReadBool(reader, 8),
                VoucherNo = ReadString(reader, 9),
                GatePassStatus = ReadString(reader, 10),
                ReceiptNoteStatus = ReadString(reader, 11),
                InActive = ReadBool(reader, 12),
                Disposed = ReadBool(reader, 13),
                DisposedType = ReadInt(reader, 14),
                Category = ReadString(reader, 15)
            });
        }

        return rows;
    }

    public async Task<FixedAssetAdjustmentModel> NewAdjustmentAsync(string? assetId = null, CancellationToken cancellationToken = default)
    {
        var model = new FixedAssetAdjustmentModel();
        if (string.IsNullOrWhiteSpace(assetId)) return model;

        model.AssetId = assetId.Trim();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await FillAdjustmentAssetAsync(connection, model, cancellationToken);
        return model;
    }

    public async Task<int> SaveAdjustmentAsync(FixedAssetAdjustmentModel model, string userName, string machineName, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(model.AssetId)) throw new InvalidOperationException("Please select asset.");

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await FillAdjustmentAssetAsync(connection, model, cancellationToken);

        const string sql = """
            DECLARE @NewID int;
            SELECT @NewID = ISNULL(MAX(EntryID), 0) + 1 FROM dbo.FixedAssets_Adjustment;
            INSERT INTO dbo.FixedAssets_Adjustment
            (EntryID, AssetID, DT, Remarks, Cost, Adjustment, Attachment_Type, PDFAttachments, UserName, MachineName, DTEntry)
            VALUES
            (@NewID, @AssetID, @DT, @Remarks, @Cost, @Adjustment, @AttachmentType, @PdfAttachments, @UserName, @MachineName, GETDATE());
            SELECT @NewID;
            """;

        var pdfFileName = Path.GetFileName(model.PdfPath ?? "");
        var attachmentType = Path.GetExtension(pdfFileName).TrimStart('.');
        if (!string.IsNullOrWhiteSpace(pdfFileName))
        {
            pdfFileName = pdfFileName.Length > 255 ? pdfFileName[^255..] : pdfFileName;
            attachmentType = attachmentType.Length > 50 ? attachmentType[..50] : attachmentType;
        }

        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@AssetID", TrimTo(model.AssetId, 50));
        command.Parameters.AddWithValue("@DT", model.AdjustmentDate);
        command.Parameters.AddWithValue("@Remarks", TrimTo(model.Remarks, 500));
        command.Parameters.AddWithValue("@Cost", model.AcquisitionCost);
        command.Parameters.AddWithValue("@Adjustment", model.AdjustmentAmount);
        command.Parameters.AddWithValue("@AttachmentType", attachmentType);
        command.Parameters.AddWithValue("@PdfAttachments", pdfFileName);
        command.Parameters.AddWithValue("@UserName", TrimTo(userName, 50));
        command.Parameters.AddWithValue("@MachineName", TrimTo(machineName, 50));
        return Convert.ToInt32(await command.ExecuteScalarAsync(cancellationToken));
    }

    public async Task<FixedAssetMoveModel> NewMoveAsync(string? assetId = null, CancellationToken cancellationToken = default)
    {
        var model = new FixedAssetMoveModel();
        if (string.IsNullOrWhiteSpace(assetId)) return model;

        model.AssetId = assetId.Trim();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await FillMoveAssetAsync(connection, model, cancellationToken);
        model.NewHeadId = model.CurrentHeadId;
        return model;
    }

    public async Task SaveMoveAsync(FixedAssetMoveModel model, string userName, string machineName, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(model.AssetId)) throw new InvalidOperationException("Please select asset.");
        if (!model.NewHeadId.HasValue) throw new InvalidOperationException("Please select new head.");

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await FillMoveAssetAsync(connection, model, cancellationToken);
        if (model.CurrentHeadId == model.NewHeadId) throw new InvalidOperationException("New head must be different from original head.");

        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);
        try
        {
            const string logSql = """
                INSERT INTO dbo.FixedAssets_Head_Changes
                (EntryID, AssetID, FA_M_RefID_Old, FA_M_RefID_New, Remarks, UserName, MachineName, DTEntry)
                SELECT ISNULL(MAX(EntryID), 0) + 1,
                       @AssetID,
                       @OldHeadId,
                       @NewHeadId,
                       @Remarks,
                       @UserName,
                       @MachineName,
                       GETDATE()
                FROM dbo.FixedAssets_Head_Changes;
                """;
            await using (var logCommand = new SqlCommand(logSql, connection, transaction))
            {
                logCommand.Parameters.AddWithValue("@AssetID", TrimTo(model.AssetId, 50));
                logCommand.Parameters.AddWithValue("@OldHeadId", model.CurrentHeadId ?? 0);
                logCommand.Parameters.AddWithValue("@NewHeadId", model.NewHeadId.Value);
                logCommand.Parameters.AddWithValue("@Remarks", TrimTo(model.Remarks, 255));
                logCommand.Parameters.AddWithValue("@UserName", TrimTo(userName, 50));
                logCommand.Parameters.AddWithValue("@MachineName", TrimTo(machineName, 50));
                await logCommand.ExecuteNonQueryAsync(cancellationToken);
            }

            const string updateSql = "UPDATE dbo.FixedAssets SET MasterFA_RefID=@NewHeadId WHERE AssetID=@AssetID;";
            await using (var updateCommand = new SqlCommand(updateSql, connection, transaction))
            {
                updateCommand.Parameters.AddWithValue("@NewHeadId", model.NewHeadId.Value);
                updateCommand.Parameters.AddWithValue("@AssetID", TrimTo(model.AssetId, 50));
                await updateCommand.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
        }
        catch
        {
            await transaction.RollbackAsync(cancellationToken);
            throw;
        }
    }

    public async Task<FixedAssetDisposalModel> NewDisposalAsync(string? assetId = null, CancellationToken cancellationToken = default)
    {
        var model = new FixedAssetDisposalModel();
        if (string.IsNullOrWhiteSpace(assetId)) return model;

        model.AssetId = assetId.Trim();
        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await FillDisposalAssetAsync(connection, model, cancellationToken);
        return model;
    }

    public async Task SaveDisposalAsync(FixedAssetDisposalModel model, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(model.AssetId)) throw new InvalidOperationException("Please select asset.");
        if (model.SalePrice < 0) throw new InvalidOperationException("Invalid sale price.");
        if (model.DisposalType is not "0" and not "1") throw new InvalidOperationException("Please select disposed or write off.");

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await FillDisposalAssetAsync(connection, model, cancellationToken);

        const string sql = """
            UPDATE dbo.FixedAssets
            SET Disposed=1,
                DTDisposal=@DTDisposal,
                SalePrice=@SalePrice,
                WDV=@WDV,
                Remarks=@Remarks,
                Disposed_Type=@DisposedType
            WHERE AssetID=@AssetID;
            """;

        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@DTDisposal", model.DisposalDate);
        command.Parameters.AddWithValue("@SalePrice", model.SalePrice);
        command.Parameters.AddWithValue("@WDV", model.Wdv);
        command.Parameters.AddWithValue("@Remarks", TrimTo(model.Remarks, 1000));
        command.Parameters.AddWithValue("@DisposedType", byte.Parse(model.DisposalType));
        command.Parameters.AddWithValue("@AssetID", TrimTo(model.AssetId, 50));
        var affected = await command.ExecuteNonQueryAsync(cancellationToken);
        if (affected == 0) throw new InvalidOperationException("Selected asset was not found.");
    }

    private static async Task FillDisposalAssetAsync(SqlConnection connection, FixedAssetDisposalModel model, CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT TOP 1 ISNULL(CAST(AcquisitionCost AS decimal(18,2)),0),
                   ISNULL(CAST(AcumulatedDepriciation AS decimal(18,2)),0)
            FROM dbo.FixedAssets
            WHERE AssetID=@AssetID AND ISNULL(Disposed,0)=0;
            """;

        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@AssetID", model.AssetId.Trim());
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken)) throw new InvalidOperationException("Selected asset was not found or is already disposed.");
        model.Wdv = Math.Round(ReadDecimal(reader, 0) - ReadDecimal(reader, 1), 2);
    }

    private static async Task FillMoveAssetAsync(SqlConnection connection, FixedAssetMoveModel model, CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT TOP 1 f.MasterFA_RefID, ISNULL(m.Description,'')
            FROM dbo.FixedAssets f
            LEFT JOIN dbo.FixedAssets_Master m ON f.MasterFA_RefID=m.EntryID
            WHERE f.AssetID=@AssetID;
            """;

        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@AssetID", model.AssetId.Trim());
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken)) throw new InvalidOperationException("Selected asset was not found.");
        model.CurrentHeadId = reader.IsDBNull(0) ? null : reader.GetInt32(0);
        model.OriginalHeadName = ReadString(reader, 1);
    }

    private static async Task FillAdjustmentAssetAsync(SqlConnection connection, FixedAssetAdjustmentModel model, CancellationToken cancellationToken)
    {
        const string sql = """
            SELECT TOP 1 AcquiredDT, ISNULL(CAST(AcquisitionCost AS decimal(18,2)),0),
                   ISNULL((SELECT SUM(CAST(Adjustment AS decimal(18,2))) FROM dbo.FixedAssets_Adjustment WHERE AssetID=@AssetID),0)
            FROM dbo.FixedAssets
            WHERE AssetID=@AssetID;
            """;

        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@AssetID", model.AssetId.Trim());
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken)) throw new InvalidOperationException("Selected asset was not found.");
        model.CapitalizationDate = ReadDate(reader, 0);
        model.AcquisitionCost = ReadDecimal(reader, 1);
        model.PreviousAdjustment = ReadDecimal(reader, 2);
    }

    public async Task<FixedAssetProfileModel?> GetAsync(int entryId, CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT TOP 1 EntryID, AssetID, ISNULL(Description,''), MasterFA_RefID, ISNULL(OldCode,''), AcquiredDT,
                   ISNULL(Location,''), ISNULL(ModelNo,''), ISNULL(SerialNo,''), ISNULL(Manufacturer,''), ISNULL(Notes,''),
                   ISNULL(EmpID,''), ISNULL(DeptID,''), ISNULL(VendAccNo,''), ISNULL(Status,0), PlacedInServiceDT,
                   ISNULL(CAST(AcquisitionCost AS decimal(18,2)),0), ISNULL(CAST(AcumulatedDepriciation AS decimal(18,2)),0),
                   ISNULL(CAST(WDV AS decimal(18,2)),0), ISNULL(RequiresCalibration,0), CalibrationFrequency,
                   ISNULL(AccumAccNo,''), ISNULL(DepriciationAccNo,''), DepriciationMethodID, ISNULL(Convention,''),
                   ISNULL(CAST(DepriciationRate AS decimal(18,2)),0), DepriciationBeginMonth, LifeInYears, ISNULL(InActive,0)
            FROM dbo.FixedAssets
            WHERE EntryID=@EntryID;
            """;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@EntryID", entryId);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        return await reader.ReadAsync(cancellationToken) ? Map(reader) : null;
    }

    public async Task<int> SaveAsync(FixedAssetProfileModel model, string userName, string machineName, CancellationToken cancellationToken = default)
    {
        Validate(model);

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

        try
        {
            await EnsureAssetIdIsUniqueAsync(connection, transaction, model, cancellationToken);

            if (model.EntryId > 0)
            {
                await UpdateAsync(connection, transaction, model, userName, machineName, cancellationToken);
            }
            else
            {
                model.EntryId = await InsertAsync(connection, transaction, model, userName, machineName, cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
            return model.EntryId;
        }
        catch
        {
            await transaction.RollbackAsync(cancellationToken);
            throw;
        }
    }

    public async Task<IReadOnlyList<FixedAssetRepairLogRow>> GetRepairLogsAsync(string assetId, CancellationToken cancellationToken = default)
    {
        const string sql = """
            SELECT ROW_NUMBER() OVER (ORDER BY EntryID DESC), DT, ISNULL(Remarks,''), ISNULL(MeterReading,0), ScheduleDT,
                   ISNULL(CAST(CostOfRepair AS decimal(18,2)),0), ISNULL(BillReference,'')
            FROM dbo.FixedAssets_RepairLog
            WHERE AssetID=@AssetID
            ORDER BY EntryID DESC;
            """;

        var rows = new List<FixedAssetRepairLogRow>();
        if (string.IsNullOrWhiteSpace(assetId)) return rows;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@AssetID", assetId.Trim());
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            rows.Add(new FixedAssetRepairLogRow
            {
                SerialNo = reader.GetInt64(0) > int.MaxValue ? int.MaxValue : (int)reader.GetInt64(0),
                Date = ReadDate(reader, 1),
                Remarks = ReadString(reader, 2),
                MeterReading = ReadInt(reader, 3),
                ScheduleDate = ReadDate(reader, 4),
                Cost = ReadDecimal(reader, 5),
                BillReference = ReadString(reader, 6)
            });
        }

        return rows;
    }

    public async Task DeleteAssetAsync(int entryId, string assetId, CancellationToken cancellationToken = default)
    {
        if (entryId <= 0 && string.IsNullOrWhiteSpace(assetId)) return;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(cancellationToken);

        try
        {
            const string sql = """
                DELETE FROM dbo.FixedAssets_Adjustment WHERE AssetID=@AssetID;
                DELETE FROM dbo.FixedAssets_Head_Changes WHERE AssetID=@AssetID;
                DELETE FROM dbo.FixedAssets_RepairLog WHERE AssetID=@AssetID;
                DELETE FROM dbo.FixedAssets_Calibration WHERE AssetID=@AssetID;
                DELETE FROM dbo.FixedAssets_FuelSlips WHERE AssetID=@AssetID;
                DELETE FROM dbo.FixedAssets_Approval WHERE AssetID=@AssetID;
                DELETE FROM dbo.FixedAssets WHERE EntryID=@EntryID;
                """;

            await using var command = new SqlCommand(sql, connection, transaction);
            command.Parameters.AddWithValue("@EntryID", entryId);
            command.Parameters.AddWithValue("@AssetID", assetId.Trim());
            await command.ExecuteNonQueryAsync(cancellationToken);
            await transaction.CommitAsync(cancellationToken);
        }
        catch
        {
            await transaction.RollbackAsync(cancellationToken);
            throw;
        }
    }

    public async Task<int> SaveMasterCategoryAsync(string description, decimal percentRate, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(description)) throw new InvalidOperationException("Category description is required.");

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        const string sql = """
            DECLARE @NewID int;
            SELECT @NewID = ISNULL(MAX(EntryID), 0) + 1 FROM dbo.FixedAssets_Master;
            INSERT INTO dbo.FixedAssets_Master (EntryID, Description, PercentRate)
            VALUES (@NewID, @Description, @PercentRate);
            SELECT @NewID;
            """;

        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@Description", description.Trim());
        command.Parameters.AddWithValue("@PercentRate", percentRate);
        return Convert.ToInt32(await command.ExecuteScalarAsync(cancellationToken));
    }

    public async Task SetActiveStatusAsync(int entryId, string assetId, bool isInactive, DateTime date, string remarks, string userName, string machineName, CancellationToken cancellationToken = default)
    {
        if (entryId <= 0 && string.IsNullOrWhiteSpace(assetId)) return;

        await using var connection = CreateConnection();
        await connection.OpenAsync(cancellationToken);

        const string sql = """
            UPDATE dbo.FixedAssets
            SET InActive = @InActive,
                InActiveDT = @DT,
                Remarks = @Remarks,
                UserName = @UserName,
                MachineName = @MachineName
            WHERE EntryID = @EntryID;
            """;

        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@EntryID", entryId);
        command.Parameters.AddWithValue("@InActive", isInactive);
        command.Parameters.AddWithValue("@DT", date);
        command.Parameters.AddWithValue("@Remarks", TrimTo(remarks, 500));
        command.Parameters.AddWithValue("@UserName", TrimTo(userName, 50));
        command.Parameters.AddWithValue("@MachineName", TrimTo(machineName, 50));
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static async Task<IReadOnlyList<FixedAssetLookupItem>> LoadLookupAsync(SqlConnection connection, string sql, CancellationToken cancellationToken)
    {
        var rows = new List<FixedAssetLookupItem>();
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            rows.Add(new FixedAssetLookupItem
            {
                Id = Convert.ToString(reader.GetValue(0)) ?? "",
                Name = ReadString(reader, 1),
                PercentRate = ReadDecimal(reader, 2)
            });
        }

        return rows;
    }

    private static async Task<IReadOnlyList<string>> LoadStringsAsync(SqlConnection connection, string sql, CancellationToken cancellationToken)
    {
        var rows = new List<string>();
        await using var command = new SqlCommand(sql, connection);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            rows.Add(ReadString(reader, 0));
        }

        return rows;
    }

    private static async Task<string> NextAssetIdAsync(SqlConnection connection, CancellationToken cancellationToken)
    {
        const string sql = "SELECT ISNULL(MAX(CASE WHEN ISNUMERIC(AssetID)=1 AND LEN(RTRIM(LTRIM(AssetID))) <= 6 AND CAST(AssetID AS int) < 900000 THEN CAST(AssetID AS int) ELSE 0 END), 0) + 1 FROM dbo.FixedAssets;";
        await using var command = new SqlCommand(sql, connection);
        var value = await command.ExecuteScalarAsync(cancellationToken);
        return Convert.ToInt32(value).ToString("0000");
    }

    private static async Task EnsureAssetIdIsUniqueAsync(SqlConnection connection, SqlTransaction transaction, FixedAssetProfileModel model, CancellationToken cancellationToken)
    {
        const string sql = "SELECT COUNT(*) FROM dbo.FixedAssets WHERE AssetID=@AssetID AND EntryID<>@EntryID;";
        await using var command = new SqlCommand(sql, connection, transaction);
        command.Parameters.AddWithValue("@AssetID", model.AssetId.Trim());
        command.Parameters.AddWithValue("@EntryID", model.EntryId);
        if (Convert.ToInt32(await command.ExecuteScalarAsync(cancellationToken)) > 0)
        {
            throw new InvalidOperationException("Asset ID already exists.");
        }
    }

    private static async Task<int> InsertAsync(SqlConnection connection, SqlTransaction transaction, FixedAssetProfileModel model, string userName, string machineName, CancellationToken cancellationToken)
    {
        const string sql = """
            INSERT INTO dbo.FixedAssets
            (AssetID, Description, AcquiredDT, Location, ModelNo, SerialNo, Manufacturer, Notes, DeptID, Status, VendAccNo,
             PlacedInServiceDT, AcquisitionCost, AccumAccNo, DepriciationAccNo, DepriciationMethodID, Convention,
             DepriciationBeginMonth, LifeInYears, InActive, MasterFA_RefID, RequiresCalibration, CalibrationFrequency,
             AcumulatedDepriciation, DepriciationRate, WDV, OldCode, EmpID, UserName, MachineName, EntryDT)
            OUTPUT INSERTED.EntryID
            VALUES
            (@AssetID, @Description, @AcquiredDT, @Location, @ModelNo, @SerialNo, @Manufacturer, @Notes, @DeptID, @Status, @VendAccNo,
             @PlacedInServiceDT, @AcquisitionCost, @AccumAccNo, @DepriciationAccNo, @DepriciationMethodID, @Convention,
             @DepriciationBeginMonth, @LifeInYears, @InActive, @MasterFA_RefID, @RequiresCalibration, @CalibrationFrequency,
             @AcumulatedDepriciation, @DepriciationRate, @WDV, @OldCode, @EmpID, @UserName, @MachineName, SYSDATETIME());
            """;

        await using var command = new SqlCommand(sql, connection, transaction);
        AddParameters(command, model, userName, machineName);
        return Convert.ToInt32(await command.ExecuteScalarAsync(cancellationToken));
    }

    private static async Task UpdateAsync(SqlConnection connection, SqlTransaction transaction, FixedAssetProfileModel model, string userName, string machineName, CancellationToken cancellationToken)
    {
        const string sql = """
            UPDATE dbo.FixedAssets
            SET AssetID=@AssetID, Description=@Description, AcquiredDT=@AcquiredDT, Location=@Location, ModelNo=@ModelNo,
                SerialNo=@SerialNo, Manufacturer=@Manufacturer, Notes=@Notes, DeptID=@DeptID, Status=@Status, VendAccNo=@VendAccNo,
                PlacedInServiceDT=@PlacedInServiceDT, AcquisitionCost=@AcquisitionCost, AccumAccNo=@AccumAccNo,
                DepriciationAccNo=@DepriciationAccNo, DepriciationMethodID=@DepriciationMethodID, Convention=@Convention,
                DepriciationBeginMonth=@DepriciationBeginMonth, LifeInYears=@LifeInYears, InActive=@InActive,
                MasterFA_RefID=@MasterFA_RefID, RequiresCalibration=@RequiresCalibration, CalibrationFrequency=@CalibrationFrequency,
                AcumulatedDepriciation=@AcumulatedDepriciation, DepriciationRate=@DepriciationRate, WDV=@WDV, OldCode=@OldCode,
                EmpID=@EmpID, UserName=@UserName, MachineName=@MachineName
            WHERE EntryID=@EntryID;
            """;

        await using var command = new SqlCommand(sql, connection, transaction);
        AddParameters(command, model, userName, machineName);
        command.Parameters.AddWithValue("@EntryID", model.EntryId);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static void AddParameters(SqlCommand command, FixedAssetProfileModel model, string userName, string machineName)
    {
        command.Parameters.AddWithValue("@AssetID", model.AssetId.Trim());
        command.Parameters.AddWithValue("@Description", model.Description.Trim());
        command.Parameters.AddWithValue("@AcquiredDT", DbValue(model.AcquiredDate));
        command.Parameters.AddWithValue("@Location", model.Location.Trim());
        command.Parameters.AddWithValue("@ModelNo", model.ModelNo.Trim());
        command.Parameters.AddWithValue("@SerialNo", model.SerialNo.Trim());
        command.Parameters.AddWithValue("@Manufacturer", model.Manufacturer.Trim());
        command.Parameters.AddWithValue("@Notes", model.Notes.Trim());
        command.Parameters.AddWithValue("@DeptID", model.DepartmentId.Trim());
        command.Parameters.AddWithValue("@Status", model.Status);
        command.Parameters.AddWithValue("@VendAccNo", model.VendorAccountNo.Trim());
        command.Parameters.AddWithValue("@PlacedInServiceDT", DbValue(model.PlacedInServiceDate));
        command.Parameters.AddWithValue("@AcquisitionCost", model.AcquisitionCost);
        command.Parameters.AddWithValue("@AccumAccNo", model.AccumulatedAccountNo.Trim());
        command.Parameters.AddWithValue("@DepriciationAccNo", model.DepreciationAccountNo.Trim());
        command.Parameters.AddWithValue("@DepriciationMethodID", DbValue(model.DepreciationMethodId));
        command.Parameters.AddWithValue("@Convention", model.Convention.Trim());
        command.Parameters.AddWithValue("@DepriciationBeginMonth", DbValue(model.DepreciationBeginMonth));
        command.Parameters.AddWithValue("@LifeInYears", DbValue(model.LifeInYears));
        command.Parameters.AddWithValue("@InActive", model.InActive);
        command.Parameters.AddWithValue("@MasterFA_RefID", DbValue(model.MasterAssetRefId));
        command.Parameters.AddWithValue("@RequiresCalibration", model.RequiresCalibration);
        command.Parameters.AddWithValue("@CalibrationFrequency", DbValue(model.CalibrationFrequency));
        command.Parameters.AddWithValue("@AcumulatedDepriciation", model.AccumulatedDepreciation);
        command.Parameters.AddWithValue("@DepriciationRate", model.DepreciationRate);
        command.Parameters.AddWithValue("@WDV", model.BookValue);
        command.Parameters.AddWithValue("@OldCode", model.OldCode.Trim());
        command.Parameters.AddWithValue("@EmpID", model.EmployeeId.Trim());
        command.Parameters.AddWithValue("@UserName", userName);
        command.Parameters.AddWithValue("@MachineName", machineName);
    }

    private static FixedAssetProfileModel Map(SqlDataReader reader)
    {
        var model = new FixedAssetProfileModel
        {
            EntryId = reader.GetInt32(0),
            AssetId = ReadString(reader, 1),
            Description = ReadString(reader, 2),
            MasterAssetRefId = ReadNullableInt(reader, 3),
            OldCode = ReadString(reader, 4),
            AcquiredDate = ReadDate(reader, 5),
            Location = ReadString(reader, 6),
            ModelNo = ReadString(reader, 7),
            SerialNo = ReadString(reader, 8),
            Manufacturer = ReadString(reader, 9),
            Notes = ReadString(reader, 10),
            EmployeeId = ReadString(reader, 11),
            DepartmentId = ReadString(reader, 12),
            VendorAccountNo = ReadString(reader, 13),
            Status = Convert.ToByte(reader.GetValue(14)),
            PlacedInServiceDate = ReadDate(reader, 15),
            AcquisitionCost = ReadDecimal(reader, 16),
            AccumulatedDepreciation = ReadDecimal(reader, 17),
            BookValue = ReadDecimal(reader, 18),
            RequiresCalibration = ReadBool(reader, 19),
            CalibrationFrequency = ReadNullableInt(reader, 20),
            AccumulatedAccountNo = ReadString(reader, 21),
            DepreciationAccountNo = ReadString(reader, 22),
            DepreciationMethodId = reader.IsDBNull(23) ? null : Convert.ToByte(reader.GetValue(23)),
            Convention = ReadString(reader, 24),
            DepreciationRate = ReadDecimal(reader, 25),
            DepreciationBeginMonth = ReadDate(reader, 26),
            LifeInYears = ReadNullableInt(reader, 27),
            InActive = ReadBool(reader, 28)
        };

        if (model.BookValue == 0)
        {
            model.BookValue = model.AcquisitionCost - model.AccumulatedDepreciation;
        }

        return model;
    }

    private static void Validate(FixedAssetProfileModel model)
    {
        if (string.IsNullOrWhiteSpace(model.AssetId)) throw new InvalidOperationException("Asset ID is required.");
        if (string.IsNullOrWhiteSpace(model.Description)) throw new InvalidOperationException("Description is required.");
        if (model.AcquisitionCost < 0) throw new InvalidOperationException("Acquisition cost cannot be negative.");
        if (model.AccumulatedDepreciation < 0) throw new InvalidOperationException("Accumulated depreciation cannot be negative.");
        model.BookValue = model.AcquisitionCost - model.AccumulatedDepreciation;
    }

    private SqlConnection CreateConnection() => new(_connectionString);
    private static object DbValue<T>(T? value) where T : struct => value.HasValue ? value.Value : DBNull.Value;
    private static object DbValue(DateTime? value) => value.HasValue ? value.Value : DBNull.Value;
    private static string ReadString(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? "" : Convert.ToString(reader.GetValue(ordinal)) ?? "";
    private static int ReadInt(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? 0 : Convert.ToInt32(reader.GetValue(ordinal));
    private static int? ReadNullableInt(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : Convert.ToInt32(reader.GetValue(ordinal));
    private static DateTime? ReadDate(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : Convert.ToDateTime(reader.GetValue(ordinal));
    private static decimal ReadDecimal(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? 0 : Convert.ToDecimal(reader.GetValue(ordinal));
    private static bool ReadBool(SqlDataReader reader, int ordinal) => !reader.IsDBNull(ordinal) && Convert.ToBoolean(reader.GetValue(ordinal));
    private static string TrimTo(string? value, int length)
    {
        value = (value ?? "").Trim();
        return value.Length <= length ? value : value[..length];
    }
}
