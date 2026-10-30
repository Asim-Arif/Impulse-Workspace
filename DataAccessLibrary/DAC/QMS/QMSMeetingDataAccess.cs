using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using DataAccessLibrary.Interface.QMS;
using DataAccessLibrary.Models.QMS;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DataAccessLibrary.DAC.QMS
{
    public class QMSMeetingDataAccess : IQMSMeetingDataAccess
    {
        private readonly string _connectionString;
        private readonly ILogger<QMSMeetingDataAccess> _logger;

        public QMSMeetingDataAccess(IConfiguration configuration, ILogger<QMSMeetingDataAccess> logger)
        {
            _connectionString = configuration.GetConnectionString("DefaultConnection")
                ?? configuration.GetConnectionString("ImpulseDatabase")
                ?? configuration.GetConnectionString("ImpulseConnection")
                ?? throw new InvalidOperationException("Connection string not found.");
            _logger = logger;
        }

        private IDbConnection CreateConnection() => new SqlConnection(_connectionString);

        public async Task<List<QMSMeetingListItemModel>> GetMeetingsAsync(QMSMeetingFilterDto filter)
        {
            try
            {
                using var db = CreateConnection();
                var sql = new StringBuilder(@"
                    SELECT 
                        m.EntryID,
                        m.Code,
                        m.Type,
                        m.DT,
                        m.Place,
                        m.Duration,
                        m.Agenda,
                        m.Discussion,
                        m.Keywords,
                        m.UserName,
                        m.DTEntry,
                        (SELECT COUNT(1) FROM QMS_Meetings_Attendees a WHERE a.RefID = m.EntryID) AS AttendeeCount,
                        (SELECT COUNT(1) FROM EntityAttachments ea WHERE ea.EntityType = 'QMSMeeting' AND ea.EntityId = CAST(m.EntryID AS VARCHAR(50))) AS AttachmentCount
                    FROM QMS_Meetings m
                    WHERE 1 = 1");

                var parameters = new DynamicParameters();

                if (filter != null)
                {
                    if (filter.DateFrom.HasValue && filter.DateTo.HasValue)
                    {
                        sql.Append(" AND m.DT >= @DateFrom AND m.DT <= @DateTo");
                        parameters.Add("DateFrom", filter.DateFrom.Value.Date);
                        parameters.Add("DateTo", filter.DateTo.Value.Date.AddDays(1).AddTicks(-1));
                    }

                    if (!string.IsNullOrWhiteSpace(filter.MeetingType))
                    {
                        sql.Append(" AND m.Type = @MeetingType");
                        parameters.Add("MeetingType", filter.MeetingType.Trim());
                    }

                    if (!string.IsNullOrWhiteSpace(filter.SearchText))
                    {
                        sql.Append(@" AND (
                            m.Code LIKE @Search 
                            OR m.Type LIKE @Search 
                            OR m.Place LIKE @Search 
                            OR m.Agenda LIKE @Search 
                            OR m.Discussion LIKE @Search 
                            OR m.Keywords LIKE @Search
                            OR EXISTS (SELECT 1 FROM QMS_Meetings_Attendees att WHERE att.RefID = m.EntryID AND att.AttendeeName LIKE @Search)
                        )");
                        parameters.Add("Search", $"%{filter.SearchText.Trim()}%");
                    }
                }

                sql.Append(" ORDER BY m.DT DESC, m.EntryID DESC");

                var results = await db.QueryAsync<QMSMeetingListItemModel>(sql.ToString(), parameters);
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching QMS meetings list.");
                throw;
            }
        }

        public async Task<QMSMeetingModel?> GetMeetingByIdAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        m.EntryID,
                        m.Code,
                        m.Type,
                        m.DT,
                        m.Place,
                        m.Duration,
                        m.Agenda,
                        m.Discussion,
                        m.Keywords,
                        m.UserName,
                        m.MachineName,
                        m.DTEntry
                    FROM QMS_Meetings m
                    WHERE m.EntryID = @EntryID";

                var meeting = await db.QueryFirstOrDefaultAsync<QMSMeetingModel>(sql, new { EntryID = entryId });
                if (meeting == null) return null;

                const string attSql = @"
                    SELECT AttendeeName
                    FROM QMS_Meetings_Attendees
                    WHERE RefID = @RefID
                    ORDER BY EntryID";

                var attendees = await db.QueryAsync<string>(attSql, new { RefID = entryId });
                meeting.Attendees = attendees.ToList();

                return meeting;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching QMS meeting by ID {EntryID}.", entryId);
                throw;
            }
        }

        public async Task<int> SaveMeetingAsync(QMSMeetingModel model, string currentUserName, string? machineName)
        {
            try
            {
                using var db = CreateConnection();
                db.Open();
                using var trans = db.BeginTransaction();

                int meetingId = model.EntryID;

                if (meetingId == 0)
                {
                    if (string.IsNullOrWhiteSpace(model.Code))
                    {
                        model.Code = await GenerateNextCodeAsync(db, trans);
                    }

                    const string insertSql = @"
                        INSERT INTO QMS_Meetings (
                            Code,
                            Type,
                            DT,
                            Place,
                            Duration,
                            Agenda,
                            Discussion,
                            Keywords,
                            UserName,
                            MachineName,
                            DTEntry
                        ) VALUES (
                            @Code,
                            @Type,
                            @DT,
                            @Place,
                            @Duration,
                            @Agenda,
                            @Discussion,
                            @Keywords,
                            @UserName,
                            @MachineName,
                            GETDATE()
                        );
                        SELECT CAST(SCOPE_IDENTITY() AS INT);";

                    meetingId = await db.ExecuteScalarAsync<int>(insertSql, new
                    {
                        model.Code,
                        model.Type,
                        model.DT,
                        model.Place,
                        model.Duration,
                        model.Agenda,
                        model.Discussion,
                        model.Keywords,
                        UserName = currentUserName,
                        MachineName = machineName ?? Environment.MachineName
                    }, trans);

                    model.EntryID = meetingId;
                }
                else
                {
                    const string updateSql = @"
                        UPDATE QMS_Meetings
                        SET 
                            Code = @Code,
                            Type = @Type,
                            DT = @DT,
                            Place = @Place,
                            Duration = @Duration,
                            Agenda = @Agenda,
                            Discussion = @Discussion,
                            Keywords = @Keywords,
                            UserName = @UserName,
                            MachineName = @MachineName
                        WHERE EntryID = @EntryID;";

                    await db.ExecuteAsync(updateSql, new
                    {
                        model.EntryID,
                        model.Code,
                        model.Type,
                        model.DT,
                        model.Place,
                        model.Duration,
                        model.Agenda,
                        model.Discussion,
                        model.Keywords,
                        UserName = currentUserName,
                        MachineName = machineName ?? Environment.MachineName
                    }, trans);
                }

                // Delete and re-insert attendees
                await db.ExecuteAsync("DELETE FROM QMS_Meetings_Attendees WHERE RefID = @RefID", new { RefID = meetingId }, trans);

                if (model.Attendees != null && model.Attendees.Any())
                {
                    const string insertAttSql = @"
                        INSERT INTO QMS_Meetings_Attendees (RefID, AttendeeName)
                        VALUES (@RefID, @AttendeeName);";

                    foreach (var att in model.Attendees)
                    {
                        if (!string.IsNullOrWhiteSpace(att))
                        {
                            await db.ExecuteAsync(insertAttSql, new
                            {
                                RefID = meetingId,
                                AttendeeName = att.Trim()
                            }, trans);
                        }
                    }
                }

                trans.Commit();
                return meetingId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving QMS meeting.");
                throw;
            }
        }

        public async Task<bool> DeleteMeetingAsync(int entryId)
        {
            try
            {
                using var db = CreateConnection();
                db.Open();
                using var trans = db.BeginTransaction();

                // Delete attendees
                await db.ExecuteAsync("DELETE FROM QMS_Meetings_Attendees WHERE RefID = @EntryID", new { EntryID = entryId }, trans);

                // Delete attachments
                await db.ExecuteAsync("DELETE FROM EntityAttachments WHERE EntityType = 'QMSMeeting' AND EntityId = @EntityId", new { EntityId = entryId.ToString() }, trans);

                // Delete meeting record
                int rows = await db.ExecuteAsync("DELETE FROM QMS_Meetings WHERE EntryID = @EntryID", new { EntryID = entryId }, trans);

                trans.Commit();
                return rows > 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting QMS meeting with ID {EntryID}.", entryId);
                throw;
            }
        }

        public async Task<string> GetNextMeetingCodeAsync()
        {
            try
            {
                using var db = CreateConnection();
                return await GenerateNextCodeAsync(db, null);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting next meeting code.");
                return "0001";
            }
        }

        private async Task<string> GenerateNextCodeAsync(IDbConnection db, IDbTransaction? trans)
        {
            const string sql = @"
                SELECT ISNULL(MAX(CAST(Code AS INT)), 0) + 1 
                FROM QMS_Meetings 
                WHERE ISNUMERIC(Code) = 1;";

            int nextNum = await db.ExecuteScalarAsync<int>(sql, null, trans);
            return nextNum.ToString("D4");
        }

        public async Task<List<string>> GetMeetingTypesAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = "SELECT DISTINCT MeetingType FROM QMS_Meetings_Types WHERE MeetingType IS NOT NULL ORDER BY MeetingType";
                var results = await db.QueryAsync<string>(sql);
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error loading meeting types.");
                return new List<string>();
            }
        }

        public async Task<List<EmployeeDropdownDto>> GetEmployeesAsync()
        {
            try
            {
                using var db = CreateConnection();
                const string sql = @"
                    SELECT 
                        empid AS EmpID,
                        name AS Name,
                        fname AS FName,
                        DeptName,
                        Designation
                    FROM VEmp
                    WHERE (Active = 1 OR Active IS NULL)
                      AND empid IS NOT NULL AND RTRIM(LTRIM(empid)) <> ''
                      AND name IS NOT NULL AND RTRIM(LTRIM(name)) <> ''
                    ORDER BY name";

                var results = await db.QueryAsync<EmployeeDropdownDto>(sql);
                return results.ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error loading employees lookup for QMS meetings.");
                return new List<EmployeeDropdownDto>();
            }
        }
    }
}

