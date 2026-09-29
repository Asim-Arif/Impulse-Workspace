using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DataAccessLibrary.Interface.Production;
using DataAccessLibrary.Models.ViewModels.Production;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.Production
{
    public class EmployeePerformanceService : IEmployeePerformanceService
    {
        private readonly IEmployeePerformanceDataAccess _dataAccess;
        private readonly ILogger<EmployeePerformanceService> _logger;

        public EmployeePerformanceService(IEmployeePerformanceDataAccess dataAccess, ILogger<EmployeePerformanceService> logger)
        {
            _dataAccess = dataAccess;
            _logger = logger;
        }

        public async Task<List<DepartmentLookupItem>> GetDepartmentsLookupAsync()
        {
            try
            {
                return await _dataAccess.GetDepartmentsLookupAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching departments for Employee Performance");
                return new List<DepartmentLookupItem>();
            }
        }

        public async Task<List<ProcessLookupItem>> GetProcessesLookupAsync()
        {
            try
            {
                return await _dataAccess.GetProcessesLookupAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching processes for Employee Performance");
                return new List<ProcessLookupItem>();
            }
        }

        public async Task<List<PerformanceEmployeeLookupItem>> GetEmployeesLookupAsync(int? deptId = null)
        {
            try
            {
                return await _dataAccess.GetEmployeesLookupAsync(deptId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching employees lookup for Employee Performance");
                return new List<PerformanceEmployeeLookupItem>();
            }
        }

        public async Task<EmployeePerformanceDashboardDto> GetPerformanceDashboardAsync(EmployeePerformanceFilter filter)
        {
            try
            {
                return await _dataAccess.GetPerformanceDashboardAsync(filter);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching Employee Performance dashboard");
                throw;
            }
        }

        public async Task<List<EmployeeLotDetailDto>> GetEmployeeLotsDetailAsync(string empId, DateTime dtFrom, DateTime dtTo)
        {
            try
            {
                return await _dataAccess.GetEmployeeLotsDetailAsync(empId, dtFrom, dtTo);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching employee lots detail for EmpID [{EmpID}]", empId);
                return new List<EmployeeLotDetailDto>();
            }
        }
    }
}
