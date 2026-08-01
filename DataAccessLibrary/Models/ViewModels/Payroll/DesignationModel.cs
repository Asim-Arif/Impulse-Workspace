using System;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class DesignationModel
    {
        public int EntryID { get; set; }
        public string Designation { get; set; } = string.Empty;
        public string Duties { get; set; } = string.Empty;
        public string RequiredSkills { get; set; } = string.Empty;
        public string RequiredTrainings { get; set; } = string.Empty;
    }
}
