using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Payroll
{
    public class GeneralSettingsModel
    {
        public int Entryid { get; set; }
        public string DeptID { get; set; } = string.Empty;
        public string EnterTime { get; set; } = "08:00";
        public string ExitTime { get; set; } = "17:00";
        public double LunchTime { get; set; } = 1.0;
        public string LunchInTime { get; set; } = "13:00";
        public string LunchOutTime { get; set; } = "14:00";
        public string FridayBreakStartTime { get; set; } = "12:30";
        public string FridayBreakEndTime { get; set; } = "14:00";
        public string FridayCheckOut { get; set; } = "17:30";
        public double WorkingHrs { get; set; } = 8.0;
        public int EmpAttCat_RefID { get; set; } = 1;

        // Legacy / Hidden Settings Columns
        public double AbsAllowed { get; set; } = 0;
        public double TotalAbsAllowed { get; set; } = 0;
        public double TimeAllowed { get; set; } = 0;
        public double NTimes { get; set; } = 0;
        public double OverTime { get; set; } = 5;
        public double SocialAmt { get; set; } = 0;
        public double DinnerAmt { get; set; } = 0;
        public double ShortTermPer { get; set; } = 0;
        public double AdvSalPer { get; set; } = 0;
        public double LongTermTime { get; set; } = 0;
        public string LongTermLimit { get; set; } = "0";
        public bool BGHoliday { get; set; } = false;
        public bool AGHoliday { get; set; } = false;
        public bool BPHoliday { get; set; } = false;
        public bool APHoliday { get; set; } = false;
        public bool DeductAbsents { get; set; } = true;
        public bool DeductLHrs { get; set; } = true;
    }

    public class EmpAttCategoryModel
    {
        public int EntryID { get; set; }
        public bool? LunchTimeScanning { get; set; }
        public bool? LateDeduction { get; set; }
        public string? LateDeductionAfter { get; set; }
        public bool? GatePassRequired { get; set; }
        public bool? LeavesRecorded { get; set; }
        public bool? OverTimePaid { get; set; }
        public bool? DeductionSpecialAbsents { get; set; }
        public float? OTRate { get; set; }
        public int? DeductionPercentage { get; set; }
        public string? Remarks { get; set; }
        public int Saturday_Early_Out_Mins { get; set; } = 0;
    }

    public class CategoryRuleItem
    {
        public string SettingName { get; set; } = string.Empty;
        public string SettingValue { get; set; } = string.Empty;
    }
}
