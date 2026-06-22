using System;

namespace DataAccessLibrary.Models.ReportModels
{
    // These records act as the "return value" from your modals
    public record DateRangeResult(bool Success, DateTime DateFrom, DateTime DateTo);
    public record DateSingleResult(bool Success, DateTime Date);

    public record DateRangeSingleSelectionResult(bool Success, DateTime DateFrom, DateTime DateTo, string SelectedValue);
    public record DateRangeSingleSelectionResultSubHeads(bool Success, DateTime DateFrom, DateTime DateTo, string SelectedValue, string SelectedValue_sub);

    public class ReportOption
    {
        public string Text { get; set; }
        public string Value { get; set; }
    }

}