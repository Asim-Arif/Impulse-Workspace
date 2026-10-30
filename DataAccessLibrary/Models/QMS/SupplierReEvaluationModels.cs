using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.QMS
{
    public class SupplierReEvaluationModel
    {
        public int EntryID { get; set; }
        public string AccNo { get; set; } = string.Empty;
        public string? AccTitle { get; set; }
        public DateTime? ReviewDT { get; set; } = DateTime.Today;
        public DateTime? EntryDT { get; set; } = DateTime.Today;
        public string? ReviewBy { get; set; }
        public string? PreparedBy { get; set; }
        public string? ApprovedBy { get; set; }
        public string? UserName { get; set; }
        public string? MachineName { get; set; }

        // Section 1: Quality Criteria
        public int? QualityOfProduction { get; set; }
        public string? QualityOfProductionComments { get; set; }

        public int? Price { get; set; }
        public string? PriceComments { get; set; }

        public int? Packing { get; set; }
        public string? PackingComments { get; set; }

        public int? RejectionRate { get; set; }
        public string? RejectionRateComments { get; set; }

        public int? DeliveryInTime { get; set; }
        public string? DeliveryInTimeComments { get; set; }

        public int? ReWorkRate { get; set; }
        public string? ReWorkRateComments { get; set; }

        public int? NoOfOrdersCompleted { get; set; }
        public string? NoOfOrdersCompletedComments { get; set; }

        public int? CancelledOrders { get; set; }
        public string? CancelledOrdersComments { get; set; }

        public int? ResponseTime { get; set; }
        public string? ResponseTimeComments { get; set; }

        // Section 2: Performance Criteria
        public int? TechnicalAbility { get; set; }
        public string? TechnicalAbilityComments { get; set; }

        public int? ProductionCapacity { get; set; }
        public string? ProductionCapacityComments { get; set; }

        // Section 3: Environment System Criteria
        public int? EnvironmentRequirements { get; set; }
        public string? EnvironmentRequirementsComments { get; set; }

        public int? LightingEtc { get; set; }
        public string? LightingEtcComments { get; set; }

        // Calculation Helpers
        public int TotalScore =>
            (QualityOfProduction ?? 0) +
            (Price ?? 0) +
            (Packing ?? 0) +
            (RejectionRate ?? 0) +
            (DeliveryInTime ?? 0) +
            (ReWorkRate ?? 0) +
            (NoOfOrdersCompleted ?? 0) +
            (CancelledOrders ?? 0) +
            (ResponseTime ?? 0) +
            (TechnicalAbility ?? 0) +
            (ProductionCapacity ?? 0) +
            (EnvironmentRequirements ?? 0) +
            (LightingEtc ?? 0);

        public int MaxScore => 130; // 13 criteria * 10 max points each

        public double ScorePercentage => MaxScore > 0 ? Math.Round((double)TotalScore / MaxScore * 100.0, 1) : 0;
    }

    public class SupplierReEvaluationListItemDto
    {
        public int EntryID { get; set; }
        public string AccNo { get; set; } = string.Empty;
        public string AccTitle { get; set; } = string.Empty;
        public DateTime? ReviewDT { get; set; }
        public DateTime? EntryDT { get; set; }
        public string? ReviewBy { get; set; }
        public string? PreparedBy { get; set; }
        public string? ApprovedBy { get; set; }

        public int QualityOfProduction { get; set; }
        public int Price { get; set; }
        public int Packing { get; set; }
        public int RejectionRate { get; set; }
        public int DeliveryInTime { get; set; }
        public int ReWorkRate { get; set; }
        public int NoOfOrdersCompleted { get; set; }
        public int CancelledOrders { get; set; }
        public int ResponseTime { get; set; }
        public int TechnicalAbility { get; set; }
        public int ProductionCapacity { get; set; }
        public int EnvironmentRequirements { get; set; }
        public int LightingEtc { get; set; }

        public int TotalScore =>
            QualityOfProduction + Price + Packing + RejectionRate + DeliveryInTime +
            ReWorkRate + NoOfOrdersCompleted + CancelledOrders + ResponseTime +
            TechnicalAbility + ProductionCapacity + EnvironmentRequirements + LightingEtc;

        public int MaxScore => 130;

        public double ScorePercentage => MaxScore > 0 ? Math.Round((double)TotalScore / MaxScore * 100.0, 1) : 0;

        public string ScoreBadgeClass => ScorePercentage switch
        {
            >= 80 => "bg-success-subtle text-success border border-success-subtle",
            >= 60 => "bg-primary-subtle text-primary border border-primary-subtle",
            >= 40 => "bg-warning-subtle text-warning-emphasis border border-warning-subtle",
            _ => "bg-danger-subtle text-danger border border-danger-subtle"
        };

        public int AttachmentCount { get; set; }
    }

    public class SupplierReEvaluationFilterDto
    {
        public DateTime? DateFrom { get; set; }
        public DateTime? DateTo { get; set; }
        public string? AccNo { get; set; }
        public string? SearchText { get; set; }
    }
}

