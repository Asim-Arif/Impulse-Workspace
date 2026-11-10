using System;
using System.Collections.Generic;

namespace DataAccessLibrary.Models.ViewModels.Gate;

public sealed class GateDashboardModel
{
    public GateDashboardSection Today { get; set; } = new();
    public GateDashboardSection LastSevenDays { get; set; } = new();
    public IReadOnlyList<GateDashboardDayModel> LastSevenDayTrend { get; set; } = Array.Empty<GateDashboardDayModel>();
}

public sealed class GateDashboardSection
{
    public int Visitors { get; set; }
    public int GateIn { get; set; }
    public int GateOutPermanent { get; set; }
    public int GateOutReturnable { get; set; }
    public int Courier { get; set; }
    public int Vehicles { get; set; }

    // Backward compatibility for any existing callers
    public int GateOut
    {
        get => GateOutPermanent + GateOutReturnable;
        set => GateOutPermanent = value;
    }

    public int Total => Visitors + GateIn + GateOutPermanent + GateOutReturnable + Courier + Vehicles;
}

public sealed class GateDashboardDayModel
{
    public DateTime Date { get; set; }
    public int Visitors { get; set; }
    public int GateIn { get; set; }
    public int GateOutPermanent { get; set; }
    public int GateOutReturnable { get; set; }
    public int Courier { get; set; }
    public int Vehicles { get; set; }

    // Backward compatibility
    public int GateOut
    {
        get => GateOutPermanent + GateOutReturnable;
        set => GateOutPermanent = value;
    }

    public int Total => Visitors + GateIn + GateOutPermanent + GateOutReturnable + Courier + Vehicles;
}
