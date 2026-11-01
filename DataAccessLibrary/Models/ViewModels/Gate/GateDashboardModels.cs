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
    public int GateOut { get; set; }
    public int Total => Visitors + GateIn + GateOut;
}

public sealed class GateDashboardDayModel
{
    public DateTime Date { get; set; }
    public int Visitors { get; set; }
    public int GateIn { get; set; }
    public int GateOut { get; set; }
    public int Total => Visitors + GateIn + GateOut;
}

