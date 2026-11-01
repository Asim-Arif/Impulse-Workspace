using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace DataAccessLibrary.Models.ViewModels.Gate;

public sealed class VehicleFilterModel
{
    public string SearchText { get; set; } = "";
    public int CategoryId { get; set; }
    public bool ShowClosed { get; set; }
}

public sealed class VehicleCategoryModel
{
    public int EntryID { get; set; }
    public string CategoryName { get; set; } = "";
    public int SortOrder { get; set; }
    public bool Active { get; set; } = true;
}

public sealed class VehicleEntryListRowModel
{
    public int EntryID { get; set; }
    public string VehicleNumber { get; set; } = "";
    public string VehicleType { get; set; } = "";
    public string DriverName { get; set; } = "";
    public string SupplierCompany { get; set; } = "";
    public string Transporter { get; set; } = "";
    public string CategoryName { get; set; } = "";
    public DateTime EntryTime { get; set; }
    public DateTime? ExitTime { get; set; }
    public string Status { get; set; } = "";
}

public sealed class VehicleEntryModel
{
    public int EntryID { get; set; }

    [Required, StringLength(50)]
    public string VehicleNumber { get; set; } = "";

    [StringLength(50)]
    public string VehicleType { get; set; } = "";

    [StringLength(100)]
    public string DriverName { get; set; } = "";

    [StringLength(50)]
    public string DriverCNIC { get; set; } = "";

    [StringLength(50)]
    public string DriverMobile { get; set; } = "";

    [StringLength(150)]
    public string SupplierCompany { get; set; } = "";

    [StringLength(150)]
    public string Transporter { get; set; } = "";

    [StringLength(255)]
    public string Purpose { get; set; } = "";

    public int? CategoryId { get; set; }
    public string CategoryName { get; set; } = "";
    public DateTime EntryTime { get; set; } = DateTime.Now;
    public DateTime? ExitTime { get; set; }
    public string? VehiclePhotoDataUrl { get; set; }
    public string Status { get; set; } = "Inside";

    [StringLength(255)]
    public string Remarks { get; set; } = "";
}

