using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace DataAccessLibrary.Models.ViewModels.Gate;

public sealed class MaterialInwardFilterModel
{
    public string SearchText { get; set; } = "";
    public bool ShowClosed { get; set; }
}

public sealed class MaterialInwardListRowModel
{
    public int EntryID { get; set; }
    public string PoNumber { get; set; } = "";
    public string SupplierName { get; set; } = "";
    public string VehicleNumber { get; set; } = "";
    public string DriverName { get; set; } = "";
    public decimal Weighment { get; set; }
    public DateTime EntryTime { get; set; }
    public string Status { get; set; } = "";
    public int LineCount { get; set; }
}

public sealed class MaterialInwardEntryModel
{
    public int EntryID { get; set; }

    [StringLength(50)]
    public string PoNumber { get; set; } = "";

    [Required, StringLength(150)]
    public string SupplierName { get; set; } = "";

    public int? VehicleEntryId { get; set; }

    [StringLength(50)]
    public string VehicleNumber { get; set; } = "";

    [StringLength(100)]
    public string DriverName { get; set; } = "";

    [StringLength(50)]
    public string DeliveryChallanNo { get; set; } = "";

    [StringLength(50)]
    public string InvoiceNo { get; set; } = "";

    public decimal? Weighment { get; set; }
    public DateTime EntryTime { get; set; } = DateTime.Now;
    public string Status { get; set; } = "Open";

    [StringLength(255)]
    public string Remarks { get; set; } = "";

    public List<MaterialInwardLineModel> Lines { get; set; } = new();
}

public sealed class MaterialInwardLineModel
{
    public int EntryID { get; set; }

    [Required, StringLength(150)]
    public string MaterialName { get; set; } = "";

    [Range(0.001, 999999999)]
    public decimal Quantity { get; set; }

    [StringLength(25)]
    public string Uom { get; set; } = "";
}

public sealed class MaterialOutwardFilterModel
{
    public string SearchText { get; set; } = "";
    public bool ShowClosed { get; set; } = true;
}

public sealed class MaterialOutwardListRowModel
{
    public int EntryID { get; set; }
    public string SalesOrderNo { get; set; } = "";
    public string DeliveryOrderNo { get; set; } = "";
    public string CustomerName { get; set; } = "";
    public string VehicleNumber { get; set; } = "";
    public string DriverName { get; set; } = "";
    public decimal Weighment { get; set; }
    public DateTime EntryTime { get; set; }
    public string Status { get; set; } = "";
    public int LineCount { get; set; }
}

public sealed class MaterialOutwardEntryModel
{
    public int EntryID { get; set; }

    [StringLength(50)]
    public string SalesOrderNo { get; set; } = "";

    [StringLength(50)]
    public string DeliveryOrderNo { get; set; } = "";

    [Required, StringLength(150)]
    public string CustomerName { get; set; } = "";

    public int? VehicleEntryId { get; set; }

    [StringLength(50)]
    public string VehicleNumber { get; set; } = "";

    [StringLength(100)]
    public string DriverName { get; set; } = "";

    public decimal? Weighment { get; set; }
    public DateTime EntryTime { get; set; } = DateTime.Now;
    public string Status { get; set; } = "Dispatched";

    [StringLength(255)]
    public string Remarks { get; set; } = "";

    public List<MaterialOutwardLineModel> Lines { get; set; } = new();
}

public sealed class MaterialOutwardLineModel
{
    public int EntryID { get; set; }

    [Required, StringLength(150)]
    public string MaterialName { get; set; } = "";

    [Range(0.001, 999999999)]
    public decimal Quantity { get; set; }

    [StringLength(25)]
    public string Uom { get; set; } = "";
}

