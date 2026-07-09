using System.Collections.Generic;
using System.Linq;
using DataAccessLibrary.Models.ViewModels.Export;
using Microsoft.AspNetCore.Components;

namespace Impulse.Components.Export.NewProformaInvoiceTabs
{
    public partial class OtherChargesTab
    {
        [Parameter] public List<ProformaOtherChargeModel> OtherCharges { get; set; } = new();
        [Parameter] public EventCallback OnChargesUpdated { get; set; }

        private ProformaOtherChargeModel NewCharge { get; set; } = new();

        private void AddCharge()
        {
            if (string.IsNullOrWhiteSpace(NewCharge.ChargesDetail) || NewCharge.ChargesAmt <= 0)
                return;

            NewCharge.SNo = OtherCharges.Count > 0 ? OtherCharges.Max(c => c.SNo) + 1 : 1;
            OtherCharges.Add(new ProformaOtherChargeModel
            {
                SNo = NewCharge.SNo,
                ChargesDetail = NewCharge.ChargesDetail,
                ChargesAmt = NewCharge.ChargesAmt
            });

            NewCharge = new ProformaOtherChargeModel();
            OnChargesUpdated.InvokeAsync();
        }

        private void RemoveCharge(ProformaOtherChargeModel charge)
        {
            OtherCharges.Remove(charge);
            
            // Reassign SNo
            for (int i = 0; i < OtherCharges.Count; i++)
            {
                OtherCharges[i].SNo = i + 1;
            }
            
            OnChargesUpdated.InvokeAsync();
        }
    }
}
