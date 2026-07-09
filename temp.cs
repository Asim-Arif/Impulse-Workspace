        private bool ShowUsageModal { get; set; }
        private List<AdvancePaymentUsageViewModel>? UsageDetails { get; set; }

        private async Task ShowUsageDetail(AdvancePaymentListViewModel item)
        {
            ShowUsageModal = true;
            UsageDetails = null; // show loading
            StateHasChanged();

            try
            {
                UsageDetails = await AdvancePaymentService.GetUsageDetailsAsync(item.EntryID);
            }
            catch (Exception ex)
            {
                NotificationService.Notify(new NotificationMessage { Severity = NotificationSeverity.Error, Summary = "Error", Detail = ex.Message });
                UsageDetails = new List<AdvancePaymentUsageViewModel>();
            }
        }

        private void CloseUsageModal()
        {
            ShowUsageModal = false;
        }
