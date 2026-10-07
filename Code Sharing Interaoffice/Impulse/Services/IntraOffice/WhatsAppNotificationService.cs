using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.IntraOffice
{
    public interface IWhatsAppNotificationService
    {
        Task<bool> SendTaskNotificationAsync(string phoneNumber, string taskTitle, string assigneeName, string priority, string? description = null, List<string>? attachments = null);
        Task<bool> SendDirectMessageAsync(string phoneNumber, string message, List<string>? attachments = null);
    }

    public class WhatsAppNotificationService : IWhatsAppNotificationService
    {
        private readonly HttpClient _httpClient;
        private readonly string _nodeApiUrl;
        private readonly ILogger<WhatsAppNotificationService> _logger;

        public WhatsAppNotificationService(HttpClient httpClient, IConfiguration configuration, ILogger<WhatsAppNotificationService> logger)
        {
            _httpClient = httpClient;
            _httpClient.Timeout = TimeSpan.FromSeconds(5); // Fast timeout so UI never hangs
            _nodeApiUrl = configuration["IntraOffice:WhatsAppNodeUrl"] ?? "http://127.0.0.1:3001/send";
            _logger = logger;
        }

        public async Task<bool> SendTaskNotificationAsync(string phoneNumber, string taskTitle, string assigneeName, string priority, string? description = null, List<string>? attachments = null)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(phoneNumber)) return false;

                var descText = string.IsNullOrWhiteSpace(description) ? "" : $"\nDescription: {description}\n";
                var message = $"*New Task Assigned*\n\nHello {assigneeName},\nYou have been assigned a new task: *{taskTitle}*\nPriority: {priority}\n{descText}\nPlease check your Impulse ERP dashboard for details.";

                var payload = new
                {
                    phoneNumber = phoneNumber,
                    message = message,
                    attachments = attachments ?? new List<string>()
                };

                var content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync(_nodeApiUrl, content);
                return response.IsSuccessStatusCode;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("WhatsApp notification service unreachable or error: {Message}", ex.Message);
                return false;
            }
        }

        public async Task<bool> SendDirectMessageAsync(string phoneNumber, string message, List<string>? attachments = null)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(phoneNumber)) return false;

                var payload = new
                {
                    phoneNumber = phoneNumber,
                    message = message,
                    attachments = attachments ?? new List<string>()
                };

                var content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync(_nodeApiUrl, content);
                return response.IsSuccessStatusCode;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("WhatsApp direct message service unreachable or error: {Message}", ex.Message);
                return false;
            }
        }
    }
}
