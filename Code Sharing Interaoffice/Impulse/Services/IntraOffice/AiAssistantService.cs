using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace Impulse.Services.IntraOffice
{
    public class ChatMessageDto
    {
        public string Role { get; set; } = "user"; // "user" or "assistant"
        public string Content { get; set; } = "";
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }

    public interface IAiAssistantService
    {
        string GetAiUserId();
        bool IsAiUser(string userId);
        Task<string> GenerateResponseAsync(string userMessage, string userId);
        Task<string> GetEnterpriseChatResponseAsync(string userMessage, string userId, System.Collections.Generic.List<ChatMessageDto>? history = null);
    }

    public class AiAssistantService : IAiAssistantService
    {
        private const string AiUserId = "IntraBot";
        private readonly IConfiguration _config;
        private readonly ILogger<AiAssistantService> _logger;

        public AiAssistantService(IConfiguration config, ILogger<AiAssistantService> logger)
        {
            _config = config;
            _logger = logger;
        }

        public string GetAiUserId() => AiUserId;
        public bool IsAiUser(string userId) => string.Equals(userId, AiUserId, StringComparison.OrdinalIgnoreCase);

        public Task<string> GetEnterpriseChatResponseAsync(string userMessage, string userId, System.Collections.Generic.List<ChatMessageDto>? history = null)
        {
            return GenerateResponseAsync(userMessage, userId);
        }

        public async Task<string> GenerateResponseAsync(string userMessage, string userId)
        {
            var apiKey = _config["AiAssistant:ApiKey"]?.Trim();
            var endpoint = _config["AiAssistant:Endpoint"] ?? "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent";

            if (string.IsNullOrEmpty(apiKey) || apiKey == "YOUR_API_KEY_HERE")
            {
                await Task.Delay(500);
                return GenerateMockResponse(userMessage);
            }

            try
            {
                using var client = new HttpClient();
                var requestUrl = $"{endpoint}?key={apiKey}";

                var payload = new
                {
                    systemInstruction = new
                    {
                        parts = new[] { new { text = "You are IntraBot, a helpful AI assistant for Impulse ERP. Be concise, professional, and friendly." } }
                    },
                    contents = new[]
                    {
                        new
                        {
                            parts = new[] { new { text = userMessage } }
                        }
                    }
                };

                var content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
                var response = await client.PostAsync(requestUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    using var doc = JsonDocument.Parse(json);
                    var reply = doc.RootElement
                        .GetProperty("candidates")[0]
                        .GetProperty("content")
                        .GetProperty("parts")[0]
                        .GetProperty("text")
                        .GetString();

                    return reply ?? "I am sorry, I couldn't generate a response.";
                }

                return GenerateMockResponse(userMessage);
            }
            catch (Exception ex)
            {
                _logger.LogWarning("AI Service error: {Message}", ex.Message);
                return GenerateMockResponse(userMessage);
            }
        }

        private string GenerateMockResponse(string userMessage)
        {
            var msg = userMessage.ToLower();
            if (msg.Contains("hello") || msg.Contains("hi"))
                return "Hello! I'm IntraBot, your Impulse ERP assistant. How can I assist you today?";
            if (msg.Contains("task") || msg.Contains("todo"))
                return "You can manage your tasks directly in the Task Board under the IntraOffice module.";
            if (msg.Contains("meeting"))
                return "You can view and schedule enterprise meetings from the Meetings dashboard.";
            if (msg.Contains("minute") || msg.Contains("memo"))
                return "Minutes and approval workflows can be tracked under Minutes Approval.";
            
            return $"Thank you for your message! I'm IntraBot. (Configured in assistant mode for Impulse ERP). You said: \"{userMessage}\"";
        }
    }
}
