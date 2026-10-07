using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;

namespace Impulse.Services.IntraOffice
{
    public class OpenRouterAiService : IAiService
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiKey;

        public OpenRouterAiService(HttpClient httpClient, IConfiguration configuration)
        {
            _httpClient = httpClient;
            _apiKey = configuration["OpenRouterApiKey"] ?? Environment.GetEnvironmentVariable("OPENROUTER_API_KEY") ?? string.Empty;
        }

        public async Task<string> GenerateResponseAsync(string prompt, string systemMessage = "You are a helpful AI assistant built into the IntraOffice corporate intranet.")
        {
            return await GenerateConversationResponseAsync(new List<(string role, string content)> { ("user", prompt) }, systemMessage);
        }

        public async Task<string> GenerateConversationResponseAsync(List<(string role, string content)> messages, string systemMessage = "You are a helpful AI assistant built into the IntraOffice corporate intranet.")
        {
            var primaryModels = new[] {
                "openai/gpt-4o-mini",
                "deepseek/deepseek-chat",
                "openai/gpt-4o",
                "meta-llama/llama-3.3-70b-instruct:free",
                "google/gemini-2.0-flash-exp:free"
            };

            var allMessages = new List<object>
            {
                new { role = "system", content = systemMessage }
            };

            foreach (var msg in messages)
            {
                allMessages.Add(new { role = msg.role, content = msg.content });
            }

            foreach (var model in primaryModels)
            {
                try
                {
                    var requestBody = new
                    {
                        model = model,
                        messages = allMessages,
                        temperature = 0.3,
                        max_tokens = 4000
                    };

                    var json = JsonSerializer.Serialize(requestBody);
                    using var request = new HttpRequestMessage(HttpMethod.Post, "https://openrouter.ai/api/v1/chat/completions");
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
                    request.Headers.Add("HTTP-Referer", "https://intraoffice.local");
                    request.Headers.Add("X-Title", "IntraCom-CRM Enterprise");
                    request.Content = new StringContent(json, Encoding.UTF8, "application/json");

                    using var response = await _httpClient.SendAsync(request);
                    if (response.IsSuccessStatusCode)
                    {
                        var responseString = await response.Content.ReadAsStringAsync();
                        using var doc = JsonDocument.Parse(responseString);
                        var choices = doc.RootElement.GetProperty("choices");
                        if (choices.GetArrayLength() > 0)
                        {
                            var result = choices[0].GetProperty("message").GetProperty("content").GetString();
                            if (!string.IsNullOrWhiteSpace(result))
                            {
                                return result;
                            }
                        }
                    }
                }
                catch
                {
                    // Continue to next fallback model
                }
            }

            return string.Empty;
        }

        public async Task<string> CheckSpellingAndGrammarAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;
            string systemMsg = "You are a professional corporate copyeditor. Your job is to fix any spelling, grammar, or punctuation errors in the provided text. Return ONLY the corrected text. Do not add any conversational filler, explanations, or quotes.";
            string prompt = $"Fix this text:\n\n{text}";
            return await GenerateResponseAsync(prompt, systemMsg);
        }

        public async Task<string> SummarizeTextAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            string systemMsg = "You are an executive assistant. Your job is to provide a concise, professional summary of the provided meeting points or minute details. Return ONLY the summary, preferably as a short paragraph or a few bullet points.";
            string prompt = $"Summarize the following:\n\n{text}";
            return await GenerateResponseAsync(prompt, systemMsg);
        }

        public async Task<string> DraftContentAsync(string topic, string context)
        {
            if (string.IsNullOrWhiteSpace(topic)) return "";
            string systemMsg = "You are an expert corporate communications manager. Draft professional, clear, and engaging content based on the provided topic and context. Return ONLY the drafted content.";
            string prompt = $"Topic: {topic}\nContext: {context}\n\nPlease draft the content.";
            return await GenerateResponseAsync(prompt, systemMsg);
        }
    }
}
