using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Impulse.Services.IntraOffice
{
    public interface IAiService
    {
        Task<string> GenerateResponseAsync(string prompt, string systemMessage = "You are a helpful AI assistant built into the IntraOffice corporate intranet.");
        Task<string> GenerateConversationResponseAsync(List<(string role, string content)> messages, string systemMessage = "You are a helpful AI assistant built into the IntraOffice corporate intranet.");
        Task<string> CheckSpellingAndGrammarAsync(string text);
        Task<string> SummarizeTextAsync(string text);
        Task<string> DraftContentAsync(string topic, string context);
    }
}
