using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json.Serialization;
using System.Web;

namespace OutlookAddIn1
{
    internal class Response
    {
        [JsonPropertyName("candidates")] // alternates included
        public List<Candidate> Candidates { get; set; }

        [JsonPropertyName("promptFeedback")]
        public PromptFeedback PromptFeedback { get; set; }
    }

    internal class Candidate
    {
        [JsonPropertyName("content")]
        public Content Content { get; set; }

        [JsonPropertyName("finishReason")]
        public string FinishReason { get; set; }
    }

    internal class PromptFeedback
    {
        [JsonPropertyName("blockReason")]
        public string BlockReason { get; set; }
    }
}
