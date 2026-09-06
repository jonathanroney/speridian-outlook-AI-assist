using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;



namespace OutlookAddIn1
{
	public class GeneratedEmailResult
	{
		public bool Success { get; set; }
		public string Subject { get; set; }
		public string Body { get; set; }
		public string ErrorMessage { get; set; }
	}

	public class GenerateEmail
    {
        private static readonly HttpClient httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(30)
        };

		// variables
        private readonly string apiKey;
        private readonly string apiEndpoint;
        private readonly string modelName;
        private string configError;
		private string Tone;
		private int MaxWords;
		private string SystemPrompt;

		public GenerateEmail(string SelectedTone, int SelectedMaxWords)
        {
            this.apiKey = ConfigurationManager.AppSettings["GEMINI_API_KEY"];
            this.apiEndpoint = ConfigurationManager.AppSettings["GEMINI_API_ENDPOINT"];
            this.modelName = ConfigurationManager.AppSettings["GEMINI_MODEL"];
			if (SelectedTone.ToLower() == "professional")
			{
				Tone = "professional";
			}
			else if (SelectedTone.ToLower() == "casual")
			{
				Tone = "casual";
			}
			else
			{
				configError = "Invalid tone selected"; // technically not a config error? good for reusability though
				return;
			}

			if (SelectedMaxWords == 0)
			{
				SystemPrompt = $"You are an assistant that drafts {Tone} emails for Microsoft Outlook. " +
			"Given the user's instructions, write an appropriate email subject and body. " +
			"Respond with ONLY raw JSON, no markdown fences, no commentary, in exactly this shape: " +
			"{\"subject\": \"...\", \"body\": \"...\"}.";
			}
			else
			{
				MaxWords = SelectedMaxWords;
				SystemPrompt = $"You are an assistant that drafts {Tone} emails for Microsoft Outlook. " +
			$"Given the user's instructions, write an appropriate email subject and body with a maximum word count of {MaxWords}. " +
			"Respond with ONLY raw JSON, no markdown fences, no commentary, in exactly this shape: " +
			"{\"subject\": \"...\", \"body\": \"...\"}.";
			}

			ErrorHandling();
        }

        private void ErrorHandling()
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                configError = "Missing API key";

            if (string.IsNullOrWhiteSpace(apiEndpoint))
                configError = "Missing API endpoint";

            if (string.IsNullOrWhiteSpace(modelName))
                configError = "Missing model name";
		}

		public async Task<GeneratedEmailResult> GenerateEmailAsync(string userInstruction)
		{
			if (configError != null)
				return Fail(configError);

			if (string.IsNullOrWhiteSpace(userInstruction))
				return Fail("Please enter instructions describing the email you want to generate.");

			var request = new Request
			{
				SystemInstruction = new Content { Parts = { new Part { Text = SystemPrompt } } },
				Contents = { new Content { Role = "user", Parts = { new Part { Text = userInstruction } } } }
			};

			string url = $"{apiEndpoint.TrimEnd('/')}/{modelName}:generateContent?key={apiKey}";
			string responseBody;
			HttpResponseMessage httpResponse;

			try
			{
				using (var content = new StringContent(JsonSerializer.Serialize(request), Encoding.UTF8, "application/json"))
				{
					httpResponse = await httpClient.PostAsync(url, content);
				}
				responseBody = await httpResponse.Content.ReadAsStringAsync();
			}
			catch (TaskCanceledException)
			{
				return Fail("The request to the AI service timed out.");
			}
			catch (HttpRequestException)
			{
				return Fail("Could not reach the AI service.");
			}

			if (!httpResponse.IsSuccessStatusCode)
			{
				switch ((int)httpResponse.StatusCode)
				{
					case 401:
					case 403:
						return Fail("Invalid API key");
					case 429:
						return Fail("API quota has been exceeded");
					default:
						return Fail($"The AI API request failed ({(int)httpResponse.StatusCode}).");
				}
			}

			Response geminiResponse;
			try
			{
				geminiResponse = JsonSerializer.Deserialize<Response>(responseBody);
			}
			catch (JsonException)
			{
				return Fail("The AI API request failed: unreadable response.");
			}

			string rawText = geminiResponse?.Candidates?
				.FirstOrDefault()?.Content?.Parts?
				.FirstOrDefault(p => !string.IsNullOrEmpty(p.Text))?.Text;

			if (string.IsNullOrWhiteSpace(rawText))
				return Fail("The AI API request failed: no content returned.");

			return ParseGeneratedEmail(rawText);
		}

		private GeneratedEmailResult ParseGeneratedEmail(string rawText)
		{
			string cleaned = rawText.Trim();

			if (cleaned.StartsWith("```"))
			{
				int firstNewline = cleaned.IndexOf('\n');
				int lastFence = cleaned.LastIndexOf("```");
				if (firstNewline >= 0 && lastFence > firstNewline)
					cleaned = cleaned.Substring(firstNewline + 1, lastFence - firstNewline - 1).Trim();
			}

			try
			{
				var payload = JsonSerializer.Deserialize<GeneratedEmailPayload>(
					cleaned, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

				if (payload != null && !string.IsNullOrWhiteSpace(payload.Body))
				{
					return new GeneratedEmailResult
					{
						Success = true,
						Subject = string.IsNullOrWhiteSpace(payload.Subject) ? "Generated Email" : payload.Subject,
						Body = payload.Body
					};
				}
			}
			catch (JsonException)
			{
				// empty to skip
			}

			return new GeneratedEmailResult
			{
				Success = true,
				Subject = "Generated Email",
				Body = rawText.Trim()
			};
		}

		private GeneratedEmailResult Fail(string message) =>
			new GeneratedEmailResult { Success = false, ErrorMessage = message };

		public void ApplyToMailItem(Outlook.MailItem mailItem, GeneratedEmailResult result)
		{
			if (mailItem == null || result == null || !result.Success)
				return;

			// Ensure Outlook COM properties are updated on the main STA thread
			if (Application.OpenForms.Count > 0 && Application.OpenForms[0].InvokeRequired)
			{
				Application.OpenForms[0].Invoke(new Action(() =>
				{
					mailItem.Subject = result.Subject;
					mailItem.Body = result.Body;
				}));
			}
			else
			{
				mailItem.Subject = result.Subject;
				mailItem.Body = result.Body;
			}
		}

		private class GeneratedEmailPayload
		{
			public string Subject { get; set; }
			public string Body { get; set; }
		}
	}
}
