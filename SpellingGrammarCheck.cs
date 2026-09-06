using System;
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
    public class SpellingGrammarCheckResult
    {
        public bool Success { get; set; }
        public string CorrectedBody { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class SpellingGrammarCheck
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
		private readonly string SystemPrompt =
			"You are a spelling and grammar correction assistant for Microsoft Outlook emails. " +
			"Correct only spelling and grammar mistakes in the given email body. " +
			"Do not change the meaning, tone, wording style, or formatting beyond what is needed to fix errors. " +
			"Do not add greetings, signatures, or commentary that were not already present. " +
			"Respond with ONLY the corrected email body as plain text, no markdown fences, no commentary, no JSON.";

		public SpellingGrammarCheck()
		{
			this.apiKey = ConfigurationManager.AppSettings["GEMINI_API_KEY"];
			this.apiEndpoint = ConfigurationManager.AppSettings["GEMINI_API_ENDPOINT"];
			this.modelName = ConfigurationManager.AppSettings["GEMINI_MODEL"];

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

		public async Task<SpellingGrammarCheckResult> SpellCheckAsync(string originalBody)
		{
			if (configError != null)
				return Fail(configError);

			if (string.IsNullOrWhiteSpace(originalBody))
				return Fail("The email body is empty, nothing to check.");

			var request = new Request
			{
				SystemInstruction = new Content { Parts = { new Part { Text = SystemPrompt } } },
				Contents = { new Content { Role = "user", Parts = { new Part { Text = originalBody } } } }
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

			string cleaned = CleanFences(rawText);

			return new SpellingGrammarCheckResult
			{
				Success = true,
				CorrectedBody = cleaned
			};
		}

		private string CleanFences(string rawText) // recommended method - dive deeper later
		{
			string cleaned = rawText.Trim();

			if (cleaned.StartsWith("```"))
			{
				int firstNewline = cleaned.IndexOf('\n');
				int lastFence = cleaned.LastIndexOf("```");
				if (firstNewline >= 0 && lastFence > firstNewline)
					cleaned = cleaned.Substring(firstNewline + 1, lastFence - firstNewline - 1).Trim();
			}

			return cleaned;
		}

		private SpellingGrammarCheckResult Fail(string message) =>
			new SpellingGrammarCheckResult { Success = false, ErrorMessage = message };

		public void ApplyToMailItem(Outlook.MailItem mailItem, SpellingGrammarCheckResult result)
		{
			if (mailItem == null || result == null || !result.Success)
				return;

			// Ensure Outlook COM properties are updated on the main STA thread
			if (Application.OpenForms.Count > 0 && Application.OpenForms[0].InvokeRequired)
			{
				Application.OpenForms[0].Invoke(new Action(() =>
				{
					mailItem.Body = result.CorrectedBody;
				}));
			}
			else
			{
				mailItem.Body = result.CorrectedBody;
			}
		}
	}
}