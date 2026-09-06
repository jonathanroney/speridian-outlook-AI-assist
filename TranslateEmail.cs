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
	public class TranslateEmailResult
	{
		public bool Success { get; set; }
		public string Subject { get; set; }
		public string Body { get; set; }
		public string ErrorMessage { get; set; }
	}

	public class TranslateEmail
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
		private string TargetLanguage;
		private string SystemPrompt;

		public TranslateEmail(string SelectedTargetLanguage)
		{
			this.apiKey = ConfigurationManager.AppSettings["GEMINI_API_KEY"];
			this.apiEndpoint = ConfigurationManager.AppSettings["GEMINI_API_ENDPOINT"];
			this.modelName = ConfigurationManager.AppSettings["GEMINI_MODEL"];
			if (string.IsNullOrWhiteSpace(SelectedTargetLanguage))
			{
				configError = "Missing target language";
				return;
			}

			TargetLanguage = SelectedTargetLanguage;

			SystemPrompt =
				$"You are a translation assistant for Microsoft Outlook emails. " +
				$"Translate the given email subject and body into {TargetLanguage}. " +
				"Preserve the original meaning and tone as closely as possible. " +
				"Do not add any new information, commentary, greetings, or signatures that were not already present. " +
				"Do not omit any part of the original content. " +
				"Respond with ONLY raw JSON, no markdown fences, no commentary, in exactly this shape: " +
				"{\"subject\": \"...\", \"body\": \"...\"}.";


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

		public async Task<TranslateEmailResult> TranslateEmailAsync(string originalSubject, string originalBody)
		{
			if (configError != null)
				return Fail(configError);

			if (string.IsNullOrWhiteSpace(originalBody))
				return Fail("Your email body is empty");

			string userText =
				$"Subject: {originalSubject}\n\n" +
				$"Body:\n{originalBody}";

			var request = new Request
			{
				SystemInstruction = new Content { Parts = { new Part { Text = SystemPrompt } } },
				Contents = { new Content { Role = "user", Parts = { new Part { Text = userText } } } }
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

			return ParseTranslatedEmail(rawText, originalSubject);
		}

		private TranslateEmailResult ParseTranslatedEmail(string rawText, string originalSubject)
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
				var payload = JsonSerializer.Deserialize<TranslatedEmailPayload>(
					cleaned, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

				if (payload != null && !string.IsNullOrWhiteSpace(payload.Body))
				{
					return new TranslateEmailResult
					{
						Success = true,
						Subject = string.IsNullOrWhiteSpace(payload.Subject) ? originalSubject : payload.Subject,
						Body = payload.Body
					};
				}
			}
			catch (JsonException)
			{
				// empty to skip
			}

			return Fail("The AI API request failed: could not parse translation.");
		}

		private TranslateEmailResult Fail(string message) =>
			new TranslateEmailResult { Success = false, ErrorMessage = message };

		public void ApplyToMailItem(Outlook.MailItem mailItem, TranslateEmailResult result)
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
					mailItem.Save();
				}));
			}
			else
			{
				mailItem.Subject = result.Subject;
				mailItem.Body = result.Body;
				mailItem.Save();
			}
		}

		private class TranslatedEmailPayload
		{
			public string Subject { get; set; }
			public string Body { get; set; }
		}
	}
}

