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
	public class ReplyAssistResult
	{
		public bool Success { get; set; }
		public string ReplyBody { get; set; }
		public string ErrorMessage { get; set; }
	}

	public class ReplyAssist
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
			"You are an assistant that drafts email replies for Microsoft Outlook. " +
			"You will be given the original email the user received, followed by the user's instructions " +
			"for how to reply to it. Write only the body of the reply email - do not include a greeting to " +
			"the assistant, do not repeat the original email, and do not include a subject line. " +
			"Follow the user's instructions for tone and content as closely as possible. " +
			"Respond with ONLY raw JSON, no markdown fences, no commentary, in exactly this shape: " +
			"{\"body\": \"...\"}.";

		public ReplyAssist()
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

		public async Task<ReplyAssistResult> GenerateReplyAsync(string originalSubject, string originalBody, string userInstructions)
		{
			if (configError != null)
				return Fail(configError);

			if (string.IsNullOrWhiteSpace(originalBody))
				return Fail("The original email has no body text to reply to.");

			if (string.IsNullOrWhiteSpace(userInstructions))
				return Fail("Please enter instructions describing how you'd like to reply.");

			string userText =
				$"Original email subject: {originalSubject}\n\n" +
				$"Original email body:\n{originalBody}\n\n" +
				$"Instructions for the reply:\n{userInstructions}";

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

			return ParseReply(rawText);
		}

		private ReplyAssistResult ParseReply(string rawText)
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
				var payload = JsonSerializer.Deserialize<ReplyPayload>(
					cleaned, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

				if (payload != null && !string.IsNullOrWhiteSpace(payload.Body))
				{
					return new ReplyAssistResult
					{
						Success = true,
						ReplyBody = payload.Body
					};
				}
			}
			catch (JsonException)
			{
				// empty to skip
			}

			// fallback
			return new ReplyAssistResult
			{
				Success = true,
				ReplyBody = rawText.Trim()
			};
		}

		private ReplyAssistResult Fail(string message) =>
			new ReplyAssistResult { Success = false, ErrorMessage = message };


		public Outlook.MailItem ApplyToReply(Outlook.MailItem originalMailItem, ReplyAssistResult result)
		{
			if (originalMailItem == null || result == null || !result.Success)
				return null;

			Outlook.MailItem replyItem = originalMailItem.Reply() as Outlook.MailItem;
			if (replyItem == null)
				return null;

			// generated text ABOVE quoted one
			replyItem.Body = result.ReplyBody + "\r\n\r\n" + replyItem.Body;
			replyItem.Display(false);

			return replyItem;
		}

		private class ReplyPayload
		{
			public string Body { get; set; }
		}


	}

}
