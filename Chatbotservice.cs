using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace OutlookAddIn1
{
	public class ChatbotResult
	{
		public bool Success { get; set; }
		public string Reply { get; set; }
		public string ErrorMessage { get; set; }
	}

	public class ChatbotService
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
			"You are a helpful assistant embedded in a Microsoft Outlook add-in. " +
			"Answer the user's questions clearly and concisely. " +
			"Respond with ONLY raw JSON, no markdown fences, no commentary, in exactly this shape: " +
			"{\"reply\": \"...\"}.";


		private class ChatMessage
		{
			public string Role;
			public string Content;
			public ChatMessage(string role, string content)
			{
				Role = role; Content = content;
			}
		}

		private readonly List<ChatMessage> history = new List<ChatMessage>();

		public ChatbotService()
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

		public void ClearSession()
		{
			history.Clear();
		}

		public async Task<ChatbotResult> SendMessageAsync(string userMessage)
		{
			if (configError != null)
				return Fail(configError);

			if (string.IsNullOrWhiteSpace(userMessage))
				return Fail("There's no message to send.");

			history.Add(new ChatMessage("user", userMessage));

			var request = new Request
			{
				SystemInstruction = new Content { Parts = { new Part { Text = SystemPrompt } } }
			};

			foreach (var message in history)
			{
				request.Contents.Add(new Content
				{
					Role = message.Role,
					Parts = { new Part { Text = message.Content } }
				});
			}

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

			var result = ParseReply(rawText);

			if (result.Success)
				history.Add(new ChatMessage("model", result.Reply));

			return result;
		}

		private ChatbotResult ParseReply(string rawText)
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
				var payload = JsonSerializer.Deserialize<ChatbotPayload>(
					cleaned, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

				if (payload != null && !string.IsNullOrWhiteSpace(payload.Reply))
				{
					return new ChatbotResult
					{
						Success = true,
						Reply = payload.Reply
					};
				}
			}
			catch (JsonException)
			{
				// empty to skip
			}

			return new ChatbotResult
			{
				Success = true,
				Reply = rawText.Trim()
			};
		}

		private ChatbotResult Fail(string message) =>
			new ChatbotResult { Success = false, ErrorMessage = message };

		private class ChatbotPayload
		{
			public string Reply { get; set; }
		}
	}
}
