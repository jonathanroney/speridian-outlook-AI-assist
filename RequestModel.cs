using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Microsoft.Office.Core;

namespace OutlookAddIn1
{
    internal class Request
    {
		[JsonPropertyName("contents")]
		public List<Content> Contents { get; set; } = new List<Content>();

		[JsonPropertyName("systemInstruction")]
		public Content SystemInstruction { get; set; }
	}

	internal class Content
	{
		[JsonPropertyName("role")]
		public string Role { get; set; }

		[JsonPropertyName("parts")]
		public List<Part> Parts { get; set; } = new List<Part>();
	}

	internal class Part
	{
		[JsonPropertyName("text")]
		public string Text { get; set; }
	}
}
