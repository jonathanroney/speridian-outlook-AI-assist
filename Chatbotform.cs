using System;
using System.Windows.Forms;

namespace OutlookAddIn1
{
	public partial class ChatbotForm : Form
	{
		private readonly ChatbotService chatbot = new ChatbotService();

		public ChatbotForm()
		{
			InitializeComponent();
		}

		private async void sendButton_Click(object sender, EventArgs e)
		{
			string userMessage = messageInputTextBox.Text.Trim();

			if (string.IsNullOrWhiteSpace(userMessage))
				return;

			AppendToHistory("You", userMessage);
			messageInputTextBox.Clear();

			SetBusyState(true);

			LoadingForm loading = new LoadingForm("Loading response...");
			loading.Show();
			loading.Refresh();

			try
			{
				var result = await chatbot.SendMessageAsync(userMessage);

				if (!result.Success)
				{
					MessageBox.Show(result.ErrorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				AppendToHistory("Assistant", result.Reply);
			}
			finally
			{
				CloseLoadingForm(loading);
				SetBusyState(false);
			}
		}

		private void clearChatButton_Click(object sender, EventArgs e)
		{
			chatbot.ClearSession();
			chatHistoryTextBox.Clear();
		}

		private void messageInputTextBox_KeyDown(object sender, KeyEventArgs e)
		{
			// enter and shift enter for send and new line
			if (e.KeyCode == Keys.Enter && !e.Shift)
			{
				e.SuppressKeyPress = true;
				sendButton_Click(sender, e);
			}
		}

		private void AppendToHistory(string speaker, string message)
		{
			if (this.InvokeRequired)
			{
				this.Invoke(new Action(() => AppendToHistory(speaker, message)));
				return;
			}

			chatHistoryTextBox.AppendText($"{speaker}: {message}{Environment.NewLine}{Environment.NewLine}");
		}

		private void SetBusyState(bool busy)
		{
			if (this.InvokeRequired)
			{
				this.Invoke(new Action(() => SetBusyState(busy)));
				return;
			}

			sendButton.Enabled = !busy;
			messageInputTextBox.Enabled = !busy;
			this.Cursor = busy ? Cursors.WaitCursor : Cursors.Default;

			if (!busy)
				messageInputTextBox.Focus();
		}

		private void CloseLoadingForm(LoadingForm loading) // for thread exception
		{
			if (loading.InvokeRequired)
			{
				loading.Invoke(new Action(() => CloseLoadingForm(loading)));
				return;
			}

			loading.Close();
		}

		private void chatHistoryTextBox_TextChanged(object sender, EventArgs e)
        {

        }
    }
}