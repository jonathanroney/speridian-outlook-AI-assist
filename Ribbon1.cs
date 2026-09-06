using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1
{
    public partial class Ribbon1
    {
        private ReplyAssistForm replyAssistForm;
		private ChatbotForm chatbotForm;

		private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

		private void ReplyAssist_Click(object sender, RibbonControlEventArgs e)
		{
			Outlook.MailItem mailItem = null;

			Outlook.Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();
			if (inspector?.CurrentItem is Outlook.MailItem inspectorItem)
			{
				mailItem = inspectorItem;
			}
			else
			{
				Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
				if (explorer?.Selection != null && explorer.Selection.Count > 0
					&& explorer.Selection[1] is Outlook.MailItem selectedItem)
				{
					mailItem = selectedItem;
				}
			}

			if (mailItem == null)
			{
				MessageBox.Show("Please open or select an email to reply to.", "Error", MessageBoxButtons.OK);
				return;
			}

			if (replyAssistForm != null && !replyAssistForm.IsDisposed)
			{
				replyAssistForm.Activate();
				return;
			}

			replyAssistForm = new ReplyAssistForm(mailItem);
			replyAssistForm.FormClosed += (s, args) => replyAssistForm = null;
			replyAssistForm.Show();
		}

		private void LanguageConversion_Click(object sender, RibbonControlEventArgs e)
        {
			Outlook.Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();
			Outlook.MailItem mailItem = inspector?.CurrentItem as Outlook.MailItem;

			if (mailItem == null)
			{
				MessageBox.Show("No MailItem open - open a mail to translate!", "Error", MessageBoxButtons.OK);
				return;
			}

			TranslateEmailForm form = new TranslateEmailForm(mailItem);
			form.ShowDialog();

		}

		private async void SpellCheck_Click(object sender, RibbonControlEventArgs e) // done
		{
			Outlook.Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();
			Outlook.MailItem mailItem = inspector?.CurrentItem as Outlook.MailItem;

			if (mailItem == null)
			{
				MessageBox.Show("Please open an email to spell check.", "Error", MessageBoxButtons.OK);
				return;
			}

			string originalBody = mailItem.Body;

			if (string.IsNullOrWhiteSpace(originalBody))
			{
				MessageBox.Show("This email has no body text to check.", "Error", MessageBoxButtons.OK);
				return;
			}

			LoadingForm loading = new LoadingForm("Checking spelling and grammar...");
			loading.Show();
			loading.Refresh(); // force it to paint immediately before the await

			try
			{
				var checker = new SpellingGrammarCheck();
				var result = await checker.SpellCheckAsync(originalBody);

				if (!result.Success)
				{
					MessageBox.Show(result.ErrorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				checker.ApplyToMailItem(mailItem, result);
			}
			finally
			{
				loading.Close();
			}
		}

		private void Chatbot_Click(object sender, RibbonControlEventArgs e)
		{
			if (chatbotForm != null && !chatbotForm.IsDisposed)
			{
				chatbotForm.Activate();
				return;
			}

			chatbotForm = new ChatbotForm();
			chatbotForm.FormClosed += (s, args) => chatbotForm = null;
			chatbotForm.Show();
		}


		private void GenerateEmailButton_Click(object sender, RibbonControlEventArgs e)
        {
			Outlook.MailItem mailItem = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

			GenEmailForm form = new GenEmailForm(mailItem);
			form.Show();
		}
    }
}
