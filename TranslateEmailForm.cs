using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1
{
    public partial class TranslateEmailForm : Form
    {
		private readonly Outlook.MailItem mailItem;
		private string SelectedLanguage;

        public TranslateEmailForm(Outlook.MailItem mailItem)
        {
            InitializeComponent();
			this.mailItem = mailItem;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private async void button1_Click(object sender, EventArgs e)
        {
			if (mailItem == null)
			{
				MessageBox.Show("No active email found to translate.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			string originalSubject = mailItem.Subject;
			string originalBody = mailItem.Body;

			if (string.IsNullOrWhiteSpace(originalBody))
			{
				MessageBox.Show("This email has no body text to translate.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			string typedLanguage = comboBox1.Text.Trim();

			if (string.IsNullOrWhiteSpace(typedLanguage))
			{
				MessageBox.Show("Please select a target language.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			bool isKnown = comboBox1.Items.Cast<string>()
				.Any(l => l.Equals(typedLanguage, StringComparison.OrdinalIgnoreCase));

			if (!isKnown)
			{
				MessageBox.Show($"\"{typedLanguage}\" isn't in the supported languages list.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			SelectedLanguage = typedLanguage;

			button1.Enabled = false;
			this.Cursor = Cursors.WaitCursor;

			LoadingForm loading = new LoadingForm($"Translating to {SelectedLanguage}...");
			loading.Show();
			loading.Refresh();

			try
			{
				var translator = new TranslateEmail(SelectedLanguage);
				var result = await translator.TranslateEmailAsync(originalSubject, originalBody);

				if (!result.Success)
				{
					MessageBox.Show(result.ErrorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				translator.ApplyToMailItem(mailItem, result);
				this.Close();
			}
			finally
			{
				loading.Close();
				button1.Enabled = true;
				this.Cursor = Cursors.Default;
			}

		}
	}
}
