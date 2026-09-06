using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1
{
    public partial class ReplyAssistForm : Form
    {
        private readonly Outlook.MailItem originalMailItem;
        public ReplyAssistForm(Outlook.MailItem originalMailItem)
        {
			InitializeComponent();
			this.originalMailItem = originalMailItem;
		}

		private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private async void button1_Click(object sender, EventArgs e)
        {
            if (originalMailItem == null)
            {
                MessageBox.Show("No email found to reply to - please open an email", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

			string originalMailContent = richTextBox2.Text.Trim();

			if (string.IsNullOrWhiteSpace(originalMailContent))
			{
				MessageBox.Show("Please enter the original email content that you wish to reply to.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			string instructions = richTextBox1.Text.Trim();

            if (string.IsNullOrWhiteSpace(instructions))
            {
                MessageBox.Show("Please enter the instructions needed to generate the reply to the email.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

			button1.Enabled = false;
			this.Cursor = Cursors.WaitCursor;

			LoadingForm loading = new LoadingForm("Drafting your reply...");
			loading.Show();
			loading.Refresh();

			try
			{
				string originalSubject = originalMailItem.Subject;

				var replyassist = new ReplyAssist();
				var result = await replyassist.GenerateReplyAsync(originalSubject, originalMailContent, instructions);

				if (!result.Success)
				{
					MessageBox.Show(result.ErrorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				replyassist.ApplyToReply(originalMailItem, result);
				this.Close();
			}
			finally
			{
				loading.Close();
				button1.Enabled = true;
				this.Cursor = Cursors.Default;
			}

		}

		private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
