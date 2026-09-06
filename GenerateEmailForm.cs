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
    public partial class GenEmailForm : Form
    {
		private readonly Outlook.MailItem mailItem;
		private int GenerateEmailMaxWords;
		private string GenerateEmailTone;

		public GenEmailForm(Outlook.MailItem mailItem)
		{
			InitializeComponent();
			this.mailItem = mailItem;
		}

		private void label2_Click(object sender, EventArgs e)
        {

        }
        private void label3_Click(object sender, EventArgs e)
        {

        }


		private void textBox1_TextChanged(object sender, EventArgs e)
		{
			if (int.TryParse(textBox1.Text, out int result))
			{
				GenerateEmailMaxWords = result;
			}
			else
			{
				MessageBox.Show("Please enter a valid number for the number of maximum words.", "Error", MessageBoxButtons.OK);
				return;
			}
		}

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
			if (ToneSelection.Text.ToLower() == "casual" || ToneSelection.Text.ToLower() == "professional")
			{
				GenerateEmailTone = ToneSelection.Text;
			}
            else
            {
                MessageBox.Show("Please select a valid tone for your email.", "Error", MessageBoxButtons.OK);
                ToneSelection.SelectedIndex = -1;
                ToneSelection.Text = string.Empty;
			}
		}

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

		private async void button1_Click(object sender, EventArgs e)
		{
			if (mailItem == null)
			{
				MessageBox.Show("Could not create a new email item.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			string instructions = richTextBox1.Text.Trim();

			if (string.IsNullOrWhiteSpace(instructions))
			{
				MessageBox.Show("Instruction box cannot be empty.", "Error", MessageBoxButtons.OK);
				return;
			}

			if (string.IsNullOrWhiteSpace(GenerateEmailTone))
			{
				MessageBox.Show("You must select a tone for your email (Professional or Casual).", "Error", MessageBoxButtons.OK);
				return;
			}

			button1.Enabled = false;
			this.Cursor = Cursors.WaitCursor;

			try
			{
				var generator = new GenerateEmail(GenerateEmailTone, GenerateEmailMaxWords);

				// This call completes on a background thread in Outlook VSTO
				var result = await generator.GenerateEmailAsync(instructions);

				// Switch back to the main UI thread for ALL window/COM operations
				SafeUIUpdate(() =>
				{
					if (!result.Success)
					{
						MessageBox.Show(result.ErrorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}

					generator.ApplyToMailItem(mailItem, result);

					this.Hide();
					mailItem.Display(false);
					this.Close();
				});
			}
			finally
			{
				SafeUIUpdate(() =>
				{
					button1.Enabled = true;
					this.Cursor = Cursors.Default;
				});
			}
		}

		private void SafeUIUpdate(Action action)
		{
			if (this.InvokeRequired)
			{
				this.Invoke(action);
			}
			else
			{
				action();
			}
		}
	}
}
