namespace OutlookAddIn1
{
	partial class ChatbotForm
	{
		private System.ComponentModel.IContainer components = null;
		private System.Windows.Forms.TextBox chatHistoryTextBox;
		private System.Windows.Forms.TextBox messageInputTextBox;
		private System.Windows.Forms.Button sendButton;
		private System.Windows.Forms.Button clearChatButton;

		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		private void InitializeComponent()
		{
			this.chatHistoryTextBox = new System.Windows.Forms.TextBox();
			this.messageInputTextBox = new System.Windows.Forms.TextBox();
			this.sendButton = new System.Windows.Forms.Button();
			this.clearChatButton = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// chatHistoryTextBox
			// 
			this.chatHistoryTextBox.Location = new System.Drawing.Point(19, 19);
			this.chatHistoryTextBox.Margin = new System.Windows.Forms.Padding(4);
			this.chatHistoryTextBox.Multiline = true;
			this.chatHistoryTextBox.Name = "chatHistoryTextBox";
			this.chatHistoryTextBox.ReadOnly = true;
			this.chatHistoryTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.chatHistoryTextBox.Size = new System.Drawing.Size(524, 374);
			this.chatHistoryTextBox.TabIndex = 0;
			this.chatHistoryTextBox.TabStop = false;
			this.chatHistoryTextBox.TextChanged += new System.EventHandler(this.chatHistoryTextBox_TextChanged);
			// 
			// messageInputTextBox
			// 
			this.messageInputTextBox.Location = new System.Drawing.Point(19, 406);
			this.messageInputTextBox.Margin = new System.Windows.Forms.Padding(4);
			this.messageInputTextBox.Multiline = true;
			this.messageInputTextBox.Name = "messageInputTextBox";
			this.messageInputTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.messageInputTextBox.Size = new System.Drawing.Size(430, 62);
			this.messageInputTextBox.TabIndex = 1;
			this.messageInputTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.messageInputTextBox_KeyDown);
			// 
			// sendButton
			// 
			this.sendButton.Location = new System.Drawing.Point(457, 406);
			this.sendButton.Margin = new System.Windows.Forms.Padding(4);
			this.sendButton.Name = "sendButton";
			this.sendButton.Size = new System.Drawing.Size(86, 38);
			this.sendButton.TabIndex = 2;
			this.sendButton.Text = "Send";
			this.sendButton.UseVisualStyleBackColor = true;
			this.sendButton.Click += new System.EventHandler(this.sendButton_Click);
			// 
			// clearChatButton
			// 
			this.clearChatButton.Location = new System.Drawing.Point(160, 487);
			this.clearChatButton.Margin = new System.Windows.Forms.Padding(4);
			this.clearChatButton.Name = "clearChatButton";
			this.clearChatButton.Size = new System.Drawing.Size(250, 38);
			this.clearChatButton.TabIndex = 3;
			this.clearChatButton.Text = "Clear Chat / New Session";
			this.clearChatButton.UseVisualStyleBackColor = true;
			this.clearChatButton.Click += new System.EventHandler(this.clearChatButton_Click);
			// 
			// ChatbotForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(120F, 120F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
			this.ClientSize = new System.Drawing.Size(562, 536);
			this.Controls.Add(this.chatHistoryTextBox);
			this.Controls.Add(this.messageInputTextBox);
			this.Controls.Add(this.sendButton);
			this.Controls.Add(this.clearChatButton);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Margin = new System.Windows.Forms.Padding(4);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ChatbotForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "AI Chatbot";
			this.ResumeLayout(false);
			this.PerformLayout();

		}
	}
}