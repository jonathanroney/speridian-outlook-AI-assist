namespace OutlookAddIn1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
			this.AIAssist = this.Factory.CreateRibbonTab();
			this.GenerateEmail = this.Factory.CreateRibbonGroup();
			this.GenerateEmailButton = this.Factory.CreateRibbonButton();
			this.SpellCheck = this.Factory.CreateRibbonButton();
			this.ReplyAssist = this.Factory.CreateRibbonButton();
			this.LangConversion = this.Factory.CreateRibbonButton();
			this.ChatBotGroup = this.Factory.CreateRibbonGroup();
			this.Chatbot = this.Factory.CreateRibbonButton();
			this.AIAssist.SuspendLayout();
			this.GenerateEmail.SuspendLayout();
			this.ChatBotGroup.SuspendLayout();
			this.SuspendLayout();
			// 
			// AIAssist
			// 
			this.AIAssist.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.AIAssist.Groups.Add(this.GenerateEmail);
			this.AIAssist.Groups.Add(this.ChatBotGroup);
			this.AIAssist.Label = "AI Assist";
			this.AIAssist.Name = "AIAssist";
			// 
			// GenerateEmail
			// 
			this.GenerateEmail.Items.Add(this.GenerateEmailButton);
			this.GenerateEmail.Items.Add(this.SpellCheck);
			this.GenerateEmail.Items.Add(this.ReplyAssist);
			this.GenerateEmail.Items.Add(this.LangConversion);
			this.GenerateEmail.Name = "GenerateEmail";
			// 
			// GenerateEmailButton
			// 
			this.GenerateEmailButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.GenerateEmailButton.Image = global::OutlookAddIn1.Properties.Resources.email;
			this.GenerateEmailButton.Label = "Generate Email";
			this.GenerateEmailButton.Name = "GenerateEmailButton";
			this.GenerateEmailButton.ShowImage = true;
			this.GenerateEmailButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GenerateEmailButton_Click);
			// 
			// SpellCheck
			// 
			this.SpellCheck.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.SpellCheck.Description = "Check your email\'s grammar and spelling.";
			this.SpellCheck.Image = global::OutlookAddIn1.Properties.Resources.spell_check;
			this.SpellCheck.Label = "Spell Check";
			this.SpellCheck.Name = "SpellCheck";
			this.SpellCheck.ShowImage = true;
			this.SpellCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SpellCheck_Click);
			// 
			// ReplyAssist
			// 
			this.ReplyAssist.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.ReplyAssist.Description = "Generate possible replies.";
			this.ReplyAssist.Image = global::OutlookAddIn1.Properties.Resources.reply_1_;
			this.ReplyAssist.Label = "Reply Assist";
			this.ReplyAssist.Name = "ReplyAssist";
			this.ReplyAssist.ShowImage = true;
			this.ReplyAssist.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReplyAssist_Click);
			// 
			// LangConversion
			// 
			this.LangConversion.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.LangConversion.Description = "Convert your email to another language.";
			this.LangConversion.Image = global::OutlookAddIn1.Properties.Resources.translate;
			this.LangConversion.Label = "Language Conversion";
			this.LangConversion.Name = "LangConversion";
			this.LangConversion.ShowImage = true;
			this.LangConversion.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LanguageConversion_Click);
			// 
			// ChatBotGroup
			// 
			this.ChatBotGroup.Items.Add(this.Chatbot);
			this.ChatBotGroup.Name = "ChatBotGroup";
			// 
			// Chatbot
			// 
			this.Chatbot.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.Chatbot.Description = "Talk to a chatbot for e-mail assistance.";
			this.Chatbot.Image = global::OutlookAddIn1.Properties.Resources.chatbot;
			this.Chatbot.Label = "Chatbot";
			this.Chatbot.Name = "Chatbot";
			this.Chatbot.ShowImage = true;
			this.Chatbot.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Chatbot_Click);
			// 
			// Ribbon1
			// 
			this.Name = "Ribbon1";
			this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Compose";
			this.Tabs.Add(this.AIAssist);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
			this.AIAssist.ResumeLayout(false);
			this.AIAssist.PerformLayout();
			this.GenerateEmail.ResumeLayout(false);
			this.GenerateEmail.PerformLayout();
			this.ChatBotGroup.ResumeLayout(false);
			this.ChatBotGroup.PerformLayout();
			this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab AIAssist;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GenerateEmail;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ChatBotGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SpellCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ReplyAssist;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LangConversion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Chatbot;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GenerateEmailButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
