using Properties;

namespace EmailOrderPrinter
{
    using EmailOrderPrinter.Classes;
    using S22.Imap;
    using System;
    using System.ComponentModel;
    using System.Drawing;
    using System.IO;
    using System.Net.Mail;
    using System.Text;
    using System.Threading;
    using System.Windows.Forms;

    public class Form1 : Form
    {
        private IContainer components = null;
        public NotifyIcon notif;
        public MenuStrip menuStrip1;
        public ToolStripMenuItem printToolStripMenuItem;
        public ToolStripMenuItem historiesToolStripMenuItem;
        public WebBrowser wb;

        public Form1()
        {
            this.InitializeComponent();
            base.Icon = Icon.FromHandle(Properties.Resources.printer1.GetHicon());
        }

        private MailMessages mailMessages = new MailMessages();
        public MailMessage Messages { get; set; }
        private void client_NewMessage(object sender, IdleMessageEventArgs e)
        {

            MailMessage message = e.Client.GetMessage(e.MessageUID, FetchOptions.Normal, true, null);

            if (message.Subject.ToLower().Contains(mailMessages.SearchCriterion.ToLower()) || message.From.Address.ToLower().Contains(mailMessages.From.ToLower()))
            {
                this.Messages = message;
                this.notif.ShowBalloonTip(100, "New Email Receive", "Printing...", ToolTipIcon.Info);
                this.notif.Visible = true;
                message.IsBodyHtml = true;
                Stream contentStream = message.AlternateViews[0].ContentStream;
                byte[] bytes = new byte[contentStream.Length];
                string str = Encoding.UTF8.GetString(bytes, 0, contentStream.Read(bytes, 0, bytes.Length));
                this.wb.DocumentText = str;
                this.wb.Dock = DockStyle.Fill;
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() =>
                    {
                        printToolStripMenuItem.Enabled = true;
                    }));
                }
                else
                {
                    printToolStripMenuItem.Enabled = true;
                }
                print();

            }


        }

        public void print()
        {
            this.wb.Print();
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }
        //sylv1@fl0w3rsh0p sylvia.flowershop
        public ImapClient imapClient;
        private void Form1_Load(object sender, EventArgs e)
        {
            var receiver = new EmailReceiver("imap.gmail.com", 993, "john.go1172015@gmail.com", "b3foreandafter", "");
            imapClient = receiver.Receive();
            imapClient.NewMessage += new EventHandler<IdleMessageEventArgs>(this.client_NewMessage);
            this.notif.Icon = Icon.FromHandle(Resources.printer1.GetHicon());
            Thread.Sleep(0x3e8);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.notif = new System.Windows.Forms.NotifyIcon(this.components);
            this.wb = new System.Windows.Forms.WebBrowser();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.printToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.historiesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // wb
            // 
            this.wb.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wb.Location = new System.Drawing.Point(0, 24);
            this.wb.Margin = new System.Windows.Forms.Padding(2);
            this.wb.MinimumSize = new System.Drawing.Size(15, 16);
            this.wb.Name = "wb";
            this.wb.Size = new System.Drawing.Size(413, 252);
            this.wb.TabIndex = 0;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.printToolStripMenuItem,
            this.historiesToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(413, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // printToolStripMenuItem
            // 
            this.printToolStripMenuItem.Enabled = false;
            this.printToolStripMenuItem.Name = "printToolStripMenuItem";
            this.printToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.printToolStripMenuItem.Text = "Print";
            this.printToolStripMenuItem.Click += new System.EventHandler(this.printToolStripMenuItem_Click);
            // 
            // historiesToolStripMenuItem
            // 
            this.historiesToolStripMenuItem.Name = "historiesToolStripMenuItem";
            this.historiesToolStripMenuItem.Size = new System.Drawing.Size(65, 20);
            this.historiesToolStripMenuItem.Text = "Histories";
            this.historiesToolStripMenuItem.Click += new System.EventHandler(this.historiesToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(413, 276);
            this.Controls.Add(this.wb);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Email Order Printer";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            print();
        }

        private void historiesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //var messages =
            //    imapClient.GetMessages(
            //        imapClient.Search(SearchCondition.From(From).Or(SearchCondition.Body(SearchCriterion))),
            //        FetchOptions.Normal, true, null);



            frmHistories frmHistories = new frmHistories(imapClient,this);
            frmHistories.ShowDialog();
        }
    }
}

