using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EmailOrderPrinter.Classes;
using S22.Imap;

namespace EmailOrderPrinter
{
    public partial class frmHistories : Form
    {
        private ImapClient imapClient;
        private MailMessages mailMessages = new MailMessages();
        private Form1 form1;

        public frmHistories(ImapClient imapClient, Form1 form1)
        {
            InitializeComponent();
            this.imapClient = imapClient;
            this.form1 = form1;
            backgroundWorker1.RunWorkerAsync();
        }

        void listofMessages()
        {

            foreach (var i in imapClient.Search(
                SearchCondition.From(mailMessages.From).Or(SearchCondition.Subject(mailMessages.SearchCriterion))))
            {

                var messages = imapClient.GetMessage(i,
                    FetchOptions.Normal, true, null);
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() =>
                    {
                        dataGridView1.Rows.Add(new object[] { i, messages.Subject.ToString(), messages.Date(), "Print" });

                    }));
                }
                else
                {
                    dataGridView1.Rows.Add(new object[] { i, messages.Subject.ToString(), messages.Date(), "Print" });

                }
            }


        }
        public IEnumerable<MailMessage> MailMessages { get; set; }

        private void frmHistories_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1 && e.ColumnIndex > -1)
            {
                if (e.ColumnIndex == 3)
                {
                    var messages = imapClient.GetMessage(Convert.ToUInt32(dataGridView1[0, e.RowIndex].Value),
                           FetchOptions.Normal, true, null);
                    this.form1.notif.Visible = true;
                    messages.IsBodyHtml = true;
                    Stream contentStream = messages.AlternateViews[0].ContentStream;
                    byte[] bytes = new byte[contentStream.Length];
                    string str = Encoding.UTF8.GetString(bytes, 0, contentStream.Read(bytes, 0, bytes.Length));
                    this.form1.wb.DocumentText = str;
                    this.form1.wb.Dock = DockStyle.Fill;
                    if (this.InvokeRequired)
                    {
                        this.Invoke(new Action(() =>
                        {
                            form1.printToolStripMenuItem.Enabled = true;
                        }));
                    }
                    else
                    {
                        form1.printToolStripMenuItem.Enabled = true;
                    }

                    this.Close();
                }
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            listofMessages();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Hide();
        }
    }
}
