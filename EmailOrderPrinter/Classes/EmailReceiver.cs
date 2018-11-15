namespace EmailOrderPrinter.Classes
{
    using S22.Imap;
    using System;
    using System.Runtime.CompilerServices;
    using System.Threading;

    internal class EmailReceiver
    {
        public event EventHandler<IdleMessageEventArgs> NewMessage;

        public EmailReceiver(string Server, int Port, string Username, string Password, string SearchCriterion)
        {
            this.Server = Server;
            this.Port = Port;
            this.Username = Username;
            this.Password = Password;
            this.SearchCriterion = SearchCriterion;
        }

        private void client_NewMessage(object sender, IdleMessageEventArgs e)
        {
        }

        public ImapClient Receive()
        {
            return new ImapClient(this.Server, this.Port, this.Username, this.Password, AuthMethod.Auto, true, null);
        }

        public ImapClient Client { get; set; }

        public string Password { get; set; }

        public int Port { get; set; }

        public string SearchCriterion { get; set; }

        public string Server { get; set; }

        public string Username { get; set; }
    }

    public class MailMessages
    {
        public uint UID { get; set; }
        public string Subject { get; set; }
        public string SearchCriterion => "[Sylvia's Flowers]";
        public string From => "no-reply@shopify.com";
    }
}

