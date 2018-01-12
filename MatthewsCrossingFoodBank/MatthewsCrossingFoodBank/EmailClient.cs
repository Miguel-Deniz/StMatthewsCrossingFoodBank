using System.Net.Mail;
using System.Windows.Forms;

namespace MatthewsCrossingFoodBank
{
    class EmailClient
    {
        private string _emailAddress;
        private string _password;
        private string _stmpServer;
        private int _port;
        private bool _enableSSl;

        public EmailClient(string emailAddress, string password, string stmpServer, int port, bool enableSSL)
        {
            _emailAddress = emailAddress;
            _password = password;
            _stmpServer = stmpServer;
            _port = port;
            _enableSSl = enableSSL;
        }

        public void sendEmail(string to, string subject, string body)
        {
            MailMessage mail = new MailMessageBuilder().from(_emailAddress).to(to).subject(subject).body(body).HTMLBody(true).build();

            if (mail != null)
            {
                SmtpClient SmtpServer = new SmtpClient(_stmpServer);

                SmtpServer.Port = _port;
                SmtpServer.Credentials = new System.Net.NetworkCredential(_emailAddress, _password);
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);
                MessageBox.Show("Mail Sent!");
            } else
            {
                MessageBox.Show("Unable to send message");
            }
        }
    }
}
