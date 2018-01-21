using System;
using System.Net.Mail;

namespace MatthewsCrossingFoodBank
{
    public class MailMessageBuilder
    {
        private string _from = null;
        private string _to = null;
        private string _subject = null;
        private string _body = null;
        private bool _isBodyHTML = true;

        public MailMessageBuilder()
        {
            // Default
        }

        public MailMessageBuilder from(string email)
        {
            _from = email;
            return this;
        }

        public MailMessageBuilder to(string email)
        {
            _to = email;
            return this;
        }

        public MailMessageBuilder subject(string sub)
        {
            _subject = sub;
            return this;
        }
        
        public MailMessageBuilder body(string body)
        {
            _body = body;
            return this;
        }

        public MailMessageBuilder HTMLBody(bool b)
        {
            _isBodyHTML = b;
            return this;
        }

        public MailMessage build()
        {
            if (_from == null || _to == null || _subject == null || _body == null)
                return null;

            try
            {
                MailMessage mail = new MailMessage();

                mail.From = new MailAddress(_from);
                mail.To.Add(_to);
                mail.Subject = _subject;
                mail.Body = _body;
                mail.IsBodyHtml = _isBodyHTML;

                return mail;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                return null;
            }
        }
    }
}
