using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Net.Mail;
using MatthewsCrossingFoodBank;

namespace MatthewsCrossingFoodBankTesting
{
    [TestClass]
    public class MailMessageBuilderTest
    {
        [TestMethod]
        public void TestNullMailMessages()
        {
            // Test incomplete mail builder
            MailMessage mail1 = new MailMessageBuilder().build();
            Assert.AreEqual(null, mail1);

            MailMessage mail2 = new MailMessageBuilder().from(null).to(null).subject(null).body(null).HTMLBody(true).build();
            Assert.AreEqual(null, mail2);

            MailMessage mail3 = new MailMessageBuilder().from("test@test.com").to("othertest@test.com").subject(null).body("Hello Test").HTMLBody(true).build();
            Assert.AreEqual(null, mail3);
        }

        [TestMethod]
        public void TestValidMailMessages()
        {
            // Test valid mail
            string from4 = "test@gmail.com";
            string to4 = "other@gmail.com";
            string subject4 = "Test Subject";
            string body4 = "Hello test";
            bool htmlBody4 = true;

            MailMessage mail4 = new MailMessageBuilder().from(from4).to(to4).subject(subject4).body(body4).HTMLBody(htmlBody4).build();

            Assert.AreNotEqual(null, mail4);
            Assert.AreEqual(from4, mail4.From.Address);
            Assert.AreEqual(true, mail4.To.Contains(new MailAddress(to4)));
            Assert.AreEqual(subject4, mail4.Subject);
            Assert.AreEqual(body4, mail4.Body);
            Assert.AreEqual(htmlBody4, mail4.IsBodyHtml);
        }
    }
}
