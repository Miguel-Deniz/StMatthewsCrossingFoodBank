using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MatthewsCrossingFoodBank
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application
        /// </summary>
        [STAThread]
        static void Main()
        {
            string emailAddress = "matthewcrossingtest1@gmail.com";
            string password = "appleorange123";
            string stmpServer = "smtp.gmail.com";
            int port = 587;
            bool enableSSL = true;
            EmailClient client = new EmailClient(emailAddress, password, stmpServer, port, enableSSL);

            //client.sendEmail("migueldeniz70@gmail.com", "Hello", "This should be in the body of the email.");

            //client.test();

            string body = EmailClient.getHTMLMonetaryDonation();
            client.sendEmail("migueldeniz70@gmail.com", "Test 2", body);
        }
    }
}
