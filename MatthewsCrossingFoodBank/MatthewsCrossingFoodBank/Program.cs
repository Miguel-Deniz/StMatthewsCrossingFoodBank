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



            MonetaryDonor donor = new MonetaryDonor();

            donor.firstName = "Jason";
            donor.lastName = "Gaytan";
            donor.streetAddress = "911 S. Mill Ave";
            donor.cityTown = "Tempe";
            donor.stateProvince = "AZ";
            donor.zipPostalCode = "85211";
            donor.salutationGreeting = "Dear Jason";
            donor.amount = "0.00";
            donor.donatedOn = DateTime.Now.ToString("MMMM dd, yyyy");

            string body = HTMLLetter.getHTML(donor);
            client.sendEmail("jasongaytan10@gmail.com", "Test", body);
        }
    }
}
