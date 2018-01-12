using Microsoft.Office.Interop.Word;
using System;
using System.Net.Mail;
using System.Windows.Forms;
//using GemBox.Document;

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

        public void sendEmail(string toEmail, string subject, string body)
        {
            System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
            SmtpClient SmtpServer = new SmtpClient(_stmpServer);

            mail.From = new MailAddress(_emailAddress);
            mail.To.Add(toEmail);
            mail.Subject = subject;
            mail.Body = body;
            mail.IsBodyHtml = true;

            SmtpServer.Port = _port;
            SmtpServer.Credentials = new System.Net.NetworkCredential(_emailAddress, _password);
            SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);
            MessageBox.Show("Mail Sent!");
        }

        public void test()
        {
            object oMissing = System.Reflection.Missing.Value;
            object oTemplatePath = "C:\\Users\\Miguel\\Desktop\\Temp\\Money_TY_Letter_Regular_Donor_2017.docx";

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = new Microsoft.Office.Interop.Word.Document();

            wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            Console.WriteLine("Total fields: {0} ", wordDoc.Fields.Count);

            foreach (Field myMergeField in wordDoc.Fields)
            {
                Range rngFieldCode = myMergeField.Code;
                string fieldText = rngFieldCode.Text;

                if (fieldText.StartsWith(" MERGEFIELD"))
                {
                    fieldText = fieldText.Remove(0, 12);
                    fieldText = fieldText.Replace(" ", "");
                    Console.WriteLine(fieldText);

                    switch (fieldText)
                    {
                        case "First_Name":
                            myMergeField.Select();
                            wordApp.Selection.TypeText("Miguel");
                            break;
                        case "Last_Name":
                            myMergeField.Select();
                            wordApp.Selection.TypeText("LastNameHere");
                            break;
                        case "Street_Address":
                            myMergeField.Select();
                            wordApp.Selection.TypeText("911 Mill Ave");
                            break;
                        case "Apartment":
                            myMergeField.Select();
                            wordApp.Selection.TypeText("Apt 500");
                            break;
                        case "CityTown":
                            myMergeField.Select();
                            wordApp.Selection.TypeText("Tempe");
                            break;
                        case "StateProvince":
                            myMergeField.Select();
                            wordApp.Selection.TypeText("AZ");
                            break;
                        case "ZipPostal_Code":
                            myMergeField.Select();
                            wordApp.Selection.TypeText("98222");
                            break;
                        default:
                            Console.WriteLine("Unhandled case");
                            // Error
                            break;
                    }
                }
            }

            string name = "C:\\Users\\Miguel\\Desktop\\Temp\\populo.docx";
            wordDoc.SaveAs2(name);

            wordDoc.Close();
            wordApp.Application.Quit();

            // Send
            
        }

        public static string getHTMLMonetaryDonation()
        {
            return @"<p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt; text-align: center;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;""><img src=""https://lh3.googleusercontent.com/jsqClq4FaMGDXuVBAEWLH4WaJZvnzm-sjgF3BepNccxVKF5Q6fsDFoo0BtpOl5RC9qbreNRhZYUGdy6c8-dkoApIksC1lwAmZ6PTTPadO807b-JGRD6FIPXFso6mbuM9I3wnQZIKACLCel9i-g"" width=""372"" height=""88"" style=""border: none; transform: rotate(0rad);""></span></p><br><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">«CURRENT_DATE»</span></p><br><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">«First_Name» «Last_Name»</span></p><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">«Street_Address»</span></p><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">«Apartment»</span></p><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">«CityTown» «StateProvince» &nbsp;«ZipPostal_Code»</span></p><br><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt; text-align: center;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; font-weight: 700; vertical-align: baseline; white-space: pre-wrap;"">«Salutation_Greeting_Dear_So_and_So»:</span></p><br><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt; text-align: center;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; font-weight: 700; vertical-align: baseline; white-space: pre-wrap;"">WE JUST WANT TO THANK YOU.</span></p><br><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt; text-align: justify;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">Thank you for showing your commitment to fighting hunger in our community by sending your generous gift of $«M_Amount» received on «Donated_On».</span></p><br><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">We are absolutely thrilled that you chose to support Matthew’s Crossing. &nbsp;Because of your gift families will be able to feed their children. &nbsp;</span></p><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">We at Matthew’s Crossing continue to provide food assistance to individuals and families in need, specifically the working poor, children, seniors and individuals with disabilities on a fixed income, families in crisis and the homeless. </span></p><br><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">We are seeing record numbers of new clients needing our help and truly, we couldn't continue to provide our critical services without your support.</span></p><br><ul style=""margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt;""><li dir=""ltr"" style=""list-style-type: disc; font-size: 11pt; font-family: &quot;Noto Sans Symbols&quot;; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline;""><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt;""><span style=""font-size: 11pt; font-family: Calibri; background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">Matthew's Crossing is a 501 (C) (3) tax-exempt organization (EIN 55-0896414) under the provisions of the Internal Revenue Code. &nbsp;</span></p></li><li dir=""ltr"" style=""list-style-type: disc; font-size: 11pt; font-family: &quot;Noto Sans Symbols&quot;; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline;""><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt;""><span style=""font-size: 11pt; font-family: Calibri; background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">We acknowledge that your gift was received in our office and credited as a tax-deductible contribution for the calendar year 2017. </span></p></li><li dir=""ltr"" style=""list-style-type: disc; font-size: 11pt; font-family: &quot;Noto Sans Symbols&quot;; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline;""><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt;""><span style=""font-size: 11pt; font-family: Calibri; background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">No goods or services were received in exchange for this donation, other than joy of giving to an organization that is helping to fight hunger in our community. </span></p></li><li dir=""ltr"" style=""list-style-type: disc; font-size: 11pt; font-family: &quot;Noto Sans Symbols&quot;; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline;""><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt;""><span style=""font-size: 11pt; font-family: Calibri; background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">Matthew's Crossing is included on the approved list of qualified charitable organizations published by the Arizona Department of Revenue and this contribution qualified for a charitable Arizona tax credit for the working poor. &nbsp;The allowable tax credit is up to $400 for single and up to $800 for joint tax returns. &nbsp;Please consult the Arizona Department of Revenue at </span><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 255); background-color: transparent; text-decoration-line: underline; vertical-align: baseline; white-space: pre-wrap;"">www.azdor.gov</span><span style=""font-size: 11pt; font-family: Calibri; background-color: transparent; vertical-align: baseline; white-space: pre-wrap;""> for more information.</span></p></li></ul><br><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt; text-align: justify;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">Please call if you have any questions. &nbsp;I can’t express enough how much we appreciate your support.</span></p><br><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt; text-align: justify;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">Sincerely,</span></p><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt; text-align: justify;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;""><img src=""https://lh5.googleusercontent.com/LFJKF5Sb-bLR30GSioqyrlnXHZKE1w6hMiwrZrtETnjWato2lVHtaWyiQYaX4gx2-4m_TybfFrimGFtKMEl64H3UOOPsRDmW3arkrCzsqywOKGLoqN1SJUFp-9HdTO6cNadsx2aUz7ZwiqC-mA"" width=""118"" height=""57"" style=""border: none; transform: rotate(0rad);"" alt=""C:\Users\Business Manager\AppData\Local\Microsoft\Windows\INetCache\Content.Word\Signature.jpg""></span></p><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt; text-align: justify;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">Jan Terhune</span></p><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt; text-align: justify;""><span style=""font-size: 11pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">Executive Director</span></p><br><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt; text-align: center; border-top: 0.5pt solid rgb(0, 0, 0); padding: 4pt 0pt 0pt;""><span style=""font-size: 8pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; vertical-align: baseline; white-space: pre-wrap;"">Matthew’s Crossing Food Bank ● 1368 N Arizona Ave Suite 112, Chandler, AZ 85225 ● (480) 857-2296 ● matthewscrossing.org</span></p><p dir=""ltr"" style=""line-height: 1.2; margin-top: 0pt; margin-bottom: 0p; margin-top: 0pt; margin-bottom: 0pt; margin-left: 25pt; margin-right: 25pt; text-align: center;""><span style=""font-size: 8pt; font-family: Calibri; color: rgb(204, 0, 0); background-color: transparent; font-style: italic; vertical-align: baseline; white-space: pre-wrap;"">Compassion</span><span style=""font-size: 8pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; font-style: italic; vertical-align: baseline; white-space: pre-wrap;""> </span><span style=""font-size: 8pt; font-family: Calibri; color: rgb(0, 176, 80); background-color: transparent; font-style: italic; vertical-align: baseline; white-space: pre-wrap;"">●</span><span style=""font-size: 8pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; font-style: italic; vertical-align: baseline; white-space: pre-wrap;""> </span><span style=""font-size: 8pt; font-family: Calibri; color: rgb(204, 0, 0); background-color: transparent; font-style: italic; vertical-align: baseline; white-space: pre-wrap;"">Dignity</span><span style=""font-size: 8pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; font-style: italic; vertical-align: baseline; white-space: pre-wrap;""> </span><span style=""font-size: 8pt; font-family: Calibri; color: rgb(0, 176, 80); background-color: transparent; font-style: italic; vertical-align: baseline; white-space: pre-wrap;"">●</span><span style=""font-size: 8pt; font-family: Calibri; color: rgb(0, 0, 0); background-color: transparent; font-style: italic; vertical-align: baseline; white-space: pre-wrap;""> </span><span style=""font-size: 8pt; font-family: Calibri; color: rgb(204, 0, 0); background-color: transparent; font-style: italic; vertical-align: baseline; white-space: pre-wrap;"">Hope</span></p><div><span style=""font-size: 8pt; font-family: Calibri; color: rgb(204, 0, 0); background-color: transparent; font-style: italic; vertical-align: baseline; white-space: pre-wrap;""><br></span></div></span></div>";
        }

        //the method to read html document
        public string RetrieveStream(string FullName)
        {
            string stream = null;
            using (System.IO.StreamReader sr = new System.IO.StreamReader(FullName))
            {
                string line = null;
                while ((line = sr.ReadLine()) != null)
                {
                    stream += line;
                }
            }

            return stream;
        }
    }
}
