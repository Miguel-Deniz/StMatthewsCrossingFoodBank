using FileHelpers;
using System;

namespace MatthewsCrossingFoodBank
{
    /// <summary>
    /// Handles parsing the data of the application.
    /// </summary>
    class InputParser
    {
        private const int FIELDS_PER_RECORD = 15;
        private static readonly string[] FIELDS = { "Donation_ID", "Donor_is_a_Company", "First_Name", "Last_Name",
                    "Email_Address", "Salutation_Greeting_Dear_So_and_So", "Street_Address", "Apartment", "City_Town",
                    "State_Province", "Zip_Postal_Code", "Donation_Type", "Donated_On", "Amount", "Weight_lbs" };

        public static Donor[] parseFile(string fileName)
        {
            FileHelperEngine<Donor> engine = new FileHelperEngine<Donor>();
            return engine.ReadFile(fileName);
        }

        /// <summary>
        /// Determines whether the given file is in the correct format
        /// </summary>
        /// 
        /// Format:
        ///      1. Donation_ID
        ///      2. Donor_is_a_Company
        ///      3. First_Name
        ///      4. Last_Name
        ///      5. Email_Address
        ///      6. Salutation_Greeting_Dear_So_and_So
        ///      7. Street_Address
        ///      8. Apartment
        ///      9. City_Town
        ///     10. State_Province
        ///     11. Zip_Postal_Code
        ///     12. Donation_Type
        ///     13. Donated_On
        ///     14. Amount
        ///     15. Weight_lbs
        /// 
        /// <param name="fileName"></param>
        /// <returns>bool</returns>
        public static bool isValidFormat(string fileName)
        {
            var detector = new FileHelpers.Detection.SmartFormatDetector();
            var formats = detector.DetectFileFormat(fileName);

            if (formats.Length > 0 && formats[0].Confidence == 100)
            {
                var delimited = formats[0].ClassBuilderAsDelimited;

                // Check for correct amount of fields
                if (delimited.Fields.Length !=  FIELDS.Length)
                {
                    return false;
                }


                // Check correct order of fields
                for (int i = 0; i < FIELDS.Length; i++)
                {
                    if (FIELDS[i] != delimited.Fields[i].FieldName) return false;
                }

                return true;
            }
            else
            {
                return false;
            }  
        }
    }
}
