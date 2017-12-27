using FileHelpers;

namespace MatthewsCrossingFoodBank
{
    /// <summary>
    /// The mapping class for a Donor.
    /// </summary>
    [DelimitedRecord(",")]
    public class Donor
    {
        public string donationID;

        public string firstName;

        public string lastName;

        public string streetAddress;

        public string apartment;

        public string cityTown;

        public string stateProvince;

        public string zipPostalCode;

        public string email;
    }
}
