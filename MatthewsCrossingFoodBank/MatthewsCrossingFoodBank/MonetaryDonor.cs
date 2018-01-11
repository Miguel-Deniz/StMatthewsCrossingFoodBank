using FileHelpers;

namespace MatthewsCrossingFoodBank
{
    /// <summary>
    /// The mapping class for a Donor.
    /// </summary>

    [IgnoreFirst]
    [DelimitedRecord(",")]
    public class MonetaryDonor
    {
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string donationID;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string company;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string firstName;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string lastName;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string emailAddress;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string spouseName;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string salutationGreeting;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string streetAddress;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string apartment;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string cityTown;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string stateProvince;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string zipPostalCode;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string donorType;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string donationType;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string sourceOfDonation;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string donatedOn;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string designatedFund;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string amount;
    }
}
