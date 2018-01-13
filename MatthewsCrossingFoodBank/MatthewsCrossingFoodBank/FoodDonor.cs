using FileHelpers;

namespace MatthewsCrossingFoodBank
{
    class FoodDonor
    {
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string companyOrganizationName;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string firstName;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string middleName;

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
        public string foodItemCategory;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string nameOfFoodItem;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string quantity;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string quantityType;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string weight;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string value;

        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string memo;
    }
}
}
