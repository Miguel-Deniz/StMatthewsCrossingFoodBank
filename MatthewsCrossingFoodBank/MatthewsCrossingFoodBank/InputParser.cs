using FileHelpers;

namespace MatthewsCrossingFoodBank
{
    /// <summary>
    /// Handles parsing the data of the application.
    /// </summary>
    class InputParser
    {
        public static Donor[] parseFile(string dataFile)
        {
            FileHelperEngine<Donor> engine = new FileHelperEngine<Donor>();
            return engine.ReadFile(dataFile);
        }   
    }
}
