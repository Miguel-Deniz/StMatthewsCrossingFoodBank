using ExcelDataReader;
using FileHelpers;
using System;
using System.Data;
using System.IO;

namespace MatthewsCrossingFoodBank
{
    /// <summary>
    /// Handles parsing the data of the application.
    /// </summary>
    class InputParser
    {
        private const int FIELDS_PER_RECORD = 15;

        public static MonetaryDonor[] parseMonetaryFile(string fileName)
        {
            FileHelperEngine<MonetaryDonor> engine = new FileHelperEngine<MonetaryDonor>();
            
            return engine.ReadFile(fileName);
        }

        
        /// <summary>
        ///     Converts the Excel sourceFile into a CSV destinationFile.
        ///     The Excel file must contain at least 1 sheet.
        /// </summary>
        /// <param name="sourceFile"></param>
        /// <param name="destinationFile"></param>
        public static void convertXLSXFileToCSV(string sourceFile, string destinationFile)
        {
            FileStream stream = File.Open(sourceFile, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            DataSet result = excelReader.AsDataSet();

            // Check if Excel file contains at least one sheet
            if (result.Tables.Count < 1)
            {
                throw new System.InvalidOperationException("Excel file does not contain at least one sheet.");
            }

            string csvData = "";

            for (int row = 0; row < result.Tables[0].Rows.Count; row++)
            {
                for (int col = 0; col < result.Tables[0].Columns.Count; col++)
                {
                    if (col == 0)
                        csvData += "\"" + result.Tables[0].Rows[row][col].ToString() + "\"";
                    else
                        csvData += ",\"" + result.Tables[0].Rows[row][col].ToString() + "\"";
                }

                csvData += "\n";
            }

            string output = destinationFile;
            StreamWriter csv = new StreamWriter(@output, false);
            csv.Write(csvData);
            csv.Close();
            excelReader.Close();
        }
    }
}
