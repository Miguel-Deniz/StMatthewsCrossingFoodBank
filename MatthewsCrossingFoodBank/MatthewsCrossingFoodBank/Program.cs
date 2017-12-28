using System;
using System.Collections.Generic;
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
           // Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());

            bool ans = InputParser.isValidFormat("C:\\Users\\Miguel\\Desktop\\testdata.csv");
            Console.WriteLine("Answer: " + ans);
        }
    }
}
