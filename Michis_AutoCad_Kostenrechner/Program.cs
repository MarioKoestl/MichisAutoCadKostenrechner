using System;
using System.Windows.Forms;

namespace ConsoleApp
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            
            var fileNameObject = new GetFileName().Value();
            var calcuationEntries = new ExtractDataFromDWG().Value(fileNameObject.FullFilePath);
            new ExportToExcel().Run(fileNameObject, calcuationEntries);
            Console.ReadLine();
        }
    }
}
 