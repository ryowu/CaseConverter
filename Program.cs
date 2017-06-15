using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaseConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            //Get all excel files from current folder
            List<string> files = GetExcelFiles();
            
            //Open each one by one, select * from sheet1 where type <> 'C' and action <> 'wait'

            //Translate into human language, store in DataTable
            DataTable outputData = new DataTable();

            //Create a new folder 'Output'
            //Create new excel and insert the data
            CreateOutputFiles(outputData);
        }

        public static List<string> GetExcelFiles()
        {
            List<string> result = new List<string>();
            return result;
        }

        public static void CreateOutputFiles(DataTable dt)
        { }
    }
}
