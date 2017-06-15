using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaseConverter
{
	public class Utility
	{
		//log
		private static StringBuilder sbLog = new StringBuilder();

		public static List<string> GetExcelFiles()
		{
            return Directory.GetFiles(Environment.CurrentDirectory, "*_SSD.xlsx").ToList();
		}

		public static void CreateOutputFiles(DataTable dt, string outputFolder)
		{
		}

		public static void AddLog(string s)
		{
			Console.WriteLine(s);
			sbLog.AppendLine(s);
		}

        public static void OutputLogToFile()
        {
            string logFile = "cc_" + DateTime.Now.ToString("s").Replace("-","").Replace(":","") + ".log";
            if (File.Exists(logFile))
            {
                File.Delete(logFile);
            }
            try
            {
                using (StreamWriter sw = new StreamWriter(logFile))
                {
                    sw.Write(sbLog.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to write log in file.\n");
                Console.WriteLine(ex.ToString());
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Open each one by one, select * from sheet1 where type <> 'C' and action <> 'wait' ...
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
		public static DataTable ReadInputFile(string filename)
		{
			DataTable result = new DataTable();
			return result;
		}

		public static DataTable TranslateIntoHumanLanguage(DataTable inputData)
		{
			DataTable result = new DataTable();
			return result;
		}
	}
}
