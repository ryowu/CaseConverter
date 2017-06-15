using System;
using System.Collections.Generic;
using System.Data;
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
			List<string> result = new List<string>();
			return result;
		}

		public static void CreateOutputFiles(DataTable dt, string outputFolder)
		{
		}

		public static void AddLog(string s)
		{
			Console.WriteLine(s);
			sbLog.AppendLine(s);
		}

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
