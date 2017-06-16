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
            return Directory.GetFiles(Environment.CurrentDirectory, "*_SSD.xls").ToList();
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
        /// Open each one by one, select * from sheet1
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static DataTable ReadInputFile(string filename)
        {
            var excelConnectionString = GetExcelConnectionString(filename);
            var dataTable = new DataTable();

            using (var excelConnection = new OleDbConnection(excelConnectionString))
            {
                excelConnection.Open();
                var dataAdapter = new OleDbDataAdapter("SELECT * FROM [Test$]", excelConnection);
                dataAdapter.Fill(dataTable);
                excelConnection.Close();
            }
            AddLog("OpenExcelFile: File successfully opened:" + filename);

            return dataTable;
        }

        public static DataTable TranslateIntoHumanLanguage(DataTable inputData)
        {
            DataTable result = new DataTable();
            result.Columns.Add(new DataColumn("steps", typeof(string)));
            result.Columns.Add(new DataColumn("verifications", typeof(string)));

            StringBuilder sbSteps = new StringBuilder();
            StringBuilder sbVerifications = new StringBuilder();
            string col4LowerValue = null;
            string col5RawValue = null;
            string col6RawValue = null;
            string col6TranslatedValue = null;
            bool stepFlag = false;  // true means last row is for steps
            bool verifyFlag = false; // true means last row is of verifications

            foreach (DataRow row in inputData.Rows)
            {
                col4LowerValue = row[4].ToString().Trim().ToLower();

                #region skip specific rows
                if (row[1].ToString().Trim().Equals("C",StringComparison.OrdinalIgnoreCase) ||
                    col4LowerValue.StartsWith("wait") ||
                    col4LowerValue.StartsWith("handleajax") ||
                    col4LowerValue.StartsWith("basicmode") ||
                    col4LowerValue.StartsWith("closebrowser") ||
                    col4LowerValue.StartsWith("windowclose") ||
                    col4LowerValue.StartsWith("comment") ||
                    col4LowerValue.StartsWith("portal") ||
                    col4LowerValue.StartsWith("attach"))
                {
                    continue;
                }
                #endregion

                col5RawValue = row[5].ToString();
                col6RawValue = row[6].ToString();
                col6TranslatedValue = TranslateColumn5Value(col6RawValue);

                #region verifications
                if (col4LowerValue.Equals("Verify", StringComparison.OrdinalIgnoreCase))
                {
                    sbVerifications.AppendLine(TranslateCheckPointString(col5RawValue, col6RawValue));
                    if (stepFlag)
                    {
                        // insert all step strings
                        result.Rows.Add(sbSteps.ToString(), "");
                        sbSteps.Clear();
                    }
                    verifyFlag = true;
                    stepFlag = false;
                }
                #endregion
                #region steps
                else
                {
                    if (col4LowerValue.Equals("LaunchApp", StringComparison.OrdinalIgnoreCase))
                    {
                        sbSteps.AppendLine("Launch Aprimo Marketing");
                    }
                    else if (col4LowerValue.Equals("LaunchPortal", StringComparison.OrdinalIgnoreCase))
                    {
                        sbSteps.AppendLine("Launch Aprimo Portal");
                    }
                    else if (col4LowerValue.Equals("LogOutExit", StringComparison.OrdinalIgnoreCase))
                    {
                        sbSteps.AppendLine("Log off current user");
                    }
                    else if (col4LowerValue.Equals("EnterData", StringComparison.OrdinalIgnoreCase))
                    {
                        sbSteps.AppendLine("Enter data, set \"" + col5RawValue + "\" as \"" + col6TranslatedValue + "\"");
                    }
                    else if (col4LowerValue.Equals("EnterComboText", StringComparison.OrdinalIgnoreCase))
                    {
                        sbSteps.AppendLine("Select combobox, set \"" + col5RawValue + "\" as \"" + col6TranslatedValue + "\"");
                    }
                    else if (col4LowerValue.StartsWith("ClickItem", StringComparison.OrdinalIgnoreCase))
                    {
                        sbSteps.AppendLine("Click \"" + col5RawValue + "\"");
                    }
                    else if (col4LowerValue.Equals("SelectItem", StringComparison.OrdinalIgnoreCase))
                    {
                        sbSteps.AppendLine("Select \"" + col5RawValue + "\"");
                    }
                    else
                    {
                        throw new Exception("Unhandled col4LowerValue:" + col4LowerValue);
                    }

                    if (verifyFlag)
                    {
                        // insert all verification strings
                        result.Rows[result.Rows.Count - 1][1] = sbVerifications.ToString();
                        sbVerifications.Clear();
                    }
                    verifyFlag = false;
                    stepFlag = true;
                }
                #endregion
            }
            return result;
        }

        private static string TranslateColumn5Value(string input)
        {
            string result = input;
            int index = result.IndexOf('|');
            if (index >= 0)
            {
                result = result.Substring(index + 1);
            }
            result = result.Replace("<$", "").Replace("$>", "");
            return result;
        }

        private static string TranslateCheckPointString(string col5Value, string col6Value)
        {
            string[] array = col6Value.Split(':');
            string result = null;
            if (array.Length > 1)
            {
                result = array[0] + ": " + TranslateColumn5Value(array[1]);
            }
            else
            {
                result = array[0];
            }
            string tmp = col5Value;
            if (col5Value.Equals("PageContent", StringComparison.OrdinalIgnoreCase))
            {
                tmp = "page content";
            }
            else if (col5Value.Equals("PageTitle", StringComparison.OrdinalIgnoreCase))
            {
                tmp = "page title";
            }

            return "Verify " + tmp + ": " + result;
        }

        private static string GetExcelConnectionString(string fileName)
        {
            //return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 12.0;";
            return @"Provider=Microsoft.Jet.OLEDB.4.0;" +
                   @"Data Source=" + fileName + ";" +
                   @"Extended Properties=" + Convert.ToChar(34).ToString() +
                   @"Excel 8.0" + Convert.ToChar(34).ToString() + ";";
        }
    }
}
