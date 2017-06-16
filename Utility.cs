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

		private static void ExportToExcel(DataTable dt, string filepath)
		{

			/*Set up work book, work sheets, and excel application*/
			Microsoft.Office.Interop.Excel.Application oexcel = new Microsoft.Office.Interop.Excel.Application();
			try
			{
				string path = AppDomain.CurrentDomain.BaseDirectory;
				object misValue = System.Reflection.Missing.Value;
				Microsoft.Office.Interop.Excel.Workbook obook = oexcel.Workbooks.Add(misValue);
				Microsoft.Office.Interop.Excel.Worksheet osheet = new Microsoft.Office.Interop.Excel.Worksheet();


				//  obook.Worksheets.Add(misValue);

				osheet = (Microsoft.Office.Interop.Excel.Worksheet)obook.Sheets["Test"];
				int colIndex = 0;
				int rowIndex = 1;

				foreach (DataColumn dc in dt.Columns)
				{
					colIndex++;
					osheet.Cells[1, colIndex] = dc.ColumnName;
				}
				foreach (DataRow dr in dt.Rows)
				{
					rowIndex++;
					colIndex = 0;

					foreach (DataColumn dc in dt.Columns)
					{
						colIndex++;
						osheet.Cells[rowIndex, colIndex] = dr[dc.ColumnName];
					}
				}

				osheet.Columns.AutoFit();

				//Release and terminate excel

				obook.SaveAs(filepath);
				obook.Close();
				oexcel.Quit();
				ReleaseObject(osheet);

				ReleaseObject(obook);

				ReleaseObject(oexcel);
				GC.Collect();
			}
			catch (Exception ex)
			{
				oexcel.Quit();
			}
		}

		private static void ReleaseObject(object o) { try { while (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0) { } } catch { } finally { o = null; } }

		public static void CreateOutputFiles(DataTable dt, string outputFolder)
		{
			string outputfileBaseName = "outputBase.xls";
			string outputFilename = Path.Combine(outputFolder, dt.TableName);
			//create new file
			File.Copy(outputfileBaseName, outputFilename);

			ExportToExcel(dt, outputFilename);
		}

		public static void AddLog(string s)
		{
			Console.WriteLine(s);
			sbLog.AppendLine(s);
		}

		public static void OutputLogToFile()
		{
			string logFile = "cc_" + DateTime.Now.ToString("s").Replace("-", "").Replace(":", "") + ".log";
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
				Console.WriteLine("Failed to write log in file.\r\n");
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
                dataTable.TableName = Path.GetFileName(filename);
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
                try
                {
                    col4LowerValue = row[4].ToString().Trim().ToLower();

                    #region skip specific rows
                    if (row[1].ToString().Trim().Equals("C", StringComparison.OrdinalIgnoreCase) ||
                        col4LowerValue.StartsWith("wait") ||
                        col4LowerValue.StartsWith("handleajax") ||
                        col4LowerValue.StartsWith("basicmode") ||
                        col4LowerValue.StartsWith("closebrowser") ||
                        col4LowerValue.StartsWith("windowclose") ||
                        col4LowerValue.StartsWith("comment") ||
                        col4LowerValue.StartsWith("portal") ||
                        col4LowerValue.StartsWith("attach") ||
                        col4LowerValue.StartsWith("storemessage"))
                    {
                        continue;
                    }
                    #endregion

                    col5RawValue = row[5].ToString();
                    col6RawValue = row[6].ToString();
                    col6TranslatedValue = TranslateColumn5Value(col6RawValue);

                    #region verifications
                    if (col4LowerValue.StartsWith("Verify", StringComparison.OrdinalIgnoreCase))
                    {
                        sbVerifications.AppendLine(TranslateCheckPointString(col5RawValue, col6RawValue,
                            col4LowerValue.StartsWith("VerifyNo", StringComparison.OrdinalIgnoreCase)));
                        if (stepFlag)
                        {
                            // insert all step strings
                            result.Rows.Add(sbSteps.ToString().TrimEnd(), "");
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
                        else if (col4LowerValue.Equals("LaunchMobileApp", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Launch Aprimo Mobile");
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
                        else if (col4LowerValue.StartsWith("ClickLink", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Click link \"" + col5RawValue + "\"");
                        }
                        else if (col4LowerValue.Equals("SelectItem", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Select \"" + col5RawValue + "\"");
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(col4LowerValue) && string.IsNullOrEmpty(col5RawValue))
                            {
                                //AddLog("The values in colum 4 and 5 are both empty, " + string.Join(";", row.ItemArray));
                            }
                            else
                            {
                                throw new Exception("Unhandled col4LowerValue:" + col4LowerValue);
                            }
                        }

                        if (verifyFlag)
                        {
                            // insert all verification strings
                            result.Rows[result.Rows.Count - 1][1] = sbVerifications.ToString().TrimEnd();
                            sbVerifications.Clear();
                        }
                        verifyFlag = false;
                        stepFlag = true;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    throw new Exception("Found exception, row number:" + row[0].ToString(), ex);
                }
            }
            result.TableName = inputData.TableName;
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

        private static string TranslateCheckPointString(string col5Value, string col6Value, bool isVerifyNo)
        {
            string[] array = col6Value.Split(':');
            string result = null;
            if (array.Length > 1)
            {
                result = array[0] + ": \"" + TranslateColumn5Value(array[1]) + "\"";
            }
            else
            {
                result = "\"" + TranslateColumn5Value(array[0]) + "\"";
            }
            //string tmp = col5Value;
            //if (col5Value.Equals("PageContent", StringComparison.OrdinalIgnoreCase))
            //{
            //    tmp = "page content";
            //}
            //else if (col5Value.Equals("PageTitle", StringComparison.OrdinalIgnoreCase))
            //{
            //    tmp = "page title";
            //}
            //return "Verify " + tmp + ": " + result + (isVerifyNo? " IS NOT displayed" : " is displayed");
            return "Verify " + result + (isVerifyNo ? " IS NOT displaying" : " is displaying");
        }

		private static string GetExcelConnectionString(string fileName)
		{
			//return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 12.0;";
			return @"Provider=Microsoft.Jet.OLEDB.4.0;" +
				   @"Data Source=" + fileName + ";" +
                   @"Extended Properties='Excel 8.0;IMEX=1;'";
		}
	}
}
