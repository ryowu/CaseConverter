using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseConverter
{
    public class Utility
    {
        public const string StrPrefixOfEnterData = "\u2022 ";
        public const string StrPrefixOfVerifyData = "\u2023 ";

        //log
        private static StringBuilder sbLog = new StringBuilder();

        public static List<string> GetExcelFiles()
        {
            return Directory.GetFiles(Environment.CurrentDirectory, "*_SSD.xls*").ToList();
        }

        public static void CreateOutputFiles(DataTable dt, string outputFolder)
        {
            string outputfileBaseName = Path.Combine(Environment.CurrentDirectory, "outputBase.xls");
            string outputFilename = Path.Combine(outputFolder,
                Program.OutputFilePrefix + (dt.TableName.ToLower().EndsWith(".xlsx") ? dt.TableName.Substring(0, dt.TableName.Length - 1) : dt.TableName));
            //create new file
            try
            {
                if (File.Exists(outputFilename))
                {
                    File.Delete(outputFilename);
                }
                File.Copy(outputfileBaseName, outputFilename);
            }
            catch (Exception ex)
            {
                throw new Exception("Delete file or copy file template error, new file name path:\n" + outputFilename, ex);
            }

            ExportToExcel(dt, outputFilename);
        }

        private static void ExportToExcel(DataTable dt, string filePath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook excelWorkBook = null;
            try
            {
                //Creae an Excel application instance
                excelApp = new Excel.Application();

                //Create an Excel workbook instance and open it from the predefined location
                excelWorkBook = excelApp.Workbooks.Open(filePath);

                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = Program.DefaultSheetName;

                for (int i = 1; i < dt.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = dt.Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    for (int k = 0; k < dt.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = dt.Rows[j].ItemArray[k].ToString();
                    }
                }

                excelWorkBook.Save();
            }
            catch (Exception ex)
            {
                throw new Exception("Export DateSet to Excel error, new file name path:\n" + filePath, ex);
            }
            finally
            {
                if (excelWorkBook != null)
                {
                    excelWorkBook.Close();
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }
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
                var dataAdapter = new OleDbDataAdapter("SELECT * FROM [" + Program.DefaultSheetName + "$]", excelConnection);
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
            result.Columns.Add(new DataColumn(Program.ColName0, typeof(string)));
            result.Columns.Add(new DataColumn(Program.ColName1, typeof(string)));

            StringBuilder sbSteps = new StringBuilder();
            //StringBuilder sbVerifications = new StringBuilder();
            //string col3RawValue = null;
            string col4LowerValue = null;
            string col5RawValue = null;
            string col6RawValue = null;
            string col6TranslatedValue = null;
            bool stepFlag = false;  // true means last row is for steps
            bool verifyFlag = false; // true means last row is of verifications
            bool isClickItemStep = false;
            string clickStr = null;
            string tempStr = null;
            //string pageName = null;
            List<string> verifyStrs = new List<string>();
            List<string> verifyNoStrs = new List<string>();

            foreach (DataRow row in inputData.Rows)
            {
                try
                {
                    col4LowerValue = row[4].ToString().Trim().ToLower();

                    #region skip specific rows
                    if (row[1].ToString().Trim().Equals("C", StringComparison.OrdinalIgnoreCase) ||
                        Program.SkipActions.Any(le => col4LowerValue.StartsWith(le)))
                    {
                        continue;
                    }
                    #endregion

                    //col3RawValue = row[3].ToString();
                    col5RawValue = row[5].ToString();
                    col6RawValue = row[6].ToString();
                    col6TranslatedValue = TranslateColumn5Value(col6RawValue);

                    isClickItemStep = col4LowerValue.StartsWith("ClickItem", StringComparison.OrdinalIgnoreCase);

                    // if the last step is a click item step and the current step is not a click step
                    // then write this click step. E.g. Click btn1 -> btn2 -> ...
                    if (clickStr != null && !isClickItemStep)
                    {
                        sbSteps.AppendLine(clickStr);
                        clickStr = null;
                    }

                    #region verifications
                    if (col4LowerValue.StartsWith("Verify", StringComparison.OrdinalIgnoreCase) ||
                        col4LowerValue.StartsWith("ChkPropIncludes", StringComparison.OrdinalIgnoreCase))
                    {
                        tempStr = StrPrefixOfVerifyData + TranslateCheckPointString(col4LowerValue, col5RawValue, col6RawValue);
                        if (col4LowerValue.StartsWith("VerifyNo", StringComparison.OrdinalIgnoreCase))
                        {
                            verifyNoStrs.Add(tempStr);
                        }
                        else
                        {
                            verifyStrs.Add(tempStr);
                        }

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
                        if (col4LowerValue.Equals("CloseBrowser", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Close browser");
                        }
                        else if (col4LowerValue.Equals("WindowClose", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Close the current window");
                        }
                        else if (col4LowerValue.Equals("CloseBrowserTitled", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Close browser which has the title:\"" + col5RawValue + "\"");
                        }
                        else if (col4LowerValue.StartsWith("ExportDownload", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Export to download");
                        }
                        else if (col4LowerValue.Equals("LaunchApp", StringComparison.OrdinalIgnoreCase))
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
                        else if (col4LowerValue.Equals("EnterData", StringComparison.OrdinalIgnoreCase) ||
                            col4LowerValue.Equals("EnterDataNoTab", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine(StrPrefixOfEnterData + col5RawValue + ": " + col6TranslatedValue);
                        }
                        else if (col4LowerValue.Equals("EnterDataWithEnter", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine(StrPrefixOfEnterData + col5RawValue + ": " + col6TranslatedValue + " (enter data then press Key Enter)");
                        }
                        else if (col4LowerValue.Equals("EnterComboText", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Select combobox, set \"" + col5RawValue + "\" as \"" + col6TranslatedValue + "\"");
                        }
                        else if (isClickItemStep)
                        {
                            // do not write step for clicking "Login"
                            if (!col5RawValue.Equals("Login", StringComparison.OrdinalIgnoreCase))
                            {
                                if (clickStr != null)
                                {
                                    clickStr += " -> " + col5RawValue;
                                }
                                else
                                {
                                    clickStr = "Click " + col5RawValue;
                                }
                            }
                        }
                        else if (col4LowerValue.StartsWith("ClickLink", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Click link \"" + col5RawValue + "\"");
                        }
                        else if (col4LowerValue.Equals("SelectItem", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Select \"" + col5RawValue + "\"" + " as \"" + col6TranslatedValue + "\"");
                        }
                        else if (col4LowerValue.Equals("SelectTree", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Select tree " + col5RawValue + " \"" + col6TranslatedValue + "\"");
                        }
                        else if (col4LowerValue.StartsWith("DragNextTo", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Drag next to \"" + col5RawValue + "\"");
                        }
                        else if (col4LowerValue.StartsWith("DragOnTo", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Drag \"" + col5RawValue + "\"");
                        }
                        else if (col4LowerValue.StartsWith("NodeProcessed", StringComparison.OrdinalIgnoreCase))
                        {
                            sbSteps.AppendLine("Node processed \"" + col5RawValue + "\"");
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
                            result.Rows[result.Rows.Count - 1][1] = GenerateVerifyStrings(verifyStrs, verifyNoStrs);
                            verifyStrs.Clear();
                            verifyNoStrs.Clear();
                        }
                        verifyFlag = false;
                        stepFlag = true;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    bool notThrow = false;
                    if (ex.Message != null && ex.Message.Contains("Unhandled col4LowerValue:"))
                    {
                        Match m = Regex.Match(ex.Message, "(?<=Unhandled col4LowerValue:)\\w*tree\\w*");
                        if (m.Success)
                        {
                            AddLog("Found col4LowerValue:" + m.Value + ", move original file to folder \"FilesWithIssue\". Will skip this row and continue to process the following rows of this file.");
                            notThrow = true;
                            string dirFilesWithIssue = Path.Combine(Environment.CurrentDirectory, "FilesWithIssue");
                            if (!Directory.Exists(dirFilesWithIssue))
                            {
                                Directory.CreateDirectory(dirFilesWithIssue);
                            }
                            if (File.Exists(Path.Combine(dirFilesWithIssue, inputData.TableName)))
                            {
                                File.Delete(Path.Combine(dirFilesWithIssue, inputData.TableName));
                            }
                            File.Copy(Path.Combine(Environment.CurrentDirectory, inputData.TableName), Path.Combine(dirFilesWithIssue, inputData.TableName));
                            File.Delete(Path.Combine(Environment.CurrentDirectory, inputData.TableName));
                        }
                    }

                    if (!notThrow)
                    {
                        throw new Exception("Found exception, row number:" + row[0].ToString(), ex);
                    }
                }

                // update the current page name
                //pageName = col3RawValue;
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

        private static string TranslateCheckPointString(string col4Value, string col5Value, string col6Value)
        {
            string[] array = col6Value.Split(':');
            string result = null;
            if (array.Length > 1)
            {
                result = array[0] + ": " + TranslateColumn5Value(array[1]);
            }
            else
            {
                result = col5Value + ": " + TranslateColumn5Value(array[0]);
            }
            if (string.IsNullOrEmpty(col6Value))
            {
                result = col5Value;
            }
            return (col4Value.Equals("VerifyLink", StringComparison.OrdinalIgnoreCase) ? "link " : "") + result;
        }

        private static string GenerateVerifyStrings(List<string> verifyStrs, List<string> verifyNoStrs)
        {
            StringBuilder sbVerifications = new StringBuilder();
            string str = "Verify the following value";
            if (verifyStrs.Count > 0)
            {
                sbVerifications.AppendLine(str + (verifyStrs.Count == 1 ? " is" : "s are") + " displaying:");
                verifyStrs.ForEach(s => sbVerifications.AppendLine(s));
                sbVerifications.AppendLine("");
            }
            if (verifyNoStrs.Count > 0)
            {
                sbVerifications.AppendLine(str + (verifyNoStrs.Count == 1 ? " is" : "s are") + " NOT displaying:");
                verifyNoStrs.ForEach(s => sbVerifications.AppendLine(s));
            }
            return sbVerifications.ToString().TrimEnd();
        }

        private static string GetExcelConnectionString(string fileName)
        {
            if (fileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            {
                return @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;IMEX=1;'";
            }
            else if (fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;IMEX=1;'";
            }
            else
            {
                throw new Exception("Excel format is not supported, file name:" + fileName);
            }
        }
    }
}
