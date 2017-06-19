using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

namespace CaseConverter
{
    class Program
    {
        public const string OutputBaseFileName = "outputBase.xls";
        private static string outputDirName;
        public static string OutputDirName
        {
            get
            {
                if (outputDirName == null)
                {
                    outputDirName = ConfigurationManager.AppSettings["outputDirName"];
                }
                return outputDirName;
            }
        }
        private static string outputFilePrefix;
        public static string OutputFilePrefix
        {
            get
            {
                if (outputFilePrefix == null)
                {
                    outputFilePrefix = ConfigurationManager.AppSettings["outputFilePrefix"];
                }
                return outputFilePrefix;
            }
        }
        private static string colName0;
        public static string ColName0
        {
            get
            {
                if (colName0 == null)
                {
                    colName0 = ConfigurationManager.AppSettings["colName0"];
                }
                return colName0;
            }
        }
        private static string colName1;
        public static string ColName1
        {
            get
            {
                if (colName1 == null)
                {
                    colName1 = ConfigurationManager.AppSettings["colName1"];
                }
                return colName1;
            }
        }
        public static string defaultSheetName;
        public static string DefaultSheetName
        {
            get
            {
                if (defaultSheetName == null)
                {
                    defaultSheetName = ConfigurationManager.AppSettings["defaultSheetName"];
                }
                return defaultSheetName;
            }
        }
        private static string[] skipActions;
        public static string[] SkipActions
        {
            get
            {
                if (skipActions == null)
                {
                    skipActions = ConfigurationManager.AppSettings["skipActions"].ToLower().Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                }
                return skipActions;
            }
        }

        static void Main(string[] args)
        {
            try
            {
                Utility.AddLog("Start to process, time:" + DateTime.Now.ToString("s"));

                //Get all excel files from current folder
                List<string> files = Utility.GetExcelFiles();

                //Create a new folder 'Output', if exist, skip
                string outputFolder = Path.Combine(Application.StartupPath, OutputDirName);
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                DataTable outputData = null;

                int processFileCount = 0;
                files.ForEach(f =>
                {
                    // skip the files which is under editing or the file is output base file
                    if (f.StartsWith("~$") || f.Equals(OutputBaseFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    ++processFileCount;

                    Utility.AddLog(string.Format("Open {0}", f));
                    //f is the full filepath and name
                    //read file to DataTable
                    DataTable inputMetadata = Utility.ReadInputFile(f);

                    //Translate into human language, store in outputData
                    outputData = Utility.TranslateIntoHumanLanguage(inputMetadata);

                    //Create new excel and insert the data
                    Utility.CreateOutputFiles(outputData, outputFolder);
                    Utility.AddLog(string.Format("Complete {0}\r\n", f));
                });

                Utility.AddLog(string.Format("All done, processed file count: {0}, time: {1}", processFileCount.ToString(), DateTime.Now.ToString("s")));
            }
            catch (Exception ex)
            {
                Utility.AddLog("Found exception:\r\n\r\n" + ex.ToString() + "\r\n");
            }
            finally
            {
                Utility.OutputLogToFile();
            }
        }
    }
}
