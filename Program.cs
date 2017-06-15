using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CaseConverter
{
	class Program
	{
		static void Main(string[] args)
		{
            try
            {
                Utility.AddLog("Start to process, time:" + DateTime.Now.ToString("s"));


                //Get all excel files from current folder
                List<string> files = Utility.GetExcelFiles();

                //Create a new folder 'Output', if exist, skip
                string folderName = "Output";
                string outputFolder = Path.Combine(Application.StartupPath, folderName);
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                DataTable outputData = null;

                files.ForEach(f =>
                {
                    Utility.AddLog(string.Format("Open {0}", f));
                    //f is the full filepath and name
                    //read file to DataTable
                    DataTable inputMetadata = Utility.ReadInputFile(f);

                    //Translate into human language, store in outputData
                    outputData = Utility.TranslateIntoHumanLanguage(inputMetadata);

                    //Create new excel and insert the data
                    Utility.CreateOutputFiles(outputData, outputFolder);
                    Utility.AddLog(string.Format("Complete {0}", f));
                });

                Utility.AddLog("All done, time:" + DateTime.Now.ToString("s"));
            }
            catch (Exception ex)
            {
                Utility.AddLog("Found exception:\n\n" + ex.ToString() + "\n");
            }
            finally
            {
                Utility.OutputLogToFile();
            }
		}
	}
}
