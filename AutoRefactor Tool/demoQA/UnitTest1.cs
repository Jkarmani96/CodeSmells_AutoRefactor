using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace demoQA
{
    [TestClass]
    public class UnitTest1
    {
        /*Test Cases for different application*/

        public static string Folder_Path = null;
        public static string Project_Path = null;

        [TestMethod]
        public void CalculatorApplication()
        {
            Folder_Path = @"D:\Maju University\4th Semester\ThesisWorking\Thesis\Project Clone";
            Project_Path = @"\MyCalculator\app\src\main\java\com\DataFlair\";

            CloneProject(@"D:\Maju University\4th Semester\ThesisWorking\Thesis_Applications\DataFlair-Calculator", Folder_Path);

            ConvertTextFileToCSV(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\ResponsesBeforeRefactoring\CalculatorApplicationResponse.txt",
                @"D:\Maju University\4th Semester\ThesisWorking\Thesis\CalculatorApplication.csv");

            getDataFrom_aDoctor(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\CalculatorApplication.csv", Folder_Path, Project_Path);

        }

        [TestMethod]
        public void HeartBeatApplication()
        {
            Folder_Path = @"D:\Maju University\4th Semester\ThesisWorking\Thesis\Project Clone\HeartbeatSampleApp";
            Project_Path = @"\app\src\main\java\com\ooyala\";

            CloneProject(@"D:\Maju University\4th Semester\ThesisWorking\Thesis_Applications\HeartbeatSampleApp", Folder_Path);

            ConvertTextFileToCSV(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\ResponsesBeforeRefactoring\HeartBeatApplicationResponse.txt", 
                @"D:\Maju University\4th Semester\ThesisWorking\Thesis\HeartBeatApplication.csv");

            getDataFrom_aDoctor(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\HeartBeatApplication.csv", Folder_Path, Project_Path);
        }
        
        [TestMethod]
        public void OmnitureSampleApplication()
        {
            Folder_Path = @"D:\Maju University\4th Semester\ThesisWorking\Thesis\Project Clone\OmnitureSampleApp";
            Project_Path = @"\app\src\main\java\com\ooyala\sample\";

            CloneProject(@"D:\Maju University\4th Semester\ThesisWorking\Thesis_Applications\OmnitureSampleApp", Folder_Path);

            ConvertTextFileToCSV(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\ResponsesBeforeRefactoring\OmnitureSampleApplicationResponse.txt", 
                @"D:\Maju University\4th Semester\ThesisWorking\Thesis\OmnitureSampleApplication.csv");

            getDataFrom_aDoctor(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\OmnitureSampleApplication.csv", Folder_Path, Project_Path);
        }

        [TestMethod]
        public void NPAWSampleApplication()
        {
            Folder_Path = @"D:\Maju University\4th Semester\ThesisWorking\Thesis\Project Clone\NPAWSampleApp";
            Project_Path = @"\app\src\main\java\com\ooyala\sample\";

            CloneProject(@"D:\Maju University\4th Semester\ThesisWorking\Thesis_Applications\NPAWSampleApp", Folder_Path);

            ConvertTextFileToCSV(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\ResponsesBeforeRefactoring\NPAWSampleApplicationResponse.txt", 
                @"D:\Maju University\4th Semester\ThesisWorking\Thesis\NPAWSampleApplication.csv");

            getDataFrom_aDoctor(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\NPAWSampleApplication.csv", Folder_Path, Project_Path);
        }

        [TestMethod]
        public void NielsenSampleApplication()
        {
            Folder_Path = @"D:\Maju University\4th Semester\ThesisWorking\Thesis\Project Clone\NielsenSampleApp";
            Project_Path = @"\app\src\main\java\com\ooyala\sample\";

            CloneProject(@"D:\Maju University\4th Semester\ThesisWorking\Thesis_Applications\NielsenSampleApp", Folder_Path);

            ConvertTextFileToCSV(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\ResponsesBeforeRefactoring\NielsenSampleApplicationResponse.txt",
                @"D:\Maju University\4th Semester\ThesisWorking\Thesis\NielsenSampleApplication.csv");

            getDataFrom_aDoctor(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\NielsenSampleApplication.csv", Folder_Path, Project_Path);
        }

        [TestMethod]
        public void AdvancedPlaybackApplication()
        {
            Folder_Path = @"D:\Maju University\4th Semester\ThesisWorking\Thesis\Project Clone\AdvancedPlaybackApp";
            Project_Path = @"\app\src\main\java\com\ooyala\sample\";

            CloneProject(@"D:\Maju University\4th Semester\ThesisWorking\Thesis_Applications\AdvancedPlaybackApp", Folder_Path);

            ConvertTextFileToCSV(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\ResponsesBeforeRefactoring\AdvancedPlaybackApplicationResponse.txt",
                @"D:\Maju University\4th Semester\ThesisWorking\Thesis\AdvancedPlaybackApplication.csv");

            getDataFrom_aDoctor(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\AdvancedPlaybackApplication.csv", Folder_Path, Project_Path);
        }

        [TestMethod]
        public void ChromecastApplication()
        {
            Folder_Path = @"D:\Maju University\4th Semester\ThesisWorking\Thesis\Project Clone\ChromecastApp";
            Project_Path = @"\app\src\main\java\com\ooyala\sample\";

            CloneProject(@"D:\Maju University\4th Semester\ThesisWorking\Thesis_Applications\ChromecastApp", Folder_Path);

            ConvertTextFileToCSV(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\ResponsesBeforeRefactoring\ChromecastApplicationResponse.txt", 
                @"D:\Maju University\4th Semester\ThesisWorking\Thesis\ChromecastApplication.csv");

            getDataFrom_aDoctor(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\ChromecastApplication.csv", Folder_Path, Project_Path);
        }

        [TestMethod]
        public void OptionsApplication()
        {
            Folder_Path = @"D:\Maju University\4th Semester\ThesisWorking\Thesis\Project Clone\OptionsApp";
            Project_Path = @"\app\src\main\java\com\ooyala\sample\";

            CloneProject(@"D:\Maju University\4th Semester\ThesisWorking\Thesis_Applications\OptionsApp", Folder_Path);

            ConvertTextFileToCSV(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\ResponsesBeforeRefactoring\OptionsApplicationResponse.txt", 
                @"D:\Maju University\4th Semester\ThesisWorking\Thesis\OptionsApplication.csv");

            getDataFrom_aDoctor(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\OptionsApplication.csv", Folder_Path, Project_Path);
        }

        [TestMethod]
        public void PulseCheckerApplication()
        {
            Folder_Path = @"D:\Maju University\4th Semester\ThesisWorking\Thesis\Project Clone\PulseCheckerApp";
            Project_Path = @"\app\src\main\java\com\ooyala\sample\";

            CloneProject(@"D:\Maju University\4th Semester\ThesisWorking\Thesis_Applications\PulseCheckerApp", Folder_Path);

            ConvertTextFileToCSV(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\ResponsesBeforeRefactoring\PulseCheckerApplicationResponse.txt",
                @"D:\Maju University\4th Semester\ThesisWorking\Thesis\PulseCheckerApplication.csv");

            getDataFrom_aDoctor(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\PulseCheckerApplication.csv", Folder_Path, Project_Path);
        }

        [TestMethod]
        public void VRSampleApplicationKotlin()
        {
            Folder_Path = @"D:\Maju University\4th Semester\ThesisWorking\Thesis\Project Clone\VRSampleAppKotlin";
            Project_Path = @"\app\src\main\java\com\ooyala\sample\";

            CloneProject(@"D:\Maju University\4th Semester\ThesisWorking\Thesis_Applications\VRSampleAppKotlin", Folder_Path);

            ConvertTextFileToCSV(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\ResponsesBeforeRefactoring\VRSampleApplicationKotlinResponse.txt",
                @"D:\Maju University\4th Semester\ThesisWorking\Thesis\VRSampleApplicationKotlin.csv");

            getDataFrom_aDoctor(@"D:\Maju University\4th Semester\ThesisWorking\Thesis\VRSampleApplicationKotlin.csv", Folder_Path, Project_Path);
        }

        /* Test cases Ends Here*/

       
        /* Below code is Algorithms for Removing the Code Smells from Applications*/

        public static string[] arr = null;
        public static string lastString = null;
        public static string secondLastString = null;

        private static void CloneProject(string sourcePath, string targetPath)
        {
            //Now Create all of the directories
            foreach (string dirPath in Directory.GetDirectories(sourcePath, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(sourcePath, targetPath));
            }

            //Copy all the files & Replaces any files with the same name
            foreach (string newPath in Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories))
            {
                File.Copy(newPath, newPath.Replace(sourcePath, targetPath), true);
            }
        }

        private static void ConvertTextFileToCSV(string SourceFile, string DesignationFile)
        {

            string[] lines, cells;
            StreamWriter csvfile;
            lines = File.ReadAllLines(SourceFile);

            var count = lines.Count();
            // Console.WriteLine(count);

            foreach (string line in lines)
            {
                //Console.WriteLine(line);
            }

            csvfile = new StreamWriter(DesignationFile);
            for (int i = 0; i < lines.Length; i++)
            {
                cells = lines[i].Split(new Char[] { '\t', ';' });
                for (int j = 0; j < cells.Length; j++)
                    csvfile.Write(cells[j] + ",");
                csvfile.WriteLine();
            }
            csvfile.Close();
        }

        public static void getDataFrom_aDoctor(string ResponseFile, string FPath, string PPath)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ResponseFile);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            List<string> MIM_CodeSmell = new List<string>();
            List<string> NLMR_CodeSmell = new List<string>();

            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && xlRange.Cells[i, j].Value2.ToString() == "1")
                    {
                        if (xlRange.Cells[1, j].Value2.ToString() == "MIM")
                            MIM_CodeSmell.Add(xlRange.Cells[i, 1].Value2.ToString());

                        if (xlRange.Cells[1, j].Value2.ToString() == "NLMR")
                            NLMR_CodeSmell.Add(xlRange.Cells[i, 1].Value2.ToString());
                    }
                }
            }

            if (MIM_CodeSmell != null)
            {
                foreach (string p in MIM_CodeSmell)
                {
                    arr = p.Split('.');
                    lastString = arr[arr.Length - 1];
                    secondLastString = arr[arr.Length - 2];

                    Console.WriteLine(secondLastString +" "+  lastString);
                    
                    string File_Path = FPath + PPath + secondLastString + @"\" + lastString + ".java";
                    
                    MemberIgnoringMethod_CodeSmell(File_Path);
                }

            }


            if (NLMR_CodeSmell != null)
            {
                Console.WriteLine( NLMR_CodeSmell.Count);
                    
                foreach (string p in NLMR_CodeSmell)
                {
                    arr = p.Split('.');
                    lastString = arr[arr.Length - 1];
                    secondLastString = arr[arr.Length - 2];

                    Console.WriteLine(secondLastString + " " + lastString);

                    string File_Path = FPath + PPath + secondLastString + @"\" + lastString + ".java";

                    NoLowMemoryResolver_CodeSmell(File_Path);
                }

            }

        
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        public static void MemberIgnoringMethod_CodeSmell(string filePath)
        {
            string text = File.ReadAllText(filePath);
            text = text.Replace("public ", "public static ");
            text = text.Replace("private ", "private static ");
            text = text.Replace("protected ", "protected static ");

            File.WriteAllText(filePath, text);
        }

        public static void NoLowMemoryResolver_CodeSmell(string filePath)
        {
           
            string[] lines = File.ReadAllLines(filePath);

            using (StreamWriter writer = new StreamWriter(filePath))
            {
                foreach(string line in lines)
                {
                    if (line.Contains("super."))
                    {
                        writer.WriteLine(line.Replace(line, line + Environment.NewLine + "\t\tsuper.onLowMemory();"));
                    }
                    else
                    {
                        writer.WriteLine(line);
                    }
                }
            }

        }


    }
}
