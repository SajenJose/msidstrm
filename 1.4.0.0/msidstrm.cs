using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Deployment.WindowsInstaller.Package;


namespace Msidstrm
{
    class msidstrm
    {
        /// <summary>
        /// Method:         dsExtractAllStreamsEngine
        /// Description:    This  method is used to extract all the data streams and also to extract all files based 
        ///                 on the option selected. Interestingly this uses WiX DTF rather than the COM Approach used
        ///                 in the other code file
        /// </summary>
        /// <param name="iFile"></param>
        /// <param name="streamFlag"></param>
        public static void dsExtractAllStreamsEngine(string iFile, int streamFlag)
        {
            InstallPackage winInstaller = new InstallPackage(iFile, DatabaseOpenMode.ReadOnly);
            winInstaller.Message += new InstallPackageMessageHandler(Console.WriteLine);
            string workingDir = System.Environment.GetEnvironmentVariable("TEMP");
            workingDir = (workingDir + "\\" + iFile);
            bool chkAppFolder = System.IO.Directory.Exists(workingDir);

            if (!chkAppFolder)
            {
                System.IO.Directory.CreateDirectory(workingDir);
            }
            

            if (streamFlag == 1)
            {
                try
                {
                    winInstaller.ExportAll(workingDir);
                    Console.Clear();
                    Program.copyrightBanner();
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("The data streams from the provided MSI Database has been extracted.");
                    Console.WriteLine("\nExtracted Streams are saved here: \n" + workingDir);
                    Console.ResetColor();
                }

                catch (System.ArgumentException ex)
                {
                    Console.Clear();
                    Program.copyrightBanner();
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("The data streams from the provided MSI Database has been extracted.");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("\nExtracted Streams are saved here: \n" + workingDir);
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine("\nWarning:");
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine("\nSome data streams within the MSI Database contains Illegal Path Characters" +
                                      "\nin its name and hence may not get extracted.", ex.Message);
                    Console.ResetColor();
                }
                
            }
            
            else if (streamFlag == 0)
            {

                try
                {
                    Console.Clear();
                    Program.copyrightBanner();
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("The input Windows Installer Database is : " + iFile);
                    Console.WriteLine("\nExtracting the files inside the Windows Installer Database, please wait...");
                    winInstaller.WorkingDirectory = workingDir;
                    TextWriter stdOutput = Console.Out;
                    Console.SetOut(TextWriter.Null);
                    winInstaller.ExtractFiles();
                    Console.SetOut(stdOutput);
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("\nExtracted files can be found at : \n" + workingDir);
                    Console.ResetColor();
                }

                catch (System.Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error:\nAn error encountered during the extract process, please check parameters." + ex.Message);
                    Console.ResetColor();
                }                             

            
            }

        }

    }
}
