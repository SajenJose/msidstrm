using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Deployment.WindowsInstaller.Package;
using Microsoft.Deployment.Compression;
using Microsoft.Deployment.Compression.Cab;


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

        /// <summary>
        /// Method:         dsCabFile
        /// Description:    This method is used to handle various functions like list a cab file contents,
        ///                 create a cab file from a directory containing files, or even extract files from
        ///                 a cab file.
        /// </summary>
        /// <param name="cabFileName"></param>
        /// <param name="cabFileLocation"></param>
        /// <param name="cabFlag"></param>
        public static void dsCabFile(string cabFileName, string cabFileLocation, int cabFlag)
        {
            if (cabFlag == 0)
            {
                Program.copyrightBanner();
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("The cabinet file {0} has the following files embedded within it: ", cabFileName);
                CabInfo CabFile = new CabInfo(cabFileName);
                Console.WriteLine("\nCab FileName" + "\t" + "Cab FileSize (KB)");
                foreach (CabFileInfo CabIndvFile in CabFile.GetFiles())
                {
                    Console.WriteLine(CabIndvFile.Name + "\t" + CabIndvFile.Length);
                }
                Console.ResetColor();
            }

            else if (cabFlag == 1)
            {
                Program.copyrightBanner();
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Cab File {0} is being created. The following files are being added now: \n", cabFileName);
                ArchiveInfo cabFile = CompressedFileType(cabFileName);
                cabFile.Pack(cabFileLocation, true, CompressionLevel.Max, cabProgress);
                Console.WriteLine("\nThe Cab File {0} is successfully created.", cabFileName);
                Console.ResetColor();
            }

            else if (cabFlag == 2)
            {
                Program.copyrightBanner();
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Cab File {0} is being extracted. The following files are being extracted now: \n", cabFileName);
                ArchiveInfo cabFile = CompressedFileType(cabFileName);
                cabFile.Unpack(cabFileLocation, cabProgress);
                Console.WriteLine("\nThe cab file {0} is extracted to {1}", cabFileName, cabFileLocation);
                Console.ResetColor();
            }
            
        }

        
        /// <summary>
        /// Method:         CompressedFileType
        /// Description:    This method is used to return the correct cab file format for the string that was taken as an input
        ///                 in the argument. This is related to the CabFile Method.
        /// </summary>
        /// <param name="cmpFileType"></param>
        /// <returns></returns>
        public static ArchiveInfo CompressedFileType(string cmpFileType)
        {
            string fileExtension = Path.GetExtension(cmpFileType).ToUpperInvariant();

            if (fileExtension == ".CAB")
            {
                return new Microsoft.Deployment.Compression.Cab.CabInfo(cmpFileType);
            }

            else 
            {
                throw new ArgumentException("Error: Unknown Cab File Extension: " + fileExtension);
            }

        }

        
        /// <summary>
        /// Method:             cabProgress
        /// Description:        This method lists the files that are currently being worked on during compression or decompression
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        public static void cabProgress(object source, ArchiveProgressEventArgs e)
        {
            if (e.ProgressType == ArchiveProgressType.StartFile)
            {                
                Console.WriteLine(e.CurrentFileName + " - file is being processed");                
            }
        }

    }
}
