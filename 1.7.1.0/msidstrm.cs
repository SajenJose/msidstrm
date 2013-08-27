using System;
using System.IO;
using System.Collections;
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


    /// <summary>
    /// Class:                  Get File Path Length
    /// Description:            This class contains methods used to obtain the File & Directory Path Lengths useful in those
    ///                         rare scenarios of Application Packaging.
    /// </summary>
    class getFilePathClassLength
    {
        
        /// <summary>
        /// Method:             getFilePathLen
        /// Description:        The core engine which reads the path from the input argument does a recursive search and shows the file
        ///                     path lengths as well as directory path lengths.
        /// </summary>
        /// <param name="filePath"></param>
        public static void getFilePathLen(string[] filePath)
        {
            
            if(File.Exists("msidstrm_GetFilePath.csv"))
            {
                File.Delete("msidstrm_GetFilePath.csv");
            }

            try
            {
                foreach (string iPath in filePath)
                {
                    if (File.Exists(iPath))
                    {
                        Program.copyrightBanner();
                        fileEngine(iPath);
                        getFilePathClassLength.statusComplete();
                    }

                    else if (Directory.Exists(iPath))
                    {
                        Program.copyrightBanner();
                        dirEngine(iPath);
                        getFilePathClassLength.statusComplete();
                    }

                    else
                    {
                        Program.copyrightBanner();
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Error: The provided input {0} is not a valid file or directory.", iPath);
                        Console.ResetColor();                         
                    }

                }


                string filesignature = "\nThis file was created using the MSI Data Stream Utility. Please contact theKludgeBin@gmail.com for any feedback or suggestions.";
                fileCsvWriter(filesignature);


            }

            catch
            {
                Program.copyrightBanner();
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: Please find the correct syntax below");
                Console.ForegroundColor = ConsoleColor.White;

                Console.WriteLine("\nSyntax: \n" +
                                    "\nmsidstrm.exe /gfp \"<Path/Directory>\"");
                Console.ResetColor();
            }

        }

        /// <summary>
        /// Method:             fileEngine
        /// Description:        This method creates an entry into the csv file for each recursive file.    
        /// </summary>
        /// <param name="iPath"></param>
        public static void fileEngine(string iPath)
        {

            string dataEntryLine = "Processing File #," + iPath + "," + iPath.Length + "\n";
            fileCsvWriter(dataEntryLine);
        }


        /// <summary>
        /// Method:             dirEngine
        /// Description:        This method is the crux of recursive search for each sub directory and then writes the directory length as well.
        /// </summary>
        /// <param name="iDir"></param>
        public static void dirEngine(string iDir)
        {
            string[] fEntries = Directory.GetFiles(iDir);

            foreach (string fileEntry in fEntries)
            {
                fileEngine(fileEntry);
            }

            string[] dEntries = Directory.GetDirectories(iDir);

            foreach (string dirEntry in dEntries)
            {
                string dataEntrylinedir = "Processing Directory #," + dirEntry + "," + dirEntry.Length + "\n";
                fileCsvWriter(dataEntrylinedir);
                dirEngine(dirEntry);
            }
        }


        public static void statusComplete()
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Status: Completed\n");
            Console.WriteLine("The output is saved in a file msidstrm_GetFilePath.csv present in the \n" +
                              "directory from where this Utility was run.");
            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.WriteLine("\nNote:\nThe msidstrm_GetFilePath.csv gets overwritten everytime this utility is run.");
            Console.ResetColor();

        }
        
        /// <summary>
        /// Method:             fileCsvWriter
        /// Description:        This method creates a file with the name msidstrm_GetFilePath.csv in the same location from where
        ///                     the Msidstrm utility was run and then also appends each file and directory path lenth entry.
        /// </summary>
        /// <param name="dataEntry"></param>
        public static void fileCsvWriter(string dataEntry)
        {
            string FileName = "msidstrm_GetFilePath.csv";


            if (!File.Exists(FileName))
            {                
                string fileHeader = "Action, File-Directory Path, File-Directory Path Length" + "\n";
                File.WriteAllText(FileName, fileHeader);
            }

            File.AppendAllText(FileName, dataEntry);
        }
    }

}
