using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using WindowsInstaller;
using System.Reflection; //namespace Assembly
using System.Diagnostics; //namespace FileVersionInfo
using System.Management; //namespace WMI tasks


namespace Msidstrm
{
    /// <summary>
    /// Program Name:   MSI Data Stream Reader
    ///
    /// Description:    This program would open the msi file that has been provided
    ///                 and read off the various data streams that are embedded within
    ///                 the MSI file. This program was written as a part of understanding
    ///                 which streams gets stripped off during caching of MSI files. Also
    ///                 added the logic to add new data streams and also to delete existing
    ///                 data streams
    ///                
    /// Logic Used:     Read / Add / Delete / Extract the values from the temporary table _Stream
    /// 
    /// Version History:
    /// 
    ///                     1.0.0 -:    First release includes the ability to read data streams
    ///                     1.1.0 -:    Included the ability to add and delete data streams, backup MSI files and also
    ///                                 provide an option for disabling backup files    
    ///                     1.2.0 -:    Included basic Error handling
    ///                     1.3.0 -:    Fixed the error in ProductVersion display and  also implemented a different approach
    ///                                 of dynamically reading the product version rather than hard coding it. Also fixed 
    ///                                 a few exception errors that is thrown if an incorrect syntax is provided (for e.g.
    ///                                 msidstrm.exe /anb msidstrm.exe - here  we have missed msi file name, which resulted
    ///                                 in an exception error. This has been taken care of.
    ///                     1.4.0 -:    Introduced two new options / features - eas / eaf, to extract data streams and extract 
    ///                                 files embedded inside cabs
    ///                     1.5.0 -:    Introduced two new features, one to read the current system's windows installer version.
    ///                                 The  second option allows to convert GUID to a compressed / packed / squished format
    ///                                 and vice versa.
    ///                     1.6.0 -:    Introduced a new feature to handle cab files, list cab files, create a cab file, extract
    ///                                 a cab file
    ///                     1.7.0 -:    Introduced a new feature to list / save the files & folder path size / length from a given
    ///                                 input path location recursively. 
    ///                     1.8.0 -:    Introduced a new feature to get CPU Architecture.
    /// </summary>
    class Program
    {
        public static int disableBackup = 0; //determines whether msi file needs to be backed up
        public static string productVersion;
        public static int extractStream = 0; //determines whether it is a extract stream or extract file
        public static int cabaction = 0;

        /// <summary>
        /// Method:         Main
        /// Description:    This is the main engine of the utility
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
           
            //Obtain the file version and display it as the product version.
            Assembly msidstrmAssembly = Assembly.GetExecutingAssembly();
            FileVersionInfo msidstrmFileVersion = FileVersionInfo.GetVersionInfo(msidstrmAssembly.Location);
            productVersion = msidstrmFileVersion.FileVersion;

            string winInstallerPackagedll = "Msidstrm.Microsoft.Deployment.WindowsInstaller.Package.dll";
            string winInstallerdll = "Msidstrm.Microsoft.Deployment.WindowsInstaller.dll";
            string winInstallerCompressiondll = "Msidstrm.Microsoft.Deployment.Compression.dll";
            string winInstallerCompressionCabdll = "Msidstrm.Microsoft.Deployment.Compression.Cab.dll";

            EmbeddedAssembly.Load(winInstallerPackagedll, "Microsoft.Deployment.WindowsInstaller.Package.dll");
            EmbeddedAssembly.Load(winInstallerdll, "Microsoft.Deployment.WindowsInstaller.dll");
            EmbeddedAssembly.Load(winInstallerCompressiondll, "Microsoft.Deployment.Compression.dll");
            EmbeddedAssembly.Load(winInstallerCompressionCabdll, "Microsoft.Deployment.Compression.Cab.dll");

            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(DomainResolve);

            if (args.Length == 0)
            {
                copyrightBanner();
                HelpMsg();
                Console.ReadKey();
            }

            else if (args[0] == "/?" || args[0] == "-h")
            {
                copyrightBanner();
                HelpMsg();
            }

            else if (args[0] == "/a" || args[0] == "/anb")
            {
                try 
                {
                    string inputFile = args[1];
                    string inputDSFile = args[2];

                    if (args[0] == "/anb")
                    {
                        disableBackup = 1;
                    }

                    dsAddEngine(inputFile, inputDSFile);
                }
                
                catch (System.IndexOutOfRangeException ex) 
                {
                    copyrightBanner();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: There was an error opening the file, please check input parameters.", ex.Message);
                    Console.ResetColor(); 
                }

            }

            else if (args[0] == "/d" || args[0] == "/dnb")
            {
                try 
                {
                    string inputFile = args[1];
                    string inputDSName = args[2];

                    if (args[0] == "/dnb")
                    {
                        disableBackup = 1;
                    }

                    dsDeleteEngine(inputFile, inputDSName);
                }

                catch (System.IndexOutOfRangeException ex)
                {
                    copyrightBanner();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: There was an error opening the file, please check input parameters.", ex.Message);
                    Console.ResetColor(); 
                }

            }

            else if (args[0] == "/eas" || args[0] == "/eaf")
            {
                try
                {
                    string inputFile = args[1];
                    if (args[0] == "/eas")
                    {
                        extractStream = 1;
                    }

                    msidstrm.dsExtractAllStreamsEngine(inputFile, extractStream);
                }

                catch (System.IndexOutOfRangeException ex)
                {
                    copyrightBanner();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: There was an error opening the file, please check input parameters.", ex.Message);
                    Console.ResetColor();
                }

            }

            else if (args[0] == "/guid")
            {

                copyrightBanner();
                int errorFlag = 0; //flag to check syntax error

                if (args.Length < 3) //Syntax arguments correct, it should be 3 /guid N/B/P/D {guid} if not correct, it will display guid help
                {
                    errorFlag = 1;
                    string errArg0 = String.Empty;
                    Guid errArg1 = new Guid();
                    Program.GUIDConverter(errArg0, errArg1, errorFlag);
                }
                
                else if (args.Length == 3) //Number of arguments correct, 
                {

                    if (args[1] == "N" || args[1] == "D" || args[1] == "B" || args[1] == "P") //Check if output format is correct
                    {
                        try //Check if GUID entered is in the correct format
                        {
                            string convertOption = args[1];
                            Guid inputGUID = new Guid(args[2]);
                            Program.GUIDConverter(convertOption, inputGUID, errorFlag);
                        }

                        catch (System.FormatException) //If guid not correct display guid help
                        {
                            errorFlag = 1;
                            string errArg0 = String.Empty;
                            Guid errArg1 = new Guid();
                            Program.GUIDConverter(errArg0, errArg1, errorFlag);
                            return;
                            
                        }
                    }

                    else //hits here if there is an error in the syntax
                    {
                        errorFlag = 1;
                        string errArg0 = String.Empty;
                        Guid errArg1 = new Guid();
                        Program.GUIDConverter(errArg0, errArg1, errorFlag);
                        return;
                    }
                }
            }

            else if (args[0] == "/ccab" || args[0] == "/lcab" || args[0] == "/ecab")//handle cab files
            {
                try 
                {
                    if (args[0] == "/ccab")//create cab file
                    {
                        string cbFileName = args[1];
                        string cbLocation = args[2];
                        cabaction = 1;
                        msidstrm.dsCabFile(cbFileName, cbLocation, cabaction);
                    }

                    else if (args[0] == "/ecab")//extract cab file
                    {
                        string cbFileName = args[1];
                        string cbLocation = args[2];
                        cabaction = 2;
                        
                        msidstrm.dsCabFile(cbFileName, cbLocation, cabaction);
                    }

                    else
                    {
                        string cbFileName = args[1];
                        string cbLocation = args[1];
                        cabaction = 0;
                        msidstrm.dsCabFile(cbFileName, cbLocation, cabaction);
                    }

                }

                catch 
                {
                    Program.copyrightBanner();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: Please find the correct syntax below");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("\nSyntax: \n" +
                                        "\nList Cab Files:      -> msidstrm /lcab <cabFileName>" +
                                        "\nCreate Cab File:     -> msidstrm /ccab <cabFileName> <Folder Containing Files>" +
                                        "\nExtract Cab File:    -> msidstrm /ecab <cabFileName> <Folder to Extract>");
                    Console.ResetColor();
                }

                
            }

            else if (args[0] == "/gfp")
            {
                try 
                {
                    string[] filePathArray = args;
                    getFilePathClassLength.getFilePathLen(filePathArray);                    
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

            else if (args[0] == "/cpu")
            {
                try 
                {
                    getcpudetails.getCPUDetails();
                    Console.ResetColor();
                }
                
                catch 
                {
                    Program.copyrightBanner();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: Please find the correct syntax below");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("\nSyntax: \n" +
                                        "\nmsidstrm.exe /cpu");
                    Console.ResetColor();
                }
            }
            else
            {
                string inputFile = args[0];
                dsReadEngine(inputFile);
                Console.ResetColor();
            }
        }

        /// <summary>
        /// Method:         HelpMsg
        /// Description:    This method displays the generic usage of the tool.
        /// </summary>
        static void HelpMsg()
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Syntax: \n     -: msidstrm.exe [options] <msifilename> [DataStream Name]\n");
            Console.WriteLine("Usage & Examples: \n"+
                              "     -: Read DataStreams: msidstrm.exe <msifilename>.msi\n"+
                              "     -: Update DataStreams: msidstrm.exe /a <msifilename>.msi <newfile>.ext\n"+
                              "     -: Delete DataStreams: msidstrm.exe /d <msifilename>.msi <datastreamname>\n"+
                              "     -: Extract DataStreams: msidstrm.exe /eas <msifilename>.msi");
            Console.WriteLine("\nOptions:" +
                              "\n     -: /a     -> Add a data stream to the msi file" +
                              "\n     -: /anb   -> Add a data stream without creating backup of msi file" +
                              "\n     -: /d     -> delete the data stream from the msi file" +
                              "\n     -: /dnb   -> Delete the data stream without creating backup of msi file" +
                              "\n     -: /eas   -> Extract all streams and tables present inside the msi file" +
                              "\n     -: /eaf   -> Extract all the files from embedded cab inside the msi file" + 
                              "\n     -: /?     -> Displays this help message");

            Console.Write("Press any key to see advanced options...");
            Console.ReadKey();
            copyrightBanner();
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Advanced Options:" +
                              "\n     -: /guid  -> Converts a GUID to Packed or Compressed format and vice versa" +
                              "\n     -:        -> Try following command to get detailed help: 'msidstrm /guid'"  +
                              "\n     -: /lcab  -> Lists the contents of a cab File" +
                              "\n     -: /ccab  -> Creates a cab File from a directory containing files" +
                              "\n     -: /ecab  -> Extracts the contents of a cab file" +
                              "\n     -: /gfp   -> Lists the FilePath Lengths for files in input Path Location." +
                              "\n     -:        -> example: msidstrm /gfp \"<path/directory>\"" +
                              "\n     -: /cpu   -> Gets the CPU architecture details.");     
            Console.ResetColor();
        }
                
        /// <summary>
        /// Method:         dsReadEngine
        /// Description:    This method is used to read the datastreams from the msi file.
        /// </summary>
        /// <param name="iFile"></param>
        
        static void dsReadEngine(string iFile)
        {
            try 
            {
                copyrightBanner();
                string dbRecords = String.Empty; //used to read each record value
                int dsCounter = 01; //count of datastreams within the MSI just used for cosmetic display

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("The Windows Installer Database provided is: " + iFile);
                Console.WriteLine("\nReading Data Streams, please wait...");
                Console.WriteLine("\nThe following Data Streams are present in " + iFile + "\n");

                //Get the type of Windows Installer 
                Type winInstallerType = Type.GetTypeFromProgID("WindowsInstaller.Installer");

                //Create the Windows Installer Object
                Installer winInstallerObj = (Installer)Activator.CreateInstance(winInstallerType);

                //Open the Windows Installer Database File
                Database winInstallerDB = winInstallerObj.OpenDatabase(iFile, MsiOpenDatabaseMode.msiOpenDatabaseModeReadOnly);

                //Create a View and Execute it
                View winInstallerView = winInstallerDB.OpenView("SELECT `Name`, `Data` FROM _Streams");
                winInstallerView.Execute(null);


                while (true)
                {
                    //Obtain the records from the created view
                    Record winInstallerRecord = winInstallerView.Fetch();

                    if (winInstallerRecord != null)
                    {
                        dbRecords = winInstallerRecord.get_StringData(1);
                        Console.WriteLine("Datastream-" + dsCounter + " : " + dbRecords);
                        dsCounter++;
                    }
                        
                    else
                    {
                        return;
                    }                    
                }

                
            }
            
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: There was an error opening the file: " + iFile, ex.Message);
                Console.ResetColor();
            }      

        }

        /// <summary>
        /// Method:         dsAddEngine
        /// Description:    This method is used to update a new file as a datastream into the msi file
        /// </summary>
        /// <param name="iFile"></param>
        /// <param name="iDSFile"></param>
        static void dsAddEngine(string iFile, string iDSFile)
        {
            try
            {
                copyrightBanner();
                string dbRecords = String.Empty; //variable to hold the record values being read
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("-: The Windows Installer Database provided is: " + iFile);
                Console.WriteLine("-: The Data Stream File to embedd is: " + iDSFile);

                if (disableBackup == 0)
                {
                    Console.WriteLine("-: Back up Windows Installer Database file : " + iFile + " to " + iFile + ".bac");
                    Random rSeed = new Random();
                    int rSeedValue = rSeed.Next(0, 100);
                    File.Copy(iFile, iFile + rSeedValue + ".bac");
                    Console.WriteLine("-: Backup Completed...");
                }


                Console.WriteLine("-: Processing Data Streams, please wait...");
                Console.WriteLine("-: Opening the Windows Installer Database File...\n");

                //Get the Type for Windows Installer
                Type winInstallerType = Type.GetTypeFromProgID("WindowsInstaller.Installer");

                //Create the windows installer object 
                Installer winInstallerObj = (Installer)Activator.CreateInstance(winInstallerType);

                //Open the Database
                Database winInstallerDB = winInstallerObj.OpenDatabase(iFile, MsiOpenDatabaseMode.msiOpenDatabaseModeTransact);

                try
                {
                    //Create a View and Execute it
                    View winInstallerView = winInstallerDB.OpenView("SELECT `Name`, `Data` FROM _Streams");
                    Record winInstallerRecord = winInstallerObj.CreateRecord(2);
                    winInstallerRecord.set_StringData(1, iDSFile);
                    winInstallerView.Execute(winInstallerRecord);
                    winInstallerRecord.SetStream(2, iDSFile);
                    winInstallerView.Modify(MsiViewModify.msiViewModifyAssign, winInstallerRecord);
                    winInstallerDB.Commit();
                    Console.WriteLine("-: The file " + iDSFile + " is now embedded into the MSI Data Streams.");
                }

                catch (System.Runtime.InteropServices.COMException ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: There was an error opening the file : " + iDSFile, ex.Message);
                    Console.ResetColor();
                }
            }

            catch (System.IndexOutOfRangeException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: There was an error opening the file: " + iFile, ex.Message);
                Console.ResetColor(); 
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: There was an error opening the file: " + iFile, ex.Message);
                Console.ResetColor();
            }                       
        }

        /// <summary>
        /// Method:         dsDeleteEngine
        /// Description:    This method is used to delete a datastream from the MSI File
        /// </summary>
        /// <param name="iFile"></param>
        /// <param name="iDSName"></param>
        static void dsDeleteEngine(string iFile, string iDSName)
        {
            try
            {
                copyrightBanner();

                string dbRecords = String.Empty; //Variable to hold the values being read using the SQL query

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("-: The Windows Installer Database provided is: " + iFile);
                Console.WriteLine("-: The Data Stream to delete is : " + iDSName);

                if (disableBackup == 0)
                {
                    Console.WriteLine("-: Back up Windows Installer Database file : " + iFile + " to " + iFile + ".bac");
                    Random rSeed = new Random();
                    int rSeedValue = rSeed.Next(0, 100);
                    File.Copy(iFile, iFile + rSeedValue + ".bac");
                    Console.WriteLine("-: Backup Completed...");
                }

                Console.WriteLine("-: Processing Data Streams, please wait...");
                Console.WriteLine("-: Opening the Windows Installer Database File...");

                //Get the Windows Installer Type
                Type winInstallerType = Type.GetTypeFromProgID("WindowsInstaller.Installer");

                //Create the Windows Installer Object
                Installer winInstallerObj = (Installer)Activator.CreateInstance(winInstallerType);

                //Open the Database
                Database winInstallerDB = winInstallerObj.OpenDatabase(iFile, MsiOpenDatabaseMode.msiOpenDatabaseModeTransact);

                //Create a View & Execute it
                View winInstallerView = winInstallerDB.OpenView("SELECT `Name` FROM _Streams WHERE `Name`=?");
                Record winInstallerRecord = winInstallerObj.CreateRecord(2);
                winInstallerRecord.set_StringData(1, iDSName);
                winInstallerView.Execute(winInstallerRecord);
                winInstallerRecord = winInstallerView.Fetch();
                if (winInstallerRecord == null)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("\n-: Error: The data stream " + iDSName + " is not present.");
                    Console.ResetColor();
                    return;
                }
                winInstallerView.Modify(MsiViewModify.msiViewModifyDelete, winInstallerRecord);
                winInstallerDB.Commit();
                Console.WriteLine("-: The data stream " + iDSName + " has now been deleted from " + iFile);
            }

            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\nError: There was an error opening the file : " + iFile, ex.Message);
                Console.ResetColor();
            }
        }
        
        /// <summary>
        /// Method:         GUIDConverter
        /// Description:    This method is used to convert from a normal GUID to a Packed or Compressed GUID
        ///                 The method also handles the vice versa conversion as well.
        /// </summary>
        public static void GUIDConverter(string outOption, Guid inpGUID, int errFlg)
        {
            
            //Error Handling
            if (errFlg == 1)
            {
                //The "N option of  the GUid to string converts removes { and - refer Guid.ToString Method - http://msdn.microsoft.com/en-us/library/97af8hh4.aspx            
                //N -> 32 digits: 00000000000000000000000000000000
                //D -> 32 digits separated by hyphens: 00000000-0000-0000-0000-000000000000
                //B -> 32 digits separated by hyphens, enclosed in braces:{00000000-0000-0000-0000-000000000000}
                //P -> 32 digits separated by hyphens, enclosed in parentheses: (00000000-0000-0000-0000-000000000000)

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Syntax: msidstrm /guid [N/D/B/P] <GUID>\n" +
                                  "\n       N -> Format: 00000000000000000000000000000000" +
                                  "\n       D -> Format: 00000000-0000-0000-0000-000000000000" +
                                  "\n       B -> Format: {00000000-0000-0000-0000-000000000000}" +
                                  "\n       P -> Format: (00000000-0000-0000-0000-000000000000)\n" +
                                  "\n       e.g. msidstrm /guid N {12345678-1234-1234-1234-123456789ABC}");
                Console.ResetColor();
                return;
            }

            else if (errFlg == 0)
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("The Input GUID  is: " + inpGUID.ToString(outOption));
                string inputGuidString = inpGUID.ToString("N"); //convert to just plain string without hypen and brackets
                char[] GuidArray = inputGuidString.ToArray(); //Convert it to an array

                //Now we need to reverse the guid characters to the format 8-4-4-16, the 8-4-4 are just plain reverses, however 16 is individual byte reversing

                int[] reverseCounter = new int[] { 8, 4, 4, 2, 2, 2, 2, 2, 2, 2, 2 };

                int startPosition = 0;

                for (int i = 0; i < reverseCounter.Length; i++)
                {
                    Array.Reverse(GuidArray, startPosition, reverseCounter[i]);
                    startPosition += reverseCounter[i];
                }

                string outputStr = new string(GuidArray);
                Guid outputGUID = new Guid(outputStr); //Conversioninto GUID is necessary to apply the Format specified during command line.
                Console.WriteLine("The Output GUID is: " + outputGUID.ToString(outOption).ToUpper());
                Console.ResetColor();
            }
        }
        
        /// <summary>
        /// Method:         copyrightBanner
        /// Description:    This method is  used to display the  banner for the utility.
        /// </summary>
        public static void copyrightBanner()
        {
            //Obtain Windows Installer File Version
            FileVersionInfo msiVersion;
            string msiDllFileName = Path.Combine(Environment.SystemDirectory, "msi.dll");
            try
            {
                msiVersion = FileVersionInfo .GetVersionInfo(msiDllFileName);
                
            }

            catch (FileNotFoundException) 
            {
                return;
            }

            //Copyright Message
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("MSI Data Stream Utility Version " + productVersion);
            Console.WriteLine("Description: This utility lists / modifies the data streams of a MSI database");
            Console.WriteLine("Windows Installer Version: " + msiVersion.ProductVersion);
            Console.WriteLine("Author: Sajen Jose | Suggestions & Feedback: theKludgeBin@gmail.com");
            Console.WriteLine("--------------------------------------------------------------------------------\n");
            Console.ResetColor();
        }

        /// <summary>
        /// Method:         DomainResolve
        /// Description:    This method is used to embedd the dll files into the executable code            
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        public static Assembly DomainResolve(object sender, ResolveEventArgs args)
        {
            return EmbeddedAssembly.Get(args.Name);
        }
    }
}
