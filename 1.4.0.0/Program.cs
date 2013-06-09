using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using WindowsInstaller;
using System.Reflection; //namespace Assembly
using System.Diagnostics; //namespace FileVersionInfo


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
    /// Logic Used:     Read / Add / Delete the values from the temporary table _Stream
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
    /// </summary>
    class Program
    {
        public static int disableBackup = 0;
        public static string productVersion;
        public static int extractStream = 0;

        static void Main(string[] args)
        {
           
            //Obtain the file version and display it as the product version.
            Assembly msidstrmAssembly = Assembly.GetExecutingAssembly();
            FileVersionInfo msidstrmFileVersion = FileVersionInfo.GetVersionInfo(msidstrmAssembly.Location);
            productVersion = msidstrmFileVersion.FileVersion;

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
                    //string inputDSName = args[2];
                    //dsExtractEngine(inputFile, inputDSName);
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

            else
            {
                string inputFile = args[0];
                dsReadEngine(inputFile);
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
        /// Method:         dsExtractEngine
        /// Description:    This method iis used to extract a datastream from the MSI.
        /// </summary>
        /// <param name="iFile"></param>
        /// <param name="iDSName"></param>
        static void dsExtractEngine(string iFile, string iDSName)
        {           
               
            copyrightBanner();
            string dbRecords = string.Empty; //holds the values during the SQL query

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("-: The Windows Installer Database provided is: " + iFile);
            Console.WriteLine("-: The Data Stream to Extract is : " + iDSName);
            Console.WriteLine("-: Processing Data Streams, please wait...");
            Console.WriteLine("-: Opening the Windows Installer Database File...");

            //Windows Installer Type Initialization
            Type winInstallerType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
                
            //Create Installer Object
            Installer winInstallerObj = (Installer)Activator.CreateInstance(winInstallerType);
                
            //Open the MSI Database
            Database winInstallerDB = winInstallerObj.OpenDatabase(iFile, MsiOpenDatabaseMode.msiOpenDatabaseModeTransact);
                
            //Create a View and Execute it
            View winInstallerView = winInstallerDB.OpenView("Select `Data` FROM `_Streams` WHERE `Name`="+ "'" + iDSName + "'");
            Record winInstallerRecord = winInstallerObj.CreateRecord(2);
            winInstallerRecord.set_StringData(1, iDSName);
            winInstallerView.Execute(winInstallerRecord);
            winInstallerRecord = winInstallerView.Fetch();
            int dataSize = winInstallerRecord.get_DataSize(1);            
            string dataSet = winInstallerRecord.ReadStream(1, dataSize, MsiReadStream.msiReadStreamBytes);    
            System.IO.File.WriteAllText(iDSName, dataSet);
//            InstallPackage winInstaller = new InstallPackage(iFile);
//            winInstaller.Message += new InstallPackageMessageHandler(Console.WriteLine);
            
            

        }


        public static void copyrightBanner()
        {
            //Copyright Message
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("MSI Data Stream Utility Version " + productVersion);
            Console.WriteLine("Description: This utility lists / modifies the data streams of a MSI database");
            Console.WriteLine("Author: Sajen Jose | Suggestions & Feedback: theKludgeBin@gmail.com");
            Console.WriteLine("--------------------------------------------------------------------------------\n");
            Console.ResetColor();
        }
    }
}
