/*
* This program consists of the following stages:
* 1. Verification of input
*    1.1. Verify that the user has supplied visual studio (VS) solution file
*    1.2. Verify that the solution file exists
* 2. Load TwinCAT project
*    2.1. Find TwinCAT project in VS solution file
*    2.2. Find which version of TwinCAT was used
* 3. Load the VS DTE and TwinCAT XAE with the right version of TwinCAT using the remote manager
*    The "right" version of TwinCAT is decided according to these rules:
*    - If TwinCAT project version is forced (by -w argument to TcUnit-Runner), go with this version, otherwise...
*    - If TwinCAT project is pinned, go with this version, otherwise...
*    - Go with latest installed version of TwinCAT
* 4. Load the solution
* 5. Check that the solution has at least one PLC-project
* 6. Clean the solution
* 7. Build the solution. Make sure that build was successful.
* 8. Set target NetId to 127.0.0.1.1.1
* 9. If user has provided 'TcUnitTaskName', iterate all PLC projects and do:
*     9.1. Find the 'TcUnitTaskName', and set the <AutoStart> to TRUE and <Disabled> to FALSE for the TIRT^ of the TASK
*     9.2. Iterate the rest of the tasks (if there are any), and set the <AutoStart> to FALSE and <Disabled> to TRUE for the TIRT^ of the task
* 10. Enable boot project autostart for all PLC projects
* 11. Activate configuration
* 12. Restart TwinCAT
* 13. Wait until TcUnit has reported all results and collect all results
* 14. Write all results to xUnit compatible XML-file
*/

using EnvDTE80;
using log4net;
using NDesk.Options;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         
using TCatSysManagerLib;
using TwinCAT.Ads;
using TwinCAT;

namespace TcUnit.TcUnit_Runner
{
    class Program
    {
        private static string VisualStudioSolutionFilePath = null;
        private static string TwinCATProjectFilePath = null;
        private static string TcUnitTaskName = null;
        private static string ForceToThisTwinCATVersion = null;
        private static string AmsNetId = null;
        private static List<int> AmsPorts = new List<int>();
        private static string Timeout = null;
        private static string PLCProjectName = null;
        private static string LibrarySavePath = null;
        private static VisualStudioInstance vsInstance;
        private static ILog log = LogManager.GetLogger("TcUnit-Runner");

        [STAThread]
        static void Main(string[] args)
        {
            //Thread.CurrentThread.Priority = ThreadPriority.Highest;
            //System.Diagnostics.Process current = System.Diagnostics.Process.GetCurrentProcess();
            //current.PriorityClass = System.Diagnostics.ProcessPriorityClass.AboveNormal;


            bool showHelp = false;
            bool enableDebugLoggingLevel = false;
            Console.CancelKeyPress += new ConsoleCancelEventHandler(CancelKeyPressHandler);
            log4net.GlobalContext.Properties["LogLocation"] = AppDomain.CurrentDomain.BaseDirectory + "\\logs";
            log4net.Config.XmlConfigurator.ConfigureAndWatch(new System.IO.FileInfo(AppDomain.CurrentDomain.BaseDirectory + "log4net.config"));

            OptionSet options = new OptionSet()
                .Add("v=|VisualStudioSolutionFilePath=", "The full path to the TwinCAT project (sln-file)", v => VisualStudioSolutionFilePath = v)
                .Add("w=|TcVersion=", "[OPTIONAL] The TwinCAT version to be used to load the TwinCAT project", w => ForceToThisTwinCATVersion = w)
                .Add("n=|PLCProjectName=", "The full name of the PLC project, eg 'NameOfProject^NameOfProject Project'", n => PLCProjectName = n)
                .Add("l=|LibraryPath=", "The full path of the library file including name and file extention (.library or .compiled-library)", l => LibrarySavePath = l)
                .Add("d|debug", "[OPTIONAL] Increase debug message verbosity", d => enableDebugLoggingLevel = d != null)
                .Add("?|h|help", h => showHelp = h != null);
            try
            {
                options.Parse(args);

            }
            catch (OptionException e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `TcUnit-Runner --help' for more information.");
                Environment.Exit(Constants.RETURN_ARGUMENT_ERROR);
            }


            if (showHelp)
            {
                DisplayHelp(options);
                Environment.Exit(Constants.RETURN_SUCCESSFULL);
            }

            /* Set logging level.
             * This is handled by changing the log4net.config file on the fly.
             * The following levels are defined in order of increasing priority:
             * - ALL
             * - DEBUG
             * - INFO
             * - WARN
             * - ERROR
             * - FATAL
             * - OFF
            */
            XmlDocument doc = new XmlDocument();
            doc.Load(AppDomain.CurrentDomain.BaseDirectory + "log4net.config");
            XmlNode root = doc.DocumentElement;
            XmlNode subNode1 = root.SelectSingleNode("root");
            XmlNode nodeForModify = subNode1.SelectSingleNode("level");
            if (enableDebugLoggingLevel)
                nodeForModify.Attributes[0].Value = "DEBUG";
            else
                nodeForModify.Attributes[0].Value = "INFO";
            doc.Save(AppDomain.CurrentDomain.BaseDirectory + "log4net.config");
            System.Threading.Thread.Sleep(500); // A tiny sleep just to make sure that log4net manages to detect the change in the file

            /* Make sure the user has supplied the path for the Visual Studio solution file.
             * Also verify that this file exists.
             */
            if (VisualStudioSolutionFilePath == null)
            {
                log.Error("Visual studio solution path not provided!");
                Environment.Exit(Constants.RETURN_VISUAL_STUDIO_SOLUTION_PATH_NOT_PROVIDED);
            }

            if (!File.Exists(VisualStudioSolutionFilePath))
            {
                log.Error("Visual studio solution " + VisualStudioSolutionFilePath + " does not exist!");
                Environment.Exit(Constants.RETURN_VISUAL_STUDIO_SOLUTION_PATH_NOT_FOUND);
            }

            LogBasicInfo();


            MessageFilter.Register();

            TwinCATProjectFilePath = TcFileUtilities.FindTwinCATProjectFile(VisualStudioSolutionFilePath);
            if (String.IsNullOrEmpty(TwinCATProjectFilePath))
            {
                log.Error("Did not find TwinCAT project file in solution. Is this a TwinCAT project?");
                Environment.Exit(Constants.RETURN_TWINCAT_PROJECT_FILE_NOT_FOUND);
            }

            if (!File.Exists(TwinCATProjectFilePath))
            {
                log.Error("TwinCAT project file " + TwinCATProjectFilePath + " does not exist!");
                Environment.Exit(Constants.RETURN_TWINCAT_PROJECT_FILE_NOT_FOUND);
            }

            string tcVersion = TcFileUtilities.GetTcVersion(TwinCATProjectFilePath);

            if (String.IsNullOrEmpty(tcVersion))
            {
                log.Error("Did not find TwinCAT version in TwinCAT project file path");
                Environment.Exit(Constants.RETURN_TWINCAT_VERSION_NOT_FOUND);
            }

            try
            {
                vsInstance = new VisualStudioInstance(@VisualStudioSolutionFilePath, tcVersion, ForceToThisTwinCATVersion);
                bool isTcVersionPinned = XmlUtilities.IsTwinCATProjectPinned(TwinCATProjectFilePath);
                log.Info("Version is pinned: " + isTcVersionPinned);
                vsInstance.Load(isTcVersionPinned);
            }
            catch
            {
                log.Error("Error loading VS DTE. Is the correct version of Visual Studio and TwinCAT installed? Is the TcUnit-Runner running with administrator privileges?");
                CleanUpAndExitApplication(Constants.RETURN_ERROR_LOADING_VISUAL_STUDIO_DTE);
            }

            try
            {
                vsInstance.LoadSolution();
            }
            catch
            {
                log.Error("Error loading the solution. Try to open it manually and make sure it's possible to open and that all dependencies are working");
                CleanUpAndExitApplication(Constants.RETURN_ERROR_LOADING_VISUAL_STUDIO_SOLUTION);
            }

            if (vsInstance.GetVisualStudioVersion() == null)
            {
                log.Error("Did not find Visual Studio version in Visual Studio solution file");
                CleanUpAndExitApplication(Constants.RETURN_ERROR_FINDING_VISUAL_STUDIO_SOLUTION_VERSION);
            }


            AutomationInterface automationInterface = new AutomationInterface(vsInstance.GetProject(), PLCProjectName);
            if (automationInterface.PlcTreeItem.ChildCount <= 0)
            {
                log.Error("No PLC-project exists in TwinCAT project");
                CleanUpAndExitApplication(Constants.RETURN_NO_PLC_PROJECT_IN_TWINCAT_PROJECT);
            }

            /* Build the solution and collect any eventual errors. Make sure to
             * filter out everything that is an error
             */

            vsInstance.CleanSolution();
            vsInstance.BuildSolution();
            System.Threading.Thread.Sleep(100);

            ErrorItems errorsBuild = vsInstance.GetErrorItems();

            int tcBuildWarnings = 0;
            int tcBuildError = 0;
            for (int i = 1; i <= errorsBuild.Count; i++)
            {
                ErrorItem item = errorsBuild.Item(i);
                if ((item.ErrorLevel != vsBuildErrorLevel.vsBuildErrorLevelLow))
                {
                    if (item.ErrorLevel == vsBuildErrorLevel.vsBuildErrorLevelMedium)
                        tcBuildWarnings++;
                    else if (item.ErrorLevel == vsBuildErrorLevel.vsBuildErrorLevelHigh)
                    {
                        tcBuildError++;
                        log.Error("Description: " + item.Description);
                        log.Error("ErrorLevel: " + item.ErrorLevel);
                        log.Error("Filename: " + item.FileName);
                    }
                }
            }

            /* If we don't have any errors, activate the configuration and 
             * start/restart TwinCAT */
            if (tcBuildError.Equals(0))
            {
                log.Info("No Build errors!");
                log.Info("Checking objects of Project:" + PLCProjectName);
                if (automationInterface.Plcproj.CheckAllObjects())
                {
                    log.Info("No Errors while checking all objects!");
                    log.Info("Trying to save as library...");

                    automationInterface.Plcproj.SaveAsLibrary(LibrarySavePath, false);
                    log.Info("Saved as library at " + LibrarySavePath);
                }
                else
                {
                    log.Info("Error(s) while checking all objects!");
                    CleanUpAndExitApplication(Constants.RETURN_CHECKALLOBJECTS_ERROR);
                }
                CleanUpAndExitApplication(Constants.RETURN_SUCCESSFULL);
            }
            else
            {
                log.Error("Build errors in project");
                CleanUpAndExitApplication(Constants.RETURN_BUILD_ERROR);
            }
        }

        static void DisplayHelp(OptionSet p)
        {
            Console.WriteLine("Usage: TcUnit-Runner [OPTIONS]");
            Console.WriteLine("Loads the TcUnit-runner program with the selected visual studio solution and TwinCAT project.");
            Console.WriteLine("Example #1: TcUnit-Runner -v \"C:\\Jenkins\\workspace\\TcProject\\TcProject.sln\"");
            Console.WriteLine("Example #2: TcUnit-Runner -v \"C:\\Jenkins\\workspace\\TcProject\\TcProject.sln\" -t \"UnitTestTask\"");
            Console.WriteLine("Example #3: TcUnit-Runner -v \"C:\\Jenkins\\workspace\\TcProject\\TcProject.sln\" -t \"UnitTestTask\" -a 192.168.4.221.1.1");
            Console.WriteLine("Example #4: TcUnit-Runner -v \"C:\\Jenkins\\workspace\\TcProject\\TcProject.sln\" -w \"3.1.4024.11\"");
            Console.WriteLine("Example #5: TcUnit-Runner -v \"C:\\Jenkins\\workspace\\TcProject\\TcProject.sln\" -u 5");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }

        /// <summary>
        /// Using the Timeout option the user may specify the longest time that the process 
        /// of this application is allowed to run. Sometimes (on low RAM machines), the
        /// DTE build process will hang and the only way to get out of this situation is
        /// to kill this process and any eventual Visual Studio process.
        /// </summary>
        static private void KillProcess(Object source, System.Timers.ElapsedEventArgs e)
        {
            log.Error("Timeout occured, killing process(es) ...");
            CleanUpAndExitApplication(Constants.RETURN_TIMEOUT);
        }

        /// <summary>
        /// Executed if user interrups the program (i.e. CTRL+C)
        /// </summary>
        static void CancelKeyPressHandler(object sender, ConsoleCancelEventArgs args)
        {
            log.Info("Application interrupted by user");
            CleanUpAndExitApplication(Constants.RETURN_SUCCESSFULL);
        }

        /// <summary>
        /// Cleans the system resources (including the VS DTE) and exits the application
        /// </summary>
        private static void CleanUpAndExitApplication(int exitCode)
        {
            try
            {
                vsInstance.Close();
            }
            catch { }

            log.Info("Exiting application...");
            MessageFilter.Revoke();
            Environment.Exit(exitCode);
        }

        /// <summary>
        /// Prints some basic information about the current run of TcUnit-Runner
        /// </summary>
        private static void LogBasicInfo()
        {
            log.Info("TcUnit-Runner build: " + Assembly.GetExecutingAssembly().GetName().Version.ToString());
            log.Info("TcUnit-Runner build date: " + Utilities.GetBuildDate(Assembly.GetExecutingAssembly()).ToShortDateString());
            log.Info("Visual Studio solution path: " + VisualStudioSolutionFilePath);
            log.Info("");
        }
    }
}
