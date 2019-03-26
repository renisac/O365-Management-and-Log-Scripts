/******************************
 * ExchangeOnlineLogPull Application
 * 
 * Author:      Donald Murchison
 *              Information Security Office 
 *              Sacramento State University
 * Date:        08/10/2017
 * 
 * Description: Console Application designed to use the ISOLogPullLibrary. The main purpose of this is to parse command line options
 *              and call the appropriate methods in the ISOLogPullLibrary. Outputs a CSV file containing Online Office 365 logs for 
 *              varying services. 
 *               
 * Usage:       ExchangeOnlineLogPull.exe 
 * Help Menu:   ExchangeOnlineLogPull.exe help
 * 
 * *****************************/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Newtonsoft.Json.Linq;
using System.IO;

namespace ExchangeOnlineLogPull
{

    // Command Line Application which can start, stop, and list subscriptions for Office 365 Online
    // Main functionality is to pull logs from Office 365 Onlin
    // Default usage requires no command line arguments and will pull the last 24 hours worth of logs
    // Use the help argument at the command line for full details of program and usage
    class Program
    {
        static void Main(string[] args)
        {
            // Declare LogPull class
            ISOLogPullLibrary.LogPull logPull = null;

            // If no arguments passed use default constructor
            if (args.Count() == 0)
            {
                logPull = new ISOLogPullLibrary.LogPull();
            }
            else
            {
                // If arguments passed on command line, call the Parse method from CommandLineParse class 
                // This method will take the string array and return a dictionary 
                Dictionary<string, string> arguments = new Dictionary<string, string>();
                arguments = ISOLogPullLibrary.CommandLineParser.Parse(args);

                // If Parse returned a null object there was an error with the agruments, details should have been printed to command line
                if (arguments == null)
                {
                    Console.WriteLine("Incorrect format: use argument help");
                    Console.WriteLine();
                    Console.WriteLine("Usage: ExchangeOnlineLogPull.exe starttime=YYYY-MM-DDTHH:MM endtime=YYYY-MM-DDTHH:MM subtype=Exchange");
                    Environment.Exit(0);
                }

                // Else if help argument was specified anywhere on the command line print help and then exit
                else if (arguments.ContainsKey("help"))
                {
                    printHelp();
                    Environment.Exit(1);
                }
                // Else if arguments in proper format run check arguments and then use to instantiate LogPull class
                else
                {
                    // CheckArgs method described in greater detail below
                    // if returns true then instantiate LogPull otherwise printe error and exit
                    if (checkArgs(arguments))
                    {
                        logPull = new ISOLogPullLibrary.LogPull(arguments);
                    }
                    else
                    {
                        Console.WriteLine("Incorrect format: use argument help");
                        Console.WriteLine();
                        Console.WriteLine("Usage: ExchangeOnlineLogPull.exe starttime=YYYY-MM-DDTHH:MM endtime=YYYY-MM-DDTHH:MM subtype=Exchange");
                        Environment.Exit(0);
                    }
                }
                // After instantiating LogPull with dictionary check for special arguments
                // These are list, start, and stop
                // We do this after instantiating LogPull so that subtype can be set for these actions 
                if (arguments.ContainsKey("list"))
                {
                    // if list is set call ListSubscription Method and then print results and exit
                    var subscriptionList = logPull.ListSubscription();
                    if (subscriptionList != null)
                    {
                        foreach (var subscription in subscriptionList)
                        {

                            Console.WriteLine(subscription);
                        }
                    }
                    else
                    {
                        Console.WriteLine(String.Format("No subscriptions found for tenant"));
                    }
                    Environment.Exit(0);
                }
                else if (arguments.ContainsKey("start"))
                {
                    // if start is set call StartSubscription, print response, and exit
                    // currently response is printing as empty
                    int response = logPull.StartSubscription();
                    if (response == 1)
                    {
                        Console.WriteLine("Successfully Started Subscription");
                    }
                    else
                    {
                        Console.WriteLine("Failed To Start Subscription");
                    }
                    Environment.Exit(0);
                }
                else if (arguments.ContainsKey("stop"))
                {
                    // if stop is set call StartSubscription, print response, and exit
                    // currently response is printing as empty
                    int response = logPull.StopSubscription();
                    if(response == 1)
                    {
                        Console.WriteLine("Successfully Stopped Subscription");
                    }else
                    {
                        Console.WriteLine("Failed To Stop Subscription");
                    }
                    
                    Environment.Exit(0);
                }
            }

            // if none of the special arguments are set, try to pull logs
            try
            {
               
                logPull.ThreadedGetLogs();
                
            }
            catch
            {
                Console.WriteLine("Error: Could not obtain logs");
                Console.WriteLine("Check Permissions on Cert and values in appOptions.config");
            }
                
            
        }

        // Method just to print proper usage of application and command line arguments
        private static void printHelp()
        {
            Console.WriteLine("********** Help Menu **********");
            Console.WriteLine();
            Console.WriteLine("Usage: ExchangeOnlineLogPull.exe starttime=YYYY-MM-DD endtime=YYYY-MM-DD subtype=Exchange output=C:\\Users\\ExampleUser\\Desktop\\outputfile");
            Console.WriteLine();
            Console.WriteLine("********** Arguments **********");
            Console.WriteLine();
            Console.WriteLine("StartTime: Optional datetime (UTC) indicating time range of content to return");
            Console.WriteLine("           Must also specify EndTime");
            Console.WriteLine();
            Console.WriteLine("           Acceptable Formats: YYYY-MM-DD");
            Console.WriteLine("                               YYYY-MM-DDTHH:MM");
            Console.WriteLine("                               YYYY-MM-DDTHH:MM:SS");
            Console.WriteLine();
            Console.WriteLine("EndTime:   Optional datetime (UTC) indicating time range of content to return");
            Console.WriteLine("           Must also specify StartTime");
            Console.WriteLine();
            Console.WriteLine("           Acceptable Formats: YYYY-MM-DD");
            Console.WriteLine("                               YYYY-MM-DDTHH:MM");
            Console.WriteLine("                               YYYY-MM-DDTHH:MM:SS");
            Console.WriteLine();
            Console.WriteLine("Relative:  Optional argument to pull logs generated between the current time and X minutes before");
            Console.WriteLine("           The same as setting EndTime to current time and StartTime to (current time - x minutes)");
            Console.WriteLine();
            Console.WriteLine("           Acceptable Values: 1 < X < 1440");
            Console.WriteLine();
            Console.WriteLine("SubType:   Optional argument to specify which SubscriptionType to use");
            Console.WriteLine("           This does not change default SubscriptionType in config file");
            Console.WriteLine("           *Case Sensitive - Acceptable parameters below");
            Console.WriteLine();
            Console.WriteLine("           Acceptable Types:");
            Console.WriteLine("                             AzureActiveDirectory");
            Console.WriteLine("                             Exchange");
            Console.WriteLine("                             SharePoint");
            Console.WriteLine("                             General");
            Console.WriteLine();
            Console.WriteLine("Output:    Optional output argument to specify output file path and name");
            Console.WriteLine("           Application will automatically append .csv extension to file name;");
            Console.WriteLine("           Default output path is current directory;");
            Console.WriteLine("           Default file name is current directory;");
            Console.WriteLine();
            Console.WriteLine("           Example: output=C:\\Users\\ExampleUser\\Desktop\\outputfile");
            Console.WriteLine();
            Console.WriteLine("Start:     Optional argument to start a subscription ");
            Console.WriteLine("           If no subtype is specified it will start the default subscription");
            Console.WriteLine("           specified in appOptions.config");
            Console.WriteLine("           If no subscription type is specified in config default is Exchange;");
            Console.WriteLine();
            Console.WriteLine("           Example: start subtype=SharePoint");
            Console.WriteLine();
            Console.WriteLine("Stop:      Optional argument to stop a subscription ");
            Console.WriteLine("           If no subtype is specified it will stop the default subscription");
            Console.WriteLine("           specified in appOptions.config");
            Console.WriteLine();
            Console.WriteLine("           Example: stop subtype=SharePoint");
            Console.WriteLine();
            Console.WriteLine("List:      Optional argument to list all subscriptions ");
            Console.WriteLine();
            Console.WriteLine("           Example: list subtype=SharePoint");
            Console.WriteLine();
            Console.WriteLine("           The application will only execute one action argument.");
            Console.WriteLine("           Priority - list, start, stop");
            Console.WriteLine();

        }

        // Method to check if all the arguments values are in the proper format
        private static bool checkArgs(Dictionary<string, string> arguments)
        {
            // Checkargs returns true by default only returns false if fails a test
            bool result = true;

            //If starttime set check if endtime is set
            if (arguments.ContainsKey("starttime"))
            {
                
                if (arguments.ContainsKey("endtime"))
                {
                    // If both starttime and endtime are set check if both are in proper date format
                    DateTime tmp;
                    string[] formats = { "yyyy-MM-dd", "yyyy-MM-ddTHH:mm", "yyyy-MM-ddTHH:mm:ss" };
                    if (DateTime.TryParseExact(arguments["starttime"].Trim(), formats,System.Globalization.CultureInfo.InvariantCulture,System.Globalization.DateTimeStyles.None, out tmp))
                    {
                        
                        if (DateTime.TryParseExact(arguments["endtime"], formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out tmp))
                        {
                           
                            //Nothing
                        }else
                        {
                            result = false;
                        }
                    }else
                    {
                        result = false;
                    }
                }
                else
                {
                    Console.WriteLine("Must specify both starttime and endtime");
                    Console.WriteLine();
                    result = false;
                }
            }

            // Same as above but first checks if endtime is set
            // the above checks would be skipped if only endtime was set, these checks will not
            // Version 2 - make a more elegant way of handling
            if (arguments.ContainsKey("endtime"))
            {
                if (arguments.ContainsKey("starttime"))
                {
                    // If both starttime and endtime are set check if both are in proper date format
                    DateTime tmp;
                    string[] formats = { "yyyy-MM-dd", "yyyy-MM-ddTHH:mm", "yyyy-MM-ddTHH:mm:ss" };
                    if (DateTime.TryParseExact(arguments["starttime"].Trim(), formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out tmp))
                    {
                        
                        if (DateTime.TryParseExact(arguments["endtime"], formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out tmp))
                        {

                            //Nothing
                        }
                        else
                        {
                            result = false;
                        }
                    }
                    else
                    {
                        result = false;
                    }
                }
                else
                {
                    Console.WriteLine("Must specify both starttime and endtime");
                    Console.WriteLine();
                    result = false;
                }
            }

            //Relative time argument
            if (arguments.ContainsKey("relative"))
            {
                int x = 0;
                if (Int32.TryParse(arguments["relative"], out x))
                {
                    if(x > 1 && x < 1440)
                    {
                        DateTime endtime;
                        DateTime starttime;
                        endtime = DateTime.UtcNow;
                        starttime = endtime.AddMinutes(x * -1);
                        
                        arguments["endtime"] = endtime.ToString("yyyy-MM-ddTHH:mm");
                        arguments["starttime"] = starttime.ToString("yyyy-MM-ddTHH:mm");
                        Console.WriteLine(arguments["endtime"] + " " + arguments["starttime"]);
                    }
                    else
                    {
                        Console.WriteLine("Relative must be between 1 and 1440");
                        Console.WriteLine();
                        result = false;
                    }
                    
                }
                else
                {
                    Console.WriteLine("Relative must be an integer");
                    Console.WriteLine();
                    result = false;
                }
            }
           
                
                // If subtype was set, make sure it matches one of the acceptable Subscription Types
                // Currently case sensitive
                // Version 2 - make case insensitive
                // Update - ToLower string then set to proper format - makes case insensitive
                if (arguments.ContainsKey("subtype"))
            {
                switch (arguments["subtype"].ToLower())
                {
                    case "azureactivedirectory":
                        arguments["subtype"] = "AzureActiveDirectory";
                        break;
                    case "exchange":
                        arguments["subtype"] = "Exchange";
                        break;
                    case "sharepoint":
                        arguments["subtype"] = "SharePoint";
                        break;
                    case "general":
                        arguments["subtype"] = "General";
                        break;
                    default:
                        result = false;
                        break;
                }
            }

            // If output is set run mulitple checks to make sure file path is OK
            // Since some users might run this as Admin do not overwrite any files
            if (arguments.ContainsKey("output"))
            {
                try
                {
                    // Set a string dir to the output path minus anything after last \ 
                    // this should be the directory where the file will be written
                    string dir = arguments["output"].Substring(0, arguments["output"].LastIndexOf('\\'));

                    // check if output string is not empty
                    if (arguments["output"] == "")
                    {
                        Console.WriteLine("No output file specified");
                        Console.WriteLine();
                        result = false;
                    }

                    // check if the file already exists
                    else if (File.Exists(arguments["output"]))
                    {
                        Console.WriteLine("File already exists: Delete or specify new file path");
                        Console.WriteLine();
                        result = false;
                    }
                    // check if the full path listed is already a directory
                    else if (Directory.Exists(arguments["output"]))
                    {
                        Console.WriteLine("Output path is a directory, please specify a file name at end");
                        Console.WriteLine();
                        result = false;
                    }
                    //check if the path minus the file name actually exists
                    else if (!Directory.Exists(dir))
                    {
                        Console.WriteLine("Directory does not exist");
                        Console.WriteLine();
                        result = false;
                    }
                    else
                    {
                        // Check if user has permissions to write to this directory
                        try
                        {
                            System.Security.AccessControl.DirectorySecurity ds = Directory.GetAccessControl(dir);
                        }
                        catch
                        {
                            Console.WriteLine("Do not have write permissions for folder.");
                            result = false;
                        }

                    }

                }
                catch
                {
                    Console.WriteLine("Error in the output argument");
                    result = false;
                }
            }
            return result;
        }
    }
}
