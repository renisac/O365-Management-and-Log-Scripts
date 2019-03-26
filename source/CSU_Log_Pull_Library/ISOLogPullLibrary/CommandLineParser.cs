/******************************
 * ISOLogPullLibrary
 * 
 * Author:      Donald Murchison
 *              Information Security Office 
 *              Sacramento State University
 * Date:        08/10/2017
 * 
 * Description: Library written to help applications interact with the Office 365 Management API to manage logs.
 *              The library will allow users to authenticate to the API, manage subscriptions, and retrieve available content.
 *              ExchangeOnlineLogPull Application is already written to use this library.
 *
 * Notes: CommandLineParser.cs is not being used by current version
 *         
 * *****************************/

 
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ISOLogPullLibrary
{
    // CommandLine Parser class so that application does not need to parse command lines for us
    // Application still needs to handle the format checking
    // Version 2 - handle format checking here
     
    public class CommandLineParser
    {

        // Static Parse method called to parse any command line args
        // Command Line Parser never needs to be instantiated
        public static Dictionary<string, string> Parse(string[] args)
        {
            // Convert input string array into a dictionary which can be passed to the LogPull constructor
            Dictionary < string, string> arguments = new Dictionary<string, string>();
            
            // For each command line argument, argv is an array based of spaces between strings
            foreach (string arg in args)
            {
                // If the current argument contains a special bash character stop adding stuff to the dictionary and return
                if (arg.Contains(">") || arg.Contains("|") || arg.Contains("<") || arg.Contains("&") || arg.Contains("&&"))
                {
                    return arguments;
                }

                // if argument contains help but not output set help key and return
                // we test that it does not conatin output to allow help to be in the output file path 
                if (arg.Contains("help") && !arg.Contains("output"))
                {
                    //If find help retet urn right away
                    // using help=sjust so that we can keep a dictionary format and allow user to just pass "help" at command line
                    arguments["help"] = "set";
                    return arguments;
                }

                else
                {
                    // Check for the other special arguments passed that do not require an "="
                    // Still check to see if this argument if the output path argument
                    // Follows same behaviior as help except does not return right away
                    // This allows user to still set subtype to something other than default
                    if (arg.Contains("list") && !arg.Contains("output"))
                    {
                        arguments["list"] = "set";
                    }
                    else if (arg.Contains("start") && !arg.Contains("starttime") &&!arg.Contains("output"))
                    {
                        arguments["start"] = "set";
                    }
                    else if (arg.Contains("stop") && !arg.Contains("output"))
                    {
                        arguments["stop"] = "set";
                    }
                    // if the argument is not one of the special arguments
                    // split by equal sign and add to dictionary
                    // Don't need to check if key is correct, if it isn't it will be ignored by application
                    else
                    {
                        try
                        {
                            var tmp = arg.Split('=');
                            if (tmp.Count() == 2)
                            {
                                //Make all keys lowercase
                                arguments[tmp[0].ToLower()] = tmp[1];
                            }
                            // If multiple equal signs in one arg
                            // Print error and return null
                            else
                            {
                                Console.WriteLine("Error: Not Proper Format");
                                return null;
                            }

                        }
                        // If no equals sign in arg
                        // Print error and return null
                        catch
                        {
                            Console.WriteLine("Error: Not Proper Format");
                            return null;
                        }
                    }
                    
                }  
            }
            return arguments;
        }
    }
}
