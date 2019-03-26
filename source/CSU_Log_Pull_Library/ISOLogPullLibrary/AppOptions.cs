
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
 * *****************************/


using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Text;
using System.Threading.Tasks;
using Galactic.Configuration;

// Todo: format checking
//       check for trailing slash

namespace ISOLogPullLibrary
{
    // This class handles retrieving the default config values from AppOptions.config file
     
    public class AppOptions
    {

        // Name of config file, will be AppOptions.config in file explorer
        // On first instance of running application this file will be created but application will exit
        // The required fields will be listed in the AppOptions.config but will be empty
        // User will need to open in text editor and enter correct values

        // Config file format is fieldName=fieldValue
        // Comments can be added to the config file by placing a semicolon at beginning of line
        private const string APP_CONFIG = "AppOptions";

        // Instance of Azure, e.g. https://login.microsoftonline.com
        // Make sure not to include trailing slash
        private string AADInstance
        {
            get; set;
        }


        // Name of Azure AD Tenant in which application is registered
        // Do not include trailing slash
        private string Tenant
        {
            get; set;
        }

        // Instance of Azure with Tenant appended, e.g. https://login.microsoftonline.com/mysacstate.onmicrosoft.com
        // Not set in Config, application handles creation of this from AADInstance and Tenant
        public string Authority
        {
            get; private set;
        }

        // Client ID used by the application to uniquely identify itself
        // Find value in Azure 
        public string ClientId
        {
            get; private set;
        }

        
        // Tenant ID used in url of manage.office.com to uniquely identify tenant
        // Find value in Azure
        public string TenantId
        {
            get; private set;
        }

        // Thumb print of the certificate being used to authenticate
        // View Docmentation for guide to retrieving this value 
        public string CertThumbPrint
        {
            get; private set;
        }

        // Resource Id for end point to authenticate to
        // This should always be https://manage.office.com
        // Make sure no rtailing slash
        // Version 2 - allow user to set to https://graph.windows.net to retrieve azure reports
        public string ResourceId
        {
            get; private set;
        }

        // This is the type of logs the user is trying to retrieve
        // Not necessary to be specified in config file
        // If it is not specified the default value, Exchange, will be used
        // Or the argument subtype can be passed on the command line
        // e.g. ExchangeOnlineLogPull.exe subtype=SharePoint
        public string SubscriptionType
        {
            get; private set;
        }

        // Location to store temporary files
        public string TempFolder
        {
            get; private set;
        }

        // debug mode
        public bool DebugMode
        {
            get; private set;
        }

        // try and use the new oauth2 if specified
        public bool UseOauth2
        {
            get; private set;
        }

        // oauth2 authority
        public string Oauth2AuthorityId
        {
            get; private set;
        }

        // oauth2 redirect ui
        public string Oauth2RedirectUi
        {
            get; private set;
        }

        // Deprecated remove during clean up
        /*public int MaxThreads
        {
            get; private set;
        }*/

        // Constructor which takes no arguments
        // reads config file into dictionary and then 
        // calls update
        public AppOptions()
        {
            Dictionary<string, string> appOptionsDictionary = new Dictionary<string, string>();
            appOptionsDictionary = Config_Reader();
            Update(appOptionsDictionary);
        }

        // Used by constructor
        // Was also created so that a GUI application could be used to edit the config file
        // GUI application has been put on hold
        private void Update(Dictionary<string, string> appOptionsDictionary)
        {
            // Get the current values from config file
            // Allows the application to only update fields in Dictionary
            // This way user can specify only one field in Dictionary to update
            Dictionary<string, string> oldConfigDictionary = new Dictionary<string, string>();
            oldConfigDictionary = Config_Reader();
            if (appOptionsDictionary != null)
            {
                // try to update with value from the passed dictionary 
                // If a fields is not set, set to old value from config file
                // If it was not in config set to empty string or a default value

                //Add format checking for all

                // Get current directory
                string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;

                // Get file name and append .config
                string path = currentDirectory + APP_CONFIG + ".config";

                // Create new StreamWriter
                // This method overwrites the whole file every time it is called
                // Allows others to write applications using library which can update the config
                using (StreamWriter wr = File.AppendText(path))
                {
                    try
                    {
                        DebugMode = bool.Parse(appOptionsDictionary["debugmode"].Trim());
                    }
                    catch
                    {
                        try
                        {
                            DebugMode = bool.Parse(oldConfigDictionary["debugmode"].Trim());
                        }
                        catch
                        {
                            DebugMode = false;
                        }
                    }
                    try
                    {
                        UseOauth2 = bool.Parse(appOptionsDictionary["useoauth2"].Trim());
                    }
                    catch
                    {
                        try
                        {
                            UseOauth2 = bool.Parse(oldConfigDictionary["useoauth2"].Trim());
                        }
                        catch
                        {
                            UseOauth2 = false;
                        }
                    }
                    try
                    {
                        // need to check for trailing "/"
                        AADInstance = appOptionsDictionary["aadinstance"];
                    }
                    catch
                    {
                        try
                        {
                            // Check if it was set in old config file
                            AADInstance = oldConfigDictionary["aadinstance"];
                        }
                        catch
                        {
                            //If not found at all add to config with default value
                            AADInstance = "https://login.microsoftonline.com";
                            wr.WriteLine("aadinstance=" + AADInstance);
                        }

                    }

                    try
                    {
                        Tenant = appOptionsDictionary["tenant"];
                    }
                    catch
                    {
                        try
                        {
                            Tenant = oldConfigDictionary["tenant"];
                        }
                        catch
                        {
                            Tenant = "";
                            wr.WriteLine("tenant=" + Tenant);
                        }

                    }
                    try
                    {
                        ClientId = appOptionsDictionary["clientid"];
                    }
                    catch
                    {
                        try
                        {
                            ClientId = oldConfigDictionary["clientid"];
                        }
                        catch
                        {
                            ClientId = "";
                            wr.WriteLine("clientid=" + ClientId);
                        }

                    }
                    try
                    {
                        TenantId = appOptionsDictionary["tenantid"];
                    }
                    catch
                    {
                        try
                        {
                            TenantId = oldConfigDictionary["tenantid"];
                        }
                        catch
                        {
                            TenantId = "";
                            wr.WriteLine("tenantid=" + TenantId);
                        }
                    }
                    try
                    {
                        CertThumbPrint = appOptionsDictionary["certthumbprint"];
                    }
                    catch
                    {
                        try
                        {
                            CertThumbPrint = oldConfigDictionary["certthumbprint"];
                        }
                        catch
                        {
                            CertThumbPrint = "";
                            wr.WriteLine("certthumbprint=" + CertThumbPrint);
                        }
                    }
                    try
                    {
                        ResourceId = appOptionsDictionary["resourceid"];
                    }
                    catch
                    {
                        try
                        {
                            ResourceId = oldConfigDictionary["resourceid"];
                        }
                        catch
                        {
                            ResourceId = "https://manage.office.com";
                            wr.WriteLine("resourceid=" + ResourceId);
                        }
                    }
                    try
                    {
                        Oauth2AuthorityId = appOptionsDictionary["oauth2authorityid"].Trim();
                    }
                    catch
                    {
                        try
                        {
                            Oauth2AuthorityId = oldConfigDictionary["oauth2authorityid"].Trim();
                        }
                        catch
                        {
                            Oauth2AuthorityId = "https://login.windows.net/common/oauth2/authorize";
                        }
                    }
                    try
                    {
                        Oauth2RedirectUi = appOptionsDictionary["oauth2redirectui"].Trim();
                    }
                    catch
                    {
                        try
                        {
                            Oauth2RedirectUi = oldConfigDictionary["oauth2redirectui"].Trim();
                        }
                        catch
                        {
                            Oauth2RedirectUi = "https://login.live.com/oauth20_desktop.srf";
                        }
                    }
                    // These two settings have been removed for now
                    // Tempfolder had not been implemented and subscriptiontype can be set at command line
                    // AppOptions now only contains the configuration information
                    /*try
                    {
                        SubscriptionType = appOptionsDictionary["subscriptiontype"];
                    }
                    catch
                    {
                        try
                        {
                            SubscriptionType = oldConfigDictionary["subscriptiontype"];
                        }
                        catch
                        {
                            SubscriptionType = "Exchange";
                            wr.WriteLine("subscriptiontype=" + SubscriptionType);
                        }
                    }
                    try
                    {
                        TempFolder = appOptionsDictionary["tempfolder"];
                    }
                    catch
                    {
                        try
                        {
                            TempFolder = oldConfigDictionary["tempfolder"];
                        }
                        catch
                        {
                            TempFolder = Path.GetTempPath();
                            wr.WriteLine("tempfolder=" + TempFolder);
                        }
                    }*/
                    // This is the only attribute which will not be found in the config file
                    // Application uses AADInstance and Tenant to create Authority
                    // Does not write to config file
                    try
                    {
                        string tmp = AADInstance + "/{0}";
                        //Authority = String.Format(CultureInfo.InvariantCulture, tmp, Tenant);
                        Authority = String.Format(CultureInfo.InvariantCulture, tmp, TenantId);

                    }
                    catch
                    {
                        Authority = "";
                    }
                    // Deprecated
                    /*try
                    {
                        MaxThreads = Int32.Parse(appOptionsDictionary["maxthreads"]);
                    }
                    catch
                    {
                        try
                        {
                            MaxThreads = Int32.Parse(oldConfigDictionary["maxthreads"]);
                        }
                        catch
                        {
                            MaxThreads = 64;
                            wr.WriteLine("maxthreads=" + MaxThreads);
                        }
                    }*/

                }

            }
            
        }


        // Method used to read the Config File
        private Dictionary<string, string> Config_Reader()
        {
            // Code for this section taken from Will Sykes' Log AI project

            // Get current directory
            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            Dictionary<string, string> appOptionsDictionary = new Dictionary<string, string>();
            //when you parse app options don't forget to handle extra = in the connection string AND comments that start with ;
            try
            {
                //the general app specific paramenters
                ConfigurationItem appConfig = new ConfigurationItem(currentDirectory, APP_CONFIG, false);

                //Create Reader for config file
                using (StringReader sr = new StringReader(appConfig.Value))
                {
                    // while more lines in file
                    while (sr.Peek() != -1)
                    {
                        string line = sr.ReadLine().Trim();

                        //Check if leading semi-colon, ignore if there is
                        if (!(line.StartsWith(";", StringComparison.CurrentCulture)))
                        {
                            //Split on equal sign
                            var tmp = line.Split('=');

                            // If equal sign in value it would have been split 
                            // If so, restore value with equal sign
                            // Otherwise, check if it has already been set
                            if (tmp.Length > 2)
                            {
                                string parts = "";
                                for (int i = 1; i < tmp.Length; i++)
                                {
                                    // Basically a join for the parts with equal signs

                                    //dont add an = to the end of the string
                                    if (!(i == tmp.Length - 1)) 
                                    {
                                        parts += tmp[i] + "=";
                                    }
                                    else
                                    {
                                        parts += tmp[i];
                                    }
                                }
                                try
                                {
                                    // Check if the value had already been specified in the config file
                                    // If not set, set it in Dictionary
                                    if (!(appOptionsDictionary.ContainsKey(tmp[0].ToLower())))
                                    {
                                        appOptionsDictionary.Add(tmp[0].ToLower(), parts.Trim());
                                    }
                                }
                                catch { }
                            }
                            else
                            {
                                // Check if the value had already been specified in the config file
                                // If not set, set it in Dictionary
                                try
                                {
                                    if (!(appOptionsDictionary.ContainsKey(tmp[0].ToLower())))
                                    {
                                        appOptionsDictionary.Add(tmp[0].ToLower(), tmp[1].Trim());
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                    return appOptionsDictionary;
                }
                
            }
            catch (Exception err)
            {
                Debug.WriteLine("Within the processappoptions method. Unable to process app options file {0}, {1}", err.Message, err.InnerException);
                return null;
            }
        } 

        // Unused with current application
        // This was the method our GUI would have called which would then call Update
        public void Config_Editor(string values)
        {
            
            Dictionary<string, string> appOptionsDictionary = new Dictionary<string, string>();

            // Read old config file
            // This way if not all the values are set we set them in dictionary
            // This makes some of the checking in the Update method redundant
            appOptionsDictionary = Config_Reader();

            // the format of string which is passed should be "fieldName1=value1;fieldname2=value2;"

            // Split the different key value pairs
            var fields = values.Split(';');
            foreach(var field in fields)
            {
                var pair = field.Split('=');
                try
                {
                    // Check if key was in old config file
                    // if not add with value, otherwise update the value for the matching key
                    // this does not handle if the same value is specified twice in the update string
                    // if it is the resulting value will be the last one specified

                    // *Note - with the syntax may not need to run if statement might be able to use
                    //         content of else statement whether or not key is already there
                    if (!(appOptionsDictionary.ContainsKey(pair[0].ToLower())))
                    {
                        appOptionsDictionary.Add(pair[0].ToLower(), pair[1]);
                    }
                    else
                    {
                        appOptionsDictionary[pair[0].ToLower()] = pair[1];
                    }
                }
                catch {
                    
                }
            }

            // get file path for AppOptions.config
            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string path = currentDirectory + APP_CONFIG + ".config";

            // Create Writer and write all values to config
            using (StreamWriter wr = new StreamWriter(path)){
                foreach (var item in appOptionsDictionary)
                {
                    wr.WriteLine(item.Key+"="+item.Value);
                }
            }
            // Call update
            // This is where a lot of things get redundant
            // This way we update the application instance with the edited values in the config file 
            Update(appOptionsDictionary);
        }
    }
}
