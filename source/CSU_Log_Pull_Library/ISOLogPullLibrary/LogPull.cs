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
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using CsvHelper;
using Serilog;
using Serilog.Sinks.File;

// Todo
// Review Everything 
// Format checking for attributes
// Especially date

namespace ISOLogPullLibrary
{
    // Log Pull Class
    // This class is all that the application must instantiate;

    // Any Errors are written to stderr 
    // To track Errors specify 2>error.log on command line

    // Each thread prints out "Thread {number} done" once finished
    // This is to double check all threads finished
    // Specify a redirect for stdout and sort at end if want to double check
    
    // The application should instantiate a new LogPull class 
    // and then call the methods StartSubscription, StopSubscription, and LogPull;

    // To specify the arguments StartTime, EndTime, SubscriptionType, or Output Path
    // pass in a dictionary when instantiating 
    public class LogPull
    {
        private Authentication Auth
        {
            get;set;
        }

        private string StartTime
        {
            get;set;
        }

        private string EndTime
        {
            get;set;
        }

        private string SubscriptionType
        {
            get; set;
        }

        private Writer SafeWriter
        {
            get;set;
        }
        
        //Constructor No args set everything to null
        public LogPull()
        {
            Auth = new ISOLogPullLibrary.Authentication();
            StartTime = null;
            EndTime = null;
            SubscriptionType = "Exchange";
            SafeWriter = new Writer(SubscriptionType);
        }
        //Constructor to instantiate based on how dictionary
        public LogPull(Dictionary<string, string> arguments)
        {
            Auth = new ISOLogPullLibrary.Authentication();
            StartTime = null;
            EndTime = null;
            SubscriptionType = "Exchange";
           
            // check for key values in dictionary and set variables
            // Checks for proper format should be done by the application
            if (arguments.ContainsKey("starttime"))
            {
                StartTime = arguments["starttime"];
                Debug.WriteLine("Start: {0}", StartTime);
            }
            if (arguments.ContainsKey("endtime"))
            {
                EndTime = arguments["endtime"];
                Debug.WriteLine("End: {0}", EndTime);
            }

            if (arguments.ContainsKey("subtype"))
            {
                SubscriptionType = arguments["subtype"];
            }
            SafeWriter = new Writer(SubscriptionType);

            if (arguments.ContainsKey("output"))
            {
                // call Writer class method to set file path
                SafeWriter.SetFilePath(arguments["output"]);
            }


                
        }


        // function to get actual logs from Microsoft
        // Checks if subscription for specified type is running
        // if it is not returns error and stops
        // if it is
        //      makes a log request
        //      gets nextpage header from response
        //      response will contain 0-200 URLS to call for actual records
        //      spins off new thread to pull down the actual records
        //      thread handles writing to the file
        //      makes next page request
        //      waits until all threads are finished
        
        public void ThreadedGetLogs()
        {
            ISOLogPullLibrary.AppOptions appOptions = new AppOptions();

            //Removed do to redirects causing error
            //Version 2 - allow use
            //ConsoleSpinner spin = new ConsoleSpinner();


            //This if statement has been deprecated for now
            // subtype in AppOptions has been removed until v2

            // If SubscriptionType was not set in command line 
            // try to use value in config file; Allows user to set own default subtype
            // If not set in config use default Exchange
            /*if (SubscriptionType == null)
            {
                if (appOptions.SubscriptionType == "" || appOptions.SubscriptionType == null)
                {
                    SubscriptionType = "Exchange";
                }
                else
                {
                    SubscriptionType = appOptions.SubscriptionType;
                }

            }*/
            try
            {
                HttpResponseMessage response;
                //HttpResponseMessage logResponse;
                AuthenticationResult authResult = Auth.Result;
                if (authResult == null)
                {
                    Console.WriteLine("Error: Error with Certificate authentication, check if cert properly installed and permissions of user");
                    return;
                }
                string content;
                string request;
                JArray logs = new JArray();
                IEnumerable<string> values = null;

                // check to see if any subscriptions have been started
                // if can't find subscriptions print error and exit

                //*Note - If you are sure subscriptions have been started but are reaching this branch
                //        make sure to check if the certificate was installed properly and that the account
                //        has proper permissions to access this cert.
                //        Service accounts were behaving in a weird way, in which the could not access cert.
                //        If cert was installed on local machine, and not current user, make sure that the 
                //        application is run as admin. 

                // call ListSubscription method to handle actual requests
                var subscriptionList = ListSubscription();
                if (subscriptionList == null || subscriptionList.Count == 0)
                {
                    //using (StreamWriter writer = new StreamWriter("SubscriptionError.log"))
                    //{
                        Console.WriteLine("Error: No Subscriptions found.");
                        int tmp = StartSubscription();
                        if (tmp == 0)
                        {
                            Console.WriteLine("Error: Failed to start subscription for {0}", SubscriptionType);
                            Console.WriteLine("Specify <start subtype={0}> to manually start subscription", SubscriptionType);
                        }
                        else
                        {
                            Console.WriteLine("Successfully Started Subscription");
                        }
                        return;
                    //}
                       
                }
                else
                {
                    // if there is a non empty response, check to see if the specified
                    // subscription type has been started
                    bool found = false;
                    string type = "Audit." + SubscriptionType;
                    foreach (var subscription in subscriptionList)
                    {
                        if (subscription.Value<string>("contentType") == type)
                        {
                            if (subscription.Value<string>("status") == "enabled")
                            {
                                Console.WriteLine(subscription);
                                found = true;
                            }  
                        }
                    }

                    if (found == true)
                    {
                        // If found, make default request or use start and end time if specified
                        if (StartTime != null && EndTime != null)
                        {
                            request = String.Format(CultureInfo.InvariantCulture, "https://manage.office.com/api/v1.0/{0}/activity/feed/subscriptions/content?contentType=Audit.{1}&startTime={2}&endTime={3}", appOptions.TenantId, SubscriptionType, StartTime, EndTime);
                            //Console.WriteLine(request);
                        }
                        else
                        {
                            request = String.Format(CultureInfo.InvariantCulture, "https://manage.office.com/api/v1.0/{0}/activity/feed/subscriptions/content?contentType=Audit.{1}", appOptions.TenantId, SubscriptionType);
                        }

                        
                        // use an array of Manual Reset Events to track if threads are done
                        // WaitAll can only handle 64 events
                        // the WaitAll is really only waiting for the last 64 threads to finish
                        // Likelihood of an earlier thread not finishing before the last 64 is very small
                        // Printing out all threads number to verify that they all finish
                        ManualResetEvent[] doneEvents = new ManualResetEvent[64];
                        int i = 0;
                        int retry = 0;
                        int throttle;
                        do
                        {
                            //spin.Turn();

                            // Get initial page request or nextpage request
                            // Pass the request and authentication result to HttpGet method to make actual Get request
                            throttle = 0;
                            do
                            {
                                response = HttpGet(request, authResult);
                                //response.Headers.TryGetValues("Status", out values);
                                try
                                {
                                    content = response.Content.ReadAsStringAsync().Result;
                                    JObject tmpObj = JObject.Parse(content);
                                    Console.WriteLine(tmpObj["error"]["code"]);
                                    if (tmpObj["error"]["code"].ToString() == "AF429")
                                    {
                                        Thread.Sleep(5000+(1000*throttle));
                                        throttle++;
                                    }
                                    else
                                    {
                                        throttle = 20;
                                    }

                                }catch
                                {
                                    throttle = 20;
                                }
                            }while(throttle < 20);
                                




                            // if the response is not null;
                            // often microsoft returns null response for next page requests
                            // library will retry the request up to 50 times sleeping for 2 seconds
                            // each time
                            
                            if (response != null)
                            {
                                // If next page header is specified in response then store in
                                // variable values 
                                response.Headers.TryGetValues("NextPageUri", out values);

                                // Read content from body of response
                                // Content will be a json array in string format of up to 200 urls
                                // Each url must be called to obtain actual records
                                content = response.Content.ReadAsStringAsync().Result;

                                // Parse string to JArray object type
                                try
                                {
                                    var records = JArray.Parse(content);

                                    // Create new manual reset event which will be passed to thread
                                    var done = new ManualResetEvent(false);
                                    doneEvents[i % 64] = done;

                                    // Create new Log Object type which has the work item method
                                    // Pass the urls as records, authentication result of certificate authentication, 
                                    // the manual reset event, and thread safe Writer, to the contrsuctor class
                                    Log l = new Log(records, authResult, done, SafeWriter);

                                    // Call ThreadPoolCallback method and add to threadpool work item queue
                                    ThreadPool.QueueUserWorkItem(l.ThreadPoolCallback, i);
                                }
                                catch
                                {
                                    Console.Error.WriteLine(content);
                                }
                                
                                
                                // if the NextPageUri header was empty, set request to null to end loop
                                // else set request to the NextPageUri
                                if (values == null)
                                {
                                    request = null;
                                }
                                else
                                {
                                    request = values.First();
                                }
                                //clear values and response, increment i, and then loop
                                values = null;
                                response = null;
                                i++;
                            }
                            // else branch from if(response !=null)
                            else
                            {

                                // Check if retry attempts below 50
                                // If it is increment sleep and the increment retry and loop
                                if (retry <50)
                                {
                                    Thread.Sleep(2000);
                                    //Console.WriteLine("Error: Retrying ...");
                                    retry++;
                                }
                                // else print error statement, current threads will finish and application
                                // will end early, there is no way to proceed because we cannot get a NextPageUri
                                else
                                {
                                    Console.Error.WriteLine("Error: Page Request came back null,");
                                    Console.Error.WriteLine("Error: Not all pages were acquired");
                                    retry = 0;
                                    request = null;
                                }

                            }
                            
                        } while (request != null);

                        // Once all the pages have been received and all the threads needed have been created
                        // application needs to wait for threads to finish, since we are using array need to check 
                        // if the array was filled or not.

                        // if less than 64 threads were created, application must resize array to correct length
                        if (i < 64)
                        {
                            Array.Resize<ManualResetEvent>(ref doneEvents, (i));
                        }

                        // Use wait handle all to wait for the last threads to finish
                        int j = i;
                        do
                        {
                            try
                            {
                                Console.WriteLine("Waiting...");
                                WaitHandle.WaitAll(doneEvents);
                                Console.WriteLine("All threads are complete.");
                                j = 0;

                            }
                            catch
                            {
                                j--;
                                if (i < 64)
                                {
                                    Array.Resize<ManualResetEvent>(ref doneEvents, (i));
                                }
                            }
                        } while (j!=0);
                        
                        SafeWriter.Stitch(i);
                    }

                    // else branch from if(found==true)
                    // reach this branch if subscriptions found but not the subtype specified
                    else
                    {
                        //using (StreamWriter writer = new StreamWriter("SubscriptionError.log"))
                        //{
                            Console.WriteLine("Error: Subscription {0} was not found", SubscriptionType);
                            int tmp = StartSubscription();
                            if (tmp == 0)
                            {
                                Console.WriteLine("Error: Failed to start subscription for {0}", SubscriptionType);
                                Console.WriteLine("Specify <start subtype={0}> to manually start subscription", SubscriptionType);
                            }else
                            {
                                Console.WriteLine("Started Subscription");
                            }
                            return;
                        //}
                            
                    }
                }

            }
            // Catch all branch
            catch (Exception ex)
            {

                Console.Error.WriteLine("{0}:{1}", ex.Message, ex.InnerException);
                Console.Error.WriteLine("{0}", ex.ToString());
                return;
            }

        }


        // Method used to start a subscription for the specified subtype 
        public int StartSubscription()
        {
            // Instaniate a new AppOptions class.
            // This ensures that we grab the most recent config file data.
            ISOLogPullLibrary.AppOptions appOptions = new AppOptions();
            try
            {
                // check to see if suscription type was specified at command line
                // if not check to see if specified in config file
                // otherwise set as default type Exchange
                if (SubscriptionType == null)
                {
                    if (appOptions.SubscriptionType == "" || appOptions.SubscriptionType == null)
                    {
                        SubscriptionType = "Exchange";
                    }
                    else
                    {
                        SubscriptionType = appOptions.SubscriptionType;
                    }

                }

                // Make post request manage.office.comapi to start specified subtype
                // get authentication from Authentication object create at LogPull instaniation   
                string request = String.Format(CultureInfo.InvariantCulture, "https://manage.office.com/api/v1.0/{0}/activity/feed/subscriptions/start?contentType=Audit.{1}", appOptions.TenantId, SubscriptionType);
                AuthenticationResult authResult = Auth.Result;
                if (appOptions.DebugMode)
                {
                    Console.WriteLine(String.Format("Start subscription requested subtype {0}. Request {1}", SubscriptionType, request));
                }
                
                
                // pass request and authentication to HttpPost method to send actual Post request
                HttpResponseMessage response = HttpPost(request, authResult);
                string content = response.Content.ReadAsStringAsync().Result;
                if (appOptions.DebugMode)
                {
                    Console.WriteLine(String.Format("HTTP Response {0}", content));
                }
                //var array = JArray.Parse(content); 
                return 1;
            }
            catch
            {
                
                return 0;
            }

            
        }

        // Method used to start a subscription for the specified subtype
        // Same as start but with stop specified.
        // Could combine the two methods in version 2 and just pass extra argument to specify start or stop 
        public int StopSubscription()
        {
            // Instaniate a new AppOptions class.
            // This ensures that we grab the most recent config file data.
            ISOLogPullLibrary.AppOptions appOptions = new AppOptions();
            try
            {
                if (SubscriptionType == null)
                {
                    if (appOptions.SubscriptionType == "" || appOptions.SubscriptionType == null)
                    {
                        SubscriptionType = "Exchange";
                    }
                    else
                    {
                        SubscriptionType = appOptions.SubscriptionType;
                    }

                }

                // Make post request to manage.office.com api to stop specified subtype
                // get authentication from Authentication object created at LogPull instaniation  
                string request = String.Format(CultureInfo.InvariantCulture, "https://manage.office.com/api/v1.0/{0}/activity/feed/subscriptions/stop?contentType=Audit.{1}", appOptions.TenantId, SubscriptionType);
                AuthenticationResult authResult = Auth.Result;
                HttpResponseMessage response = HttpPost(request, authResult);
                //string content = response.Content.ReadAsStringAsync().Result;
                //var array = JArray.Parse(content);
                return 1;
            }
            catch
            {
                return 0;
            }
        }

        // Method to make api call to list subscriptions and return result
        public JArray ListSubscription()
        {
            // Instaniate a new AppOptions class.
            // This ensures that we grab the most recent config file data.
            ISOLogPullLibrary.AppOptions appOptions = new AppOptions();
            try
            {
                // Make get request to manage.office.com api to list all subscriptions
                // get authentication from Authentication object create at LogPull instaniation 
                string request = String.Format(CultureInfo.InvariantCulture, "https://manage.office.com/api/v1.0/{0}/activity/feed/subscriptions/list", appOptions.TenantId);
                if (appOptions.DebugMode)
                {
                    Console.WriteLine(String.Format("API request {0}", request));
                }
                AuthenticationResult authResult = Auth.Result;
                HttpResponseMessage response = HttpGet(request, authResult);

                string content = null;
                try
                {
                    content = response.Content.ReadAsStringAsync().Result;
                }
                catch
                {
                    Console.WriteLine("Failed to read HTTP return. HTTP code {0}", response.StatusCode);
                }
                if (appOptions.DebugMode)
                {
                    Console.WriteLine(String.Format("HTTP response {0}", content));
                }
                var array = JArray.Parse(content);
                return array;
            }
            catch
            {
                return null;
            }
        }

        // Method which handles the actual http get request  
        private static HttpResponseMessage HttpGet(string request, AuthenticationResult result)
        {
            // Microsoft often sends the url in the page request but when the url is called to quickly it returns null
            // Here we attempt to call the url up to 100 times sleeping for 1 second each time
            // If after the retries, print error
            // This will not stop the program, each error statement of this type means roughly 3-5 records are missing
            // This should happen very infrequently, increase retries or sleep interval if happening often

            //*Note - this loop actually makes the retries for NextPageUri 50*100
            //        An HttpGet error message might also mean a NextPageUri could not be called
            //         
            int i = 0;
            while (i < 100)
            {
                try
                {
                    // Create HttpClient Object
                    // Version 2 - create one Client object at time of instantiation with base url
                    HttpClient httpClient = new HttpClient();

                    // Create new request message with action Get and target equal to request
                    HttpRequestMessage httpRequest = new HttpRequestMessage(HttpMethod.Get, request);

                    //Set Authorization Header Bearer value from authentication result
                    httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);

                    // Get response through async call and return it
                    HttpResponseMessage httpResponse = httpClient.SendAsync(httpRequest).Result;
                    return httpResponse;
                }
                // When the url returns null it throws an exception so we use catch
                // to sleep and increment retry
                catch 
                {
                    if (i < 100)
                    {
                        Thread.Sleep(1000);
                        //Console.WriteLine("Retry #{0}", i);
                        i++;
                    }
                    
                }
            }
            // If all retries fail print error
            // Version 2 - add way to determine if this was a NextPageUri request or a record url request
            //Console.Error.WriteLine("HttpGet Error: Missing record");
            return null;

        }

        // Method which handles the actual http post request
        // Same method as HttpGet but creates request wiht action post
        // Version 2 - Combine methods and specify action as argument
        private HttpResponseMessage HttpPost(string request, AuthenticationResult result)
        {
            // Retries much less in this method because post is only used for start and stop subscription
            // Not likely to have an error
            int i = 0;
            while (i <20)
            {
                try
                {
                    HttpClient httpClient = new HttpClient();
                    HttpRequestMessage httpRequest = new HttpRequestMessage(HttpMethod.Post, request);
                    //Set Authorization Header Bearer value
                    httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                    HttpResponseMessage httpResponse = httpClient.SendAsync(httpRequest).Result;
                    return httpResponse;
                }
                catch
                {
                    if (i < 20)
                    {
                        Thread.Sleep(2000);
                        //Console.WriteLine("Retry #{0}", i);
                        i++;
                    }

                }
            }
            Console.Error.WriteLine("HttpPost Error");
            return null;

        }



        // Class used to pass information to the threads
        private class Log
        {
            // JArray which holds the urls which need to be called to get records
            private JArray Records
            {
                get;set;
            }
            // ManualReset Event to set at end to signal done
            private ManualResetEvent DoneEvent
            {
                get;set;
            }
            // Authentication Result from cert authentication during LogPull instantiation
            private AuthenticationResult AuthResult
            {
                get;set;
            }
            // Thread Safe Writer used to write logs directly from thread so we don't waste memory storing records and then writing
            private Writer SafeWriter
            {
                get;set;
            }

            // Constructor
            public Log(JArray recs, AuthenticationResult result, ManualResetEvent done, Writer writer)
            {
                Records = recs;
                AuthResult = result;
                DoneEvent = done;
                SafeWriter = writer;
                
            }

            // Method called and passed to user work item queue
            public void ThreadPoolCallback(Object threadContext)
            {

                HttpResponseMessage logResponse;
                string request;
                string content = null;
                
                // foreach url in the JArray make an http request
                foreach (var rec in Records)
                {
                    try
                    {
                        // pull actual url from object in JArray
                        request = String.Format(CultureInfo.InvariantCulture, rec.Value<string>("contentUri"));

                        // pass request to HttpGet method
                        logResponse = HttpGet(request, AuthResult);

                        // if response is not null read body as string
                        // then write to file specified using the one instance of Writer from the LogPull class
                        if (logResponse != null)
                        {
                            content = logResponse.Content.ReadAsStringAsync().Result;

                            if(content != null)
                            {
                                SafeWriter.WriteToFile(JArray.Parse(content), threadContext.ToString() );
                            }
                            else
                            {
                                Console.WriteLine("Error: Null response in thread");
                            }

                        }        
                    }
                    catch
                    {
                        // Just to not crash
                        //Debug.WriteLine("Error when making request to content uri");
                    }

                    
  

                }
                // Write that thread has finished and then set ManualReset Event
                // We write that it has finished to double check all threads finished later on
                Console.WriteLine("Thread {0} Done", threadContext);
                DoneEvent.Set();
            }
        }

        // Class for a thread safe writer
        // Should be instantiated once for the LogPull class
        // And then passed that reference should be passed to the Log class
        // This way we can make sure threads are not writing at the same time
        public class Writer
        {
            // only one attribute this is the file path to output file
            // Current default path is year-month-day-hour-minute-subtype.log.csv
            // The date is taken from time ran not the logs it is pulling
            // and it is written to current directory
            // Version 2 - allow this default path to be set in config file
            // come up with better naming convention 
            public string Filepath
            {
                get; set;
            }
            public string Temppath
            {
                get; set;
            }
            public int FileNum
            {
                get;set;
            }
            // Locker to make this thread safe
            private static object locker = new object();

            // Constructor which takes the set subtype 
            // ****Bug - look in to fixing
            //           if default subtype in config file changed need to change default subtype passed to writer
            //           Currently, it will write wrong subtype if using config value different than exchange
            public Writer(string subtype)
            {
                //Setting Default File Path
                string date = DateTime.Now.ToString("yyyy-MM-dd-HH-mm");
                string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string path = currentDirectory + date + "-"+subtype+".log";
                Filepath = path;
                FileNum = 1;
            }

            // Method used to change file path if output was set as command line parameter
            // One workaround for bug above would be to set file path when checking config file
            // Will probably use this for now 
            // Version 2 - use more elegant solution

            // Go back through make sure this is not security risk
            public void SetFilePath(string output)
            {
                // Run through error checking
                // Application will not overwrite any existing files or create any directories
                if (output == "")
                {
                    Console.WriteLine("Error: No output file specified: Using Default File Path");
                    Console.WriteLine();
                }
                else if (File.Exists(output))
                {
                    Console.WriteLine("Error: File already exists: Using Default File Path");
                    Console.WriteLine();
                }
                else if (Directory.Exists(output))
                {
                    Console.WriteLine("Error: Output path is a directory, please specify a file name at end");
                    Console.WriteLine();
                }
                else
                {
                    string dir = output.Substring(0, output.LastIndexOf('\\'));
                    if (!Directory.Exists(dir))
                    {
                        Console.WriteLine("Error: Directory does not exist: Using Default File Path");
                        Console.WriteLine();
                    }
                    else
                    {
                        Filepath = output;
                    }
                }
            }
           
            //Method to actually write the logs to a file
            public void WriteToFile(JArray logs, string threadContext)
            {
               
                string path = Filepath + threadContext + ".csv";
                using (StreamWriter writer2 = new StreamWriter(path, true))
                using (CsvWriter csv = new CsvWriter(writer2))
                {
                    //using tab delimited since some UserAgent strings contain commas
                    csv.Configuration.Delimiter = "\t";

                    // To convert to csv had to pull each JObject out individually
                    // and then oull each property out of JObject and write as field in csv
                    // then call NextRecord to let csvWriter know to move to next line
                    foreach (JObject log in logs.Children<JObject>())
                    {
                        foreach (JProperty field in log.Properties())
                        {
                            csv.WriteField(field.Value.ToString());
                            //string.Join("\t", new string[] { })
                        }
                        csv.NextRecord();
                    }
                    
                }

            }

            public void Stitch(int threads)
            {
                int i;
                for (i = 0; i < threads; i++)
                {
                    try
                    {
                        string path = Filepath + i.ToString() + ".csv";
                        using (StreamWriter writer = new StreamWriter(Filepath + ".csv", true))
                        using (StreamReader reader = new StreamReader(path))
                        {
                            if (File.Exists(path))
                            {
                                while (reader.Peek() >= 0)
                                {
                                    writer.WriteLine(reader.ReadLine());
                                }
                            }
                        }
                        if (File.Exists(path))
                        {
                            File.Delete(path);
                        }
                    }catch
                    {
                        Console.WriteLine("Error: "+ i.ToString() + " Thread Output not Found");
                    }
                    
                }
            }
        }

    }

}
