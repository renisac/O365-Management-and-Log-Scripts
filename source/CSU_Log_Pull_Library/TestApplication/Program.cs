using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ISOLogPullLibrary;


namespace TestApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            LogPull logPull = new ISOLogPullLibrary.LogPull();
            //var subscriptionList = logPull.ListSubscription();
            //foreach (var subscription in subscriptionList)
            //{
            //    Console.WriteLine(subscription);
            //}
            logPull.StartSubscription();
            Environment.Exit(0);
           
        }
    }
}
